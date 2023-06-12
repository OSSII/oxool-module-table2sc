
#include <OxOOL/Module/Base.h>
#include <OxOOL/HttpHelper.h>
#include <OxOOL/ModuleManager.h>
#include <OxOOL/ConvertBroker.h>

#include <common/Log.hpp>

#include <Poco/TemporaryFile.h>
#include <Poco/FileStream.h>
#include <Poco/String.h>
#include <Poco/MemoryStream.h>
#include <Poco/Net/HTTPRequest.h>
#include <Poco/Net/HTTPResponse.h>
#include <Poco/Net/HTMLForm.h>

#include <Poco/Data/SessionPool.h>
#include <Poco/Data/Session.h>
#include <Poco/Data/RecordSet.h>
#include <Poco/Data/SQLite/Connector.h>

#include <Poco/JSON/Object.h>

#include <Poco/Timestamp.h>

using namespace Poco::Data::Keywords;

class Table2SC : public OxOOL::Module::Base
{
public:
    Table2SC()
    {
        // 註冊 SQLite 連結
        Poco::Data::SQLite::Connector::registerConnector();
    }

    ~Table2SC()
    {
        Poco::Data::SQLite::Connector::unregisterConnector();
    }

    void initialize() override
    {
        // 初始化 sqlite 資料庫
        /// TODO: 只是日誌而已，不需動用資料庫紀錄，未來將改成 log 檔
        auto session = getDataSession();
        session << "CREATE TABLE IF NOT EXISTS logging ("
                << "rec_id      INTEGER PRIMARY KEY AUTOINCREMENT,"
                << "status      TINYINT NOT NULL,"
                << "source_ip   VARCHAR(32) NOT NULL,"
                << "title       TEXT NOT NULL,"
                << "cost        VARCHAR(32) NOT NULL,"
                << "timestamp   TIMESTAMP DEFAULT CURRENT_TIMESTAMP,"
                << "msg         TEXT NOT NULL)", now;

        // 刪除超過一年的舊紀錄
        session << "DELETE FROM logging WHERE (strftime('%s', 'now') "
                << "- strftime('%s', timestamp)) > 86400 * 365", now;

    }

    void handleRequest(const Poco::Net::HTTPRequest& request,
                       const std::shared_ptr<StreamSocket>& socket) override
    {
        // 紀錄開始時間點
        const std::chrono::steady_clock::time_point startTime = std::chrono::steady_clock::now();

        // 紀錄來源 IP
        const std::string sourceIP = socket->isLocal() ? "127.0.0.1" : socket->clientAddress();

        LOG_INF(logTitle() << "Received request with method '"
                           << request.getMethod() << "' from "
                           << socket->clientAddress());

        // 只是 HEAD 的話，回應 200 OK 的 http header
        if (OxOOL::HttpHelper::isHEAD(request))
        {
            OxOOL::HttpHelper::sendResponseAndShutdown(socket);
            return;
        }
        // 不是 POST 的話，回應
        else if (!OxOOL::HttpHelper::isPOST(request))
        {
            addRecord(false, socket, "(none)", startTime, "No HTTP POST method is used.");
            OxOOL::HttpHelper::sendErrorAndShutdown(
                Poco::Net::HTTPResponse::HTTP_METHOD_NOT_ALLOWED, socket);
            return;
        }

        // 讀取 HTTML Form.
        Poco::MemoryInputStream message(&socket->getInBuffer()[0],
                                        socket->getInBuffer().size());
        const Poco::Net::HTMLForm form(request, message);

        const std::string& title = form.has("title") ? form.get("title") : "noname";
        const std::string& content = form.has("content") ? form.get("content") : "";
        const std::string& toFormat = form.has("format") ? form.get("format") : "ods";

        std::string htmlTemplate = MULTILINE_STRING(
<!doctype html>
<html>
    <head>
        <meta charset="utf-8">
        <title>%TITLE%</title>
    </head>
    <body>
        %CONTENT%
    </body>
</html>
        );

        // 替換內容
        Poco::replaceInPlace(htmlTemplate, std::string("%TITLE%"), title);
        Poco::replaceInPlace(htmlTemplate, std::string("%CONTENT%"), content);

        // 製作暫存路徑
        const Poco::Path tmpPath = Poco::Path::forDirectory(Poco::TemporaryFile::tempName());
        Poco::File(tmpPath).createDirectories();
        chmod(tmpPath.toString().c_str(), S_IXUSR | S_IWUSR | S_IRUSR);

        // 製作暫存檔
        const std::string tmpFile = tmpPath.toString() + "/" + title + ".xls";
        LOG_INF(logTitle() << "Create temporary file: " << tmpFile);
        Poco::FileOutputStream outputStream(tmpFile, std::ios::binary|std::ios::trunc);
        outputStream << htmlTemplate;
        outputStream.close();

        // 取得轉檔用的 Broker
        auto docBroker = OxOOL::ConvertBroker::create(tmpFile, toFormat);

        // 檔案載入完畢後，觸發這理，執行後續處理
        docBroker->loadedCallback([=]() {
            // 全選
            docBroker->sendMessageToKit("uno .uno:SelectAll");
            // 最佳化欄位寬度，額外增加 0.2 公分
            docBroker->sendMessageToKit("uno .uno:SetOptimalColumnWidth?aExtraWidth:short=200");
            // 另存檔案，並傳給 client
            docBroker->saveAsDocument();
            addRecord(true, socket, title, startTime, "Convert to '" + toFormat + "' succeeded.");
        });

        // 非唯讀模式載入檔案
        LOG_INF(logTitle() << "Load file: " << tmpFile);
        if (!docBroker->loadDocument(socket))
        {
            addRecord(false, socket, title, startTime,
                "Failed to create Client Session on docKey ["
                + docBroker->getDocKey() + "].");
        }
    }

    //
    std::string handleAdminMessage(const StringVector& tokens) override
    {
        // 傳回最新的紀錄
        if (tokens.equals(0, "refreshLog"))
        {
            std::string recId = tokens[1];
            auto session = getDataSession();

            // SQL query.
            Poco::Data::Statement select(session);
            select << "SELECT * FROM logging", now;
            Poco::Data::RecordSet rs(select);

            std::size_t cols = rs.columnCount(); // 取欄位數

            // 遍歷所有資料列
            std::string result("logData [");
            for (auto row : rs)
            {
                // 轉為 JSON 物件
                Poco::JSON::Object json;
                for (std::size_t col = 0; col < cols ; col++)
                    json.set(rs.columnName(col), row.get(col));

                std::ostringstream oss;
                json.stringify(oss);
                // 轉為 json 字串
                result.append(oss.str()).append(",");
            }
            // 去掉最後一的 ',' 號
            if (rs.rowCount() > 0)
                result.pop_back();

            result.append("]");

            return result;

        }

        return "";
    }

private:
    /// @brief 取得可用的 data session
    /// @return Poco::Data::Session
    Poco::Data::Session getDataSession()
    {
        static std::string dbName = getDocumentRoot() + "/log.db";
        static Poco::Data::SessionPool sessionPool("SQLite", dbName);
        return sessionPool.get();
    }

    void addRecord(int success,
                   const std::shared_ptr<StreamSocket>& socket,
                   std::string title,
                   const std::chrono::steady_clock::time_point startTime,
                   std::string msg)
    {
        // 現在時間
        const std::chrono::steady_clock::time_point endTime = std::chrono::steady_clock::now();
        // 時間差
        const auto timeSinceStartMs =
            std::chrono::duration_cast<std::chrono::milliseconds>(endTime - startTime).count();
        std::string cost = std::to_string(timeSinceStartMs) + " ms.";
        // 來源 IP
        std::string sourceIP = socket->clientAddress();

        auto session = getDataSession();
        session << "INSERT INTO logging (status, source_ip, title, cost, msg) "
                << "VALUES(?, ?, ?, ?, ?)",
                use(success), use(sourceIP), use(title), use(cost), use(msg), now;
    }
};

OXOOL_MODULE_EXPORT(Table2SC);
