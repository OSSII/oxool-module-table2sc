/* -*- js-indent-level: 8 -*- */
/*

*/
/* global Admin $ SERVICE_ROOT */
Admin.SocketBroker({

	_module: null,

	// 完整的 API 位址
	_fullServiceURI: "",

	onSocketOpen: function() {
		this.socket.send('getModuleInfo'); // 取得本模組資訊
	},

	onSocketClose: function() {
		console.debug('on socket close!');
	},

	onSocketMessage: function(e) {
		var textMsg = e.data;
		if (typeof e.data !== 'string') {
			textMsg = '';
		}

		// 模組資訊
		if (textMsg.startsWith('moduleInfo '))  {
			var jsonIdx = textMsg.indexOf('{');
			if (jsonIdx > 0) {
				this._module = JSON.parse(textMsg.substring(jsonIdx));
				// 紀錄完整的 API 位址
				this._fullServiceURI = window.location.origin + SERVICE_ROOT + this._module.serviceURI;
				this.initializeComponent();
				this.initializeLogginTable();
				this.socket.send('refreshLog');
			}
		// 日誌內容
		} else if (textMsg.startsWith('logData ')) {
			var jsonIdx = textMsg.indexOf('[');
			if (jsonIdx > 0) {
				var dataArray = JSON.parse(textMsg.substring(jsonIdx));
				this._loggingTable.clear();
				if (dataArray.length > 0) {
					this._loggingTable.rows.add(dataArray).draw();
				}
			}
		} else {
			console.debug('warning! received an unknown message:"' + textMsg + '"');
		}
	},

	initializeComponent: function() {
		$('#tbl2scURI').text(this._fullServiceURI);

		document.getElementById('demoForm').setAttribute('action', this._fullServiceURI);
		document.getElementById('demoForm').removeAttribute('id');
		// 把表單範例 HTML 轉成普通文字
		var formExample = document.getElementById('formExample');
		var innerHTML = formExample.innerHTML;
		formExample.innerHTML = "";
		formExample.innerText = innerHTML;
		formExample.classList.remove('d-none'); // 顯示

		var $form = $('#convertDemo');
		$form.attr('action', this._fullServiceURI);
	},

	initializeLogginTable: function() {
		this._loggingTable = $("#logging_table").DataTable({
			order: [[1, 'desc']],
			columns: [
				{data: 'status'},
				{data: 'timestamp'},
				{data: 'source_ip'},
				{data: 'cost'},
				{data: 'title'},
				{data: 'msg'},
			],
			columnDefs: [
				{
					targets: 0, // status
					render: function(data, type, row) {
						return data ? '<span class="text-success">' + _('Success') + '</span>' :
									  '<span class="text-danger">' + _('Fail') + '</span>'
					},
				},
				{
					targets: 1, // timestamp
					render: function(data, type, row) {
						return new Date(Date.parse(data)).toLocaleString();
					},
				},
				{	// 標題和訊息不需排序功能
					targets: [4, 5],
					orderable: false
				}
			],
			language: {
				url: window.location.origin + SERVICE_ROOT +
					this._module.adminServiceURI + 'js/l10n/' + String.locale + '.json'
			}
		});

		$('#refreshLog').click(function() {
			this.socket.send('refreshLog');
		}.bind(this));
	},

});