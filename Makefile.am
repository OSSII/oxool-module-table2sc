
MODULE_NAME = Table2SC
MODULE_DIR = @LOOLWSD_MODULE_DATA_DIR@/${MODULE_NAME}

include_paths = -I$(abs_top_builddir)/src/include \
		-I$(abs_top_builddir)/common \
		-I$(abs_top_builddir)/net \
		-I$(abs_top_builddir)/wsd

EXTRA_DIST = admin/*

install-data-hook:
	mkdir -p $(DESTDIR)$(MODULE_DIR)
	cp -aR admin $(DESTDIR)$(MODULE_DIR)/

uninstall-hook:
	rm -fr $(DESTDIR)$(MODULE_DIR)

configdir = @LOOLWSD_MODULE_CONFIG_DIR@
dist_config_DATA = Table2SC.xml

# Module install path
moduledir = @LOOLWSD_MODULE_DIR@
module_LTLIBRARIES = Table2SC.la

Table2SC_la_CPPFLAGS = -pthread ${include_paths}

Table2SC_la_LDFLAGS = -avoid-version -module -lPocoDataSQLite

Table2SC_la_SOURCES = src/Table2SC.cpp

