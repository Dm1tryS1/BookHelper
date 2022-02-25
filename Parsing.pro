QT += core gui multimedia axcontainer gui-private

TARGET = BookHelp
TEMPLATE = app

SOURCES += \
    sources/main.cpp \
    sources/mainwindow.cpp \
    sources/metadatawidget.cpp \
    sources/renamewidget.cpp \

HEADERS += \
    headers/mainwindow.h \
    headers/metadatawidget.h \
    headers/renamewidget.h \

FORMS += \
    forms/mainwindow.ui \
    forms/metadatawidget.ui \
    forms/renamewidget.ui

RESOURCES += \
    rc/img.qrc

RC_FILE = rc/icon.rc
