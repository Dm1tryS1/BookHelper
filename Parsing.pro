QT += core gui multimedia axcontainer gui-private

TARGET = BookHelp
TEMPLATE = app

SOURCES += \
    main.cpp \
    mainwindow.cpp \
    metadatawidget.cpp \
    renamewidget.cpp

HEADERS += \
    mainwindow.h \
    metadatawidget.h \
    renamewidget.h \

FORMS += \
    mainwindow.ui \
    metadatawidget.ui \
    renamewidget.ui

RESOURCES += \
    img.qrc

RC_FILE = icon.rc
