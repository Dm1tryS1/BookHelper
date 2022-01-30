QT     += core gui multimedia axcontainer

TARGET = Parsing
TEMPLATE = app

SOURCES += \
    main.cpp \
    mainwindow.cpp \
    metadatawidget.cpp \
    renamewidget.cpp

HEADERS += \
    mainwindow.h \
    metadatawidget.h \
    renamewidget.h

FORMS += \
    mainwindow.ui \
    metadatawidget.ui \
    renamewidget.ui

RESOURCES += \
    img.qrc

RC_FILE = icon.rc