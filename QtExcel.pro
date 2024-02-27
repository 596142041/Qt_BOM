QT       += core gui

greaterThan(QT_MAJOR_VERSION, 4): QT += widgets

CONFIG += c++17
TARGET = Qt_BOM
DEFINES += QT_DEPRECATED_WARNINGS
DEFINES += QT_MESSAGELOGCONTEXT #定义编译选项。QT_DEPRECATED_WARNINGS表示当Qt的某些功能被标记为过时的，那么编译器会发出警告
# You can make your code fail to compile if it uses deprecated APIs.
# In order to do so, uncomment the following line.
#DEFINES += QT_DISABLE_DEPRECATED_BEFORE=0x060000    # disables all the APIs deprecated before Qt 6.0.0
#使用的老的Excel库
#include(xlsx/qtxlsx.pri)

#新的Excel库文件
# QXlsx code for Application Qt project
QXLSX_PARENTPATH=./QXlsx         # current QXlsx path is . (. means curret directory)
QXLSX_HEADERPATH=./QXlsx/header/  # current QXlsx header path is ./header/
QXLSX_SOURCEPATH=./QXlsx/source/  # current QXlsx source path is ./source/
include(./QXlsx/QXlsx.pri)


SOURCES += \
    LogHandler.cpp \
    json_resolve.cpp \
    main.cpp \
    mainwindow.cpp \
    qstring_cmp.cpp

HEADERS += \
    LogHandler.h \
    json_resolve.h \
    mainwindow.h \
    qstring_cmp.h

FORMS += \
    mainwindow.ui
# Default rules for deployment.
RC_ICONS = check2.ico
qnx: target.path = /tmp/$${TARGET}/bin
else: unix:!android: target.path = /opt/$${TARGET}/bin
!isEmpty(target.path): INSTALLS += target
