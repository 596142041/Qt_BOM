#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QtGui>
#include <qdebug.h>
#include <QMainWindow>
#include <QFileDialog>
#include <QMessageBox>
#include <QProgressDialog>
#include <QModelIndex>
#include <QTableWidgetItem>
#include <QSettings>
#include <QDir>
#include <QStringList>
#include <QCloseEvent>
#include <QAction>
#include <iostream>
#include <string.h>
#include <string>
#include <map>
#include <vector>
#include <QFile>
#include <QTableWidgetItem>
#include <QDateTime>
#include <QFileInfo>
#include <QSize>

//Excel库
#include "xlsxdocument.h"
#include "xlsxcellrange.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"

//-----------------------
#include "qstring_cmp.h"
#include "json_resolve.h" //json文件解析
#include "LogHandler.h"
#include "rm_dup.h"
//---------------------
#include "qaesencryption.h" //AES加密测试

//---------------------------
QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();
    QStringList Read_colum(QXlsx::Document *pDocument,int start_row,int column);
    QStringList Read_colum_List(const QString File_Name,int start_row,int column);
private slots:
    void on_pushButton_open_clicked();
    void on_pushButton_open_old_clicked();
    void on_pushButton_open_cmp_clicked();
    void on_pushButton_tst_clicked();

    void on_actionrm_dup_triggered();
private:
    Ui::MainWindow *ui;
    void Excel_update();
    QSize Wind_Info();
    QString File_Name_New;
    QString File_Name_Old;
    QString Write_xlsx_name;
    Qstring_cmp *str_cmp;
    QTableWidgetItem *cell_Item;
    Json_resolve *json;
    bool tst_btn_enable;
    bool log_enable;
    bool default_open;
    int write_row;
    int read_start_row;
    QXlsx::Document *Write_xlsx;
    QXlsx::Document *Read_New_BOM;
    QXlsx::Document *Read_Old_BOM;
    //-------------
    QByteArray key16;
    QByteArray key24;
    QByteArray key32;
    QByteArray iv;
    QByteArray in;
    QByteArray outECB128;
    QByteArray outECB192;
    QByteArray outECB256;
    QByteArray inCBC128;
    QByteArray outCBC128;
    QByteArray outOFB128;
};
#endif // MAINWINDOW_H
