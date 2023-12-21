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

//---
#include "xlsxdocument.h"
#include "xlsxchartsheet.h"
#include "xlsxcellrange.h"
#include "xlsxchart.h"
#include "xlsxrichstring.h"
#include "xlsxworkbook.h"
#include "qstring_cmp.h"

#include "json_resolve.h"
//---------------------
typedef enum _COLUMN_HEAD_INDEX {
    Model_Name_B    = 1,
    Factory_B       = 2,
    Description_B   = 3,
    Point_B         = 4,
    Quantity_B      = 5,
    Model_Name_A    = 6,
    Factory_A       = 7,
    Description_A   = 8,
    Point_A         = 9,
    Quantity_A      = 10,
    Change_type     = 11
}COLUMN_HEAD_INDEX;
//列宽
typedef enum _COLUMN_With {
    Model_Name_With = 30,
    Factory_With = 17,
    Description_With = 20,
    Point_With = 90,
    Quantity_With = 9,
    Change_type_With = 16
}COLUMN_With;
//Excel中各项列号
typedef enum _Excel_Column_INDEX
{
    Quantity_Column = 3,
    Point_Column = 4,
    MPN_Column = 5,
    Factory_Column = 6,
    MPN1_Column = 7,
    Factory1_Column = 8,
    Column_OFFSET =2
}Excel_Column_INDEX;
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
    QString Read_colum(const QString Name);
    QString Read_colum_All(const QString File_Name,int start_row,int column);
    QStringList Read_colum_List(const QString File_Name,int start_row,int column);
    QString Read_cell(const QString Name,int row, int column);
    int Get_Row(const QString File_Name,const QString str,int column);

private slots:
    void on_pushButton_open_clicked();
    void on_pushButton_open_old_clicked();
    void on_pushButton_open_cmp_clicked();
    void on_pushButton_tst_clicked();

private:
    Ui::MainWindow *ui;
    QString File_Name_New;
    QString File_Name_Old;
    void Excel_SetTitle(QXlsx::Document *pDocument);
    Qstring_cmp *str_cmp;
    QTableWidgetItem *cell_Item;
    Json_resolve *json;
    void Excel_update();
    QXlsx::Document *Write_xlsx;
    QString Write_xlsx_name;
};
#endif // MAINWINDOW_H
