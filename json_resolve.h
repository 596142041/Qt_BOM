#ifndef JSON_RESOLVE_H
#define JSON_RESOLVE_H
#define CONFIG_NAME "config.json" //配置文件名
//json各个组名字
#define CONFIG_BOM_Column_Index "BOM_Column_Index"
#define CONFIG_Write_Column_Width "Write_Column_Width"
#define CONFIG_Write_Column_Index "Write_Column_Index"
#include <QObject>
#include <QtGui>
#include <qdebug.h>
#include <QFileDialog>
#include <QMessageBox>
#include <QProgressDialog>
#include <QModelIndex>
#include <QSettings>
#include <QStringList>
#include <QAction>
#include <QDateTime>

#include <QCoreApplication>
#include <iostream>
#include <QString>
#include <QTextStream>
#include <QFile>
#include <QDir>
#include <QFileInfo>

#include <QJsonDocument>
#include <QJsonParseError>
#include <QJsonObject>
#include <QJsonArray>
#include <QJsonValue>
#include <QJsonValueRef>
typedef enum _COLUMN_HEAD_INDEX
{
    Change_date     = 1,
    Indx_cnt        = 2,//待修改
    Model_Name_B    = 3,
    Factory_B       = 4,
    Description_B   = 5,
    Point_B         = 6,
    Quantity_B      = 7,
    Model_Name_A    = 8,
    Factory_A       = 9,
    Description_A   = 10,
    Point_A         = 11,
    Quantity_A      = 12,
    Change_type     = 13
}COLUMN_HEAD_INDEX;  //写入表格中每个列
//列宽
typedef enum _BMP_COLUMN_Width
{
    Quantity_With = 7,
    Indx_Width = 8,
    Description_With = 11,
    Change_type_With = 16,
    Date_Width = 17,
    Factory_With = 17,
    Model_Name_With = 30,
    Point_With = 90
}COLUMN_With;
//Excel中各项列号
typedef enum _BOM_Column_INDEX
{
    Quantity_Column = 3,
    Point_Column = 4,
    MPN_Column = 5,
    MPN1_Column = 7,
    Factory_Column = 6,
    Factory1_Column = 8,
    Column_OFFSET =2
}Excel_Column_INDEX;
//
typedef struct _Write_Column_Index
{
    int Change_type;
    int Description_A;
    int Description_B;
    int Factory_A;
    int Factory_B;
    int Model_Name_A;
    int Model_Name_B;
    int Point_A;
    int Point_B;
    int Quantity_A;
    int Quantity_B;
    int Change_date;
    int Indx_cnt;
}Write_Column_Index;
typedef struct _BOM_Column_Index
{
    int Column_OFFSET;
    int Quantity_Column;
    int Point_Column;
    int MPN1_Column;
    int MPN_Column;
    int Factory_Column;
    int Factory1_Column;
}BOM_Column_Index;
typedef struct _Write_Column_Width
{
    int MPN_width;
    int Factory_width;
    int Description_width;
    int Point_width;
    int Quantity_width;
    int Change_type_width;
    int Date_Width;
    int Indx_Width;
}Write_Column_Width;
class Json_resolve : public QObject
{
    Q_OBJECT
public:
    explicit Json_resolve(QObject *parent = nullptr);
    void Json_Resolve(const QString file_name);
    void Json_Set_KeyValue(const QString File_Name,const QString key_name,const QString value);
    QString Json_Get_KeyValue(const QString File_Name,const QString key_name);
    bool Json_Get_Bool(const QString File_Name,const QString key_name);
    int Json_Get_Int(const QString File_Name,const QString key_name);
    double Json_Get_Float(const QString File_Name,const QString key_name);
    void Json_update(const QString File_Name);
    void BOM_Parm_Init();
    BOM_Column_Index BOM_excel_column;
    Write_Column_Width Wirte_Column_width;
    Write_Column_Index Write_Column_index;
signals:

};

#endif // JSON_RESOLVE_H
