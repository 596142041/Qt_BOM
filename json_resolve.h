#ifndef JSON_RESOLVE_H
#define JSON_RESOLVE_H

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
typedef struct _Write_HEAD_INDEX
{
    int Model_Name_B;
    int Factory_B;
    int Description_B;
    int Point_B;
    int Quantity_B;
    int Model_Name_A;
    int Factory_A;
    int Description_A;
    int Point_A;
    int Quantity_A;
    int Change_type;
}Write_HEAD_INDEX;
typedef struct _Excel_Column
{
int Quantity_Column;
int Point_Column;
int MPN_Column;
int Factory_Column;
int MPN1_Column;
int Factory1_Column;
int Column_OFFSET;
}Excel_Column;
typedef struct _Column_Width {
    int MPN_width;
    int Factory_width;
    int Description_width;
    int Point_width;
    int Quantity_width;
    int Change_type_width;
}Column_Width;
class Json_resolve : public QObject
{
    Q_OBJECT
public:
    explicit Json_resolve(QObject *parent = nullptr);
    void Json_Resolve(const QString file_name);
    void Json_Set_KeyValue(const QString File_Name,const QString key,const QString value);
    QString Json_Get_KeyValue(const QString File_Name,const QString key);
    void Json_update(const QString File_Name);
    Excel_Column BOM_excel_column;
    Column_Width Wirte_Column_Width;
    Write_HEAD_INDEX Write_excel_index;
signals:

};

#endif // JSON_RESOLVE_H
