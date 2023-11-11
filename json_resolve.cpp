#include "json_resolve.h"

Json_resolve::Json_resolve(QObject *parent)
    : QObject{parent}
{
    BOM_excel_column = {0};
    Wirte_Column_Width = {0};
    Write_excel_index = {0};
}
void Json_resolve::Json_Resolve(const QString file_name)
{
    QFile file_json(file_name);
    if(file_json.open (QIODevice::ReadOnly|QIODevice::Text) == false)
    {
        qDebug()<<"文件错误";

    }
    //读取所有内容
    QString str_all = file_json.readAll ();
    file_json.close ();
    // 字符串格式化为JSON
    QJsonParseError err_rpt;
    QJsonDocument  root_document = QJsonDocument::fromJson(str_all.toUtf8(), &err_rpt);
    if(err_rpt.error != QJsonParseError::NoError )//文件转换错误
    {
        qDebug()<<"错误类型:"<<err_rpt.errorString ();
    }
    // 获取到Json字符串的根节点
    QJsonObject root_object = root_document.object();
    QStringList keys = root_object.keys();//获取所有节点
    qDebug()<<"节点数量:"<<keys.count ();
    foreach (const QString& key, keys)
    {
        qDebug()<<"所有节点:"<<key;
    }
    //    // 解析blog字段
    //    QString blog = root_object.find("OUT_Excel_INDEX").value().toString();
    //    qDebug()<<"blog addr:"<<blog;
    //读取单一对象
    QJsonObject get_dict_ptr = root_object.find(tr("Column_Width")).value().toObject();
    QStringList OUT_Excel_INDEX_keys = get_dict_ptr.keys();//获取所有节点
    QVariantMap map = get_dict_ptr.toVariantMap();
    foreach (const QString& key, OUT_Excel_INDEX_keys)
    {
        qDebug()<<"当前节点:"<<key<<"vale"<<map[key].toInt();
    }
}
void Json_resolve::Json_Set_KeyValue(const QString File_Name,const QString key,const QString value)
{
    QFile file_json(File_Name);
    if(file_json.open (QFile::ReadOnly | QFile::Text) == false)
    {
        qDebug()<<"文件错误";

    }
    //读取所有内容
    QTextStream stream(&file_json);
    stream.setCodec("UTF-8");		// 设置读取编码是UTF8

    QString str_all = stream.readAll();
    file_json.close();
    // 字符串格式化为JSON
    QJsonParseError json_err;
    QJsonDocument  root_document = QJsonDocument::fromJson(str_all.toUtf8(), &json_err);
    // 获取根 { }
    QJsonObject rootObj = root_document.object();

    //修改某个节点
    rootObj[key]=value;
    //最后，再将跟节点对象{ }重新设置给QJsonDocument对象，在重新写入文件即可！
    // 将object设置为本文档的主要对象
    root_document.setObject(rootObj);

    // 重写打开文件，覆盖原有文件，达到删除文件全部内容的效果
    QFile writeFile(File_Name);
    if (!writeFile.open(QFile::WriteOnly | QFile::Truncate))
    {
         qDebug()<<"文件错误";
        return;
    }
    // 将修改后的内容写入文件
    QTextStream wirteStream(&writeFile);
    wirteStream.setCodec("UTF-8");		// 设置读取编码是UTF8
    wirteStream << root_document.toJson();		// 写入文件
    writeFile.close();					// 关闭文件
}
QString Json_resolve::Json_Get_KeyValue(const QString File_Name,const QString key)
{
    QString ret;
    QFile file_json(File_Name);
    if(file_json.open (QFile::ReadOnly | QFile::Text) == false)
    {
        qDebug()<<"文件错误";

    }
    //读取所有内容
    QTextStream stream(&file_json);
    stream.setCodec("UTF-8");		// 设置读取编码是UTF8

    QString str_all = stream.readAll();
    file_json.close();
    // 字符串格式化为JSON
    QJsonParseError json_err;
    QJsonDocument  root_document = QJsonDocument::fromJson(str_all.toUtf8(), &json_err);
    if(json_err.error != QJsonParseError::NoError )//文件转换错误
    {
        qDebug()<<"错误类型:"<<json_err.errorString ();
    }
    // 获取到Json字符串的根节点
    QJsonObject root_object = root_document.object();//根节点
    ret = root_object.find(key).value().toString();
    return ret;
}
void Json_resolve::Json_update(const QString File_Name)
{

}
