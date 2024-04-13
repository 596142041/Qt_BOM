#include "json_resolve.h"

Json_resolve::Json_resolve(QObject *parent)
    : QObject{parent}
{
//    BOM_excel_column   = {0};
//    Wirte_Column_width = {0};
//    Write_Column_index = {0};
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
    //解析blog字段
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
void Json_resolve::Json_Set_KeyValue(const QString File_Name,const QString key_name,const QString value)
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
    rootObj[key_name]=value;
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
bool Json_resolve::Json_Get_Bool(const QString File_Name,const QString key_name)
{
    bool ret;
    QFile file_json(File_Name);
    if(file_json.open (QFile::ReadOnly | QFile::Text) == false)
    {
        qDebug()<<"文件错误";
        return 0;
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
    ret = root_object.find(key_name).value().toBool();
    return ret;
}
int Json_resolve::Json_Get_Int(const QString File_Name,const QString key_name)
{
    int ret;
    QFile file_json(File_Name);
    if(file_json.open (QFile::ReadOnly | QFile::Text) == false)
    {
        qDebug()<<"文件错误";
        return 0;
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
        return 0;
    }
    // 获取到Json字符串的根节点
    QJsonObject root_object = root_document.object();//根节点
    ret = root_object.find(key_name).value().toInt ();
    return ret;
}
double Json_resolve::Json_Get_Float(const QString File_Name,const QString key_name)
{
    double ret;
    QFile file_json(File_Name);
    if(file_json.open (QFile::ReadOnly | QFile::Text) == false)
    {
        qDebug()<<"文件错误";
        return 0;
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
    ret = root_object.find(key_name).value().toDouble ();
    return ret;
}

QString Json_resolve::Json_Get_KeyValue(const QString File_Name,const QString key_name)
{
    QString ret;
    QFile file_json(File_Name);
    if(file_json.open (QFile::ReadOnly | QFile::Text) == false)
    {
        qDebug()<<"文件错误";
        return 0;
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
    ret = root_object.find(key_name).value().toString();
    return ret;
}

/*
获取Excel表格列参数
分别读取Write_Column_Width(写入的表格列宽度),BOM_Column(原始BOM的所在的列),Write_Column_INDEX(写入的列序列)
*/
void Json_resolve::Json_update(const QString File_Name)
{
    QFile file_json(File_Name);
    QJsonValue interestValue;
    QJsonObject interestObj;
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
    // 获取根
    QJsonObject rootObj = root_document.object();

    QStringList keys = rootObj.keys();//获取所有节点
    qDebug()<<"节点数量:"<<keys.count ();
    foreach (const QString& key, keys)
    {
        qDebug()<<"list:"<<key;
    }
    //获取各个子节点
    interestValue = rootObj.value("Write_Column_Width");
    if(interestValue.type () == QJsonValue::Object )
    {
        interestObj = interestValue.toObject();
        Wirte_Column_width.Change_type_width = interestObj.value("Change_type_Width").toInt ();
        Wirte_Column_width.Description_width = interestObj.value("Description_Width").toInt ();
        Wirte_Column_width.Factory_width = interestObj.value("Factory_Width").toInt ();
        Wirte_Column_width.MPN_width = interestObj.value("Model_Name_Width").toInt ();
        Wirte_Column_width.Point_width = interestObj.value("Point_Width").toInt ();
        Wirte_Column_width.Quantity_width = interestObj.value("Quantity_Width").toInt ();
    }
    else
    {
         qDebug()<<"Write_Column_Width OBJ ERR"<<interestValue.type ();
    }
    interestValue = rootObj.value("BOM_Column");
    if(interestValue.type () == QJsonValue::Object )
    {
        interestObj = interestValue.toObject();
        BOM_excel_column.Column_OFFSET = interestObj.value("Column_OFFSET").toInt ();
        BOM_excel_column.Factory_Column = interestObj.value("Factory_Column").toInt ();
        BOM_excel_column.Factory1_Column = interestObj.value("Factory1_Column").toInt ();
        BOM_excel_column.MPN_Column = interestObj.value("MPN_Column").toInt ();
        BOM_excel_column.MPN1_Column = interestObj.value("MPN1_Column").toInt ();
        BOM_excel_column.Point_Column = interestObj.value("Point_Column").toInt ();
        BOM_excel_column.Quantity_Column = interestObj.value("Quantity_Column").toInt ();
    }
    else
    {
        qDebug()<<"BOM_Column OBJ ERR"<<interestValue.type ();
    }
    interestValue = rootObj.value("Write_Column_INDEX");
    if(interestValue.type () == QJsonValue::Object )
    {
        interestObj = interestValue.toObject();
        Write_Column_index.Change_type = interestObj.value("Change_type").toInt ();
        Write_Column_index.Description_A = interestObj.value("Description_A").toInt ();
        Write_Column_index.Description_B = interestObj.value("Description_B").toInt ();
        Write_Column_index.Factory_A = interestObj.value("Factory_A").toInt ();
        Write_Column_index.Factory_B = interestObj.value("Factory_B").toInt ();
        Write_Column_index.Model_Name_A = interestObj.value("Model_Name_A").toInt ();
        Write_Column_index.Model_Name_B = interestObj.value("Model_Name_B").toInt ();
        Write_Column_index.Point_A = interestObj.value("Point_A").toInt ();
        Write_Column_index.Point_B = interestObj.value("Point_B").toInt ();
        Write_Column_index.Quantity_A = interestObj.value("Quantity_A").toInt ();
        Write_Column_index.Quantity_B = interestObj.value("Quantity_B").toInt ();
        qDebug()<<"Write_Column_index.Change_type:"<<Write_Column_index.Change_type;
        qDebug()<<"Write_Column_index.Description_A:"<<Write_Column_index.Description_A;
        qDebug()<<"Write_Column_index.Description_B:"<<Write_Column_index.Description_B;
        qDebug()<<"Write_Column_index.Factory_A:"<<Write_Column_index.Factory_A;
    }
    else
    {
        qDebug()<<"Write_Column_INDEX OBJ ERR"<<interestValue.type ();
    }

}
