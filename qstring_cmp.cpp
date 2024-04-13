#include "qstring_cmp.h"

Qstring_cmp::Qstring_cmp(QObject *parent)
    : QObject{parent}
{
    cmpA.clear ();
    cmpB.clear ();
}
void Qstring_cmp::CMP_set_srting(const QString str_cmpA,const QString str_cmpB)
{
    cmpA = str_cmpA;
    cmpB = str_cmpB;
}
void Qstring_cmp::String_Cmp()
{
    QString string_all = cmpA.replace(QString("\n"), QString(","));
    QString  stringA   = cmpA.replace(QString("\n"), QString(","));
    QString  stringB   = cmpB.replace(QString("\n"), QString(","));
    /*--------------------------------------------------------------------*/
    QStringList listA = stringA.split(QLatin1Char(','), Qt::SkipEmptyParts);
    QStringList listB = stringB.split(QLatin1Char(','), Qt::SkipEmptyParts);
    string_all = string_all.append(',');
    string_all = string_all.append(stringB);
    QStringList list_all = string_all.split(QLatin1Char(','), Qt::SkipEmptyParts);
    QStringList distin_diff;//表示不同
    QStringList distin_same;//表示重复项
    QString  string_outA;
    QString  string_outB;
/*
思路是把list中的每一个都用来比较，如果当前的值在distin_diff不存在就把该值加入distin_diff,如果存在说明该值就是重复项
*/
    foreach (const QString& filename, list_all)//遍历
    {
        if(!distin_diff.contains(filename))
        {
            distin_diff.append(filename);
        }
        else
        {
            if(!distin_same.contains(filename))
            {
                distin_same.append(filename);
            }
        }
    }
    foreach (const QString& filename, distin_same)//遍历
    {
        distin_diff.removeAt(distin_diff.indexOf(filename));
    }
    foreach (const QString& filename, distin_diff)//遍历
    {
        if(listA.indexOf(filename) != -1)//说明在A
        {
            string_outA.append(filename+ ", ");
        }
        else
        {
            string_outB.append(filename + ", ");
        }
    }
    string_outA.remove(string_outA.size()-1,1);
    if(string_outA.size() != 0)
    {
        string_outA.insert(0,", ");
    }
    string_outB.remove(string_outB.size()-1,1);
    if(string_outB.size() != 0)
    {
        string_outB.insert(0,", ");
    }
    diff_A = string_outA;
    diff_B = string_outB;
    same_str = distin_same.join(", ");
}
QString Qstring_cmp::CMP_get_diff_A()
{
    return diff_A;
}

QString Qstring_cmp::CMP_get_diff_B()
{
    return diff_B;
}
QString Qstring_cmp::CMP_get_same()
{
    return same_str;
}



void Qstring_cmp::CMP_set_srtlist(const QStringList str_listA,const QStringList str_listB)
{
    cmp_listA = str_listA;
    cmp_listB = str_listB;
}
void Qstring_cmp::String_Cmp_list()
{
    QStringList list_all;
    QStringList distin_diff;//表示不同
    QStringList distin_same;//表示重复项
    same_strlist.clear ();//此处需要注意全局变量的问题
    diffA_list.clear ();
    diffB_list.clear ();
    list_all.clear ();
    list_all = cmp_listA+cmp_listB;
    //列出相同项目和不同项
    foreach (const QString& filename, list_all)//遍历
    {
        if(!distin_diff.contains(filename))
        {
            distin_diff.append(filename);
        }
        else
        {
            if(!distin_same.contains(filename))
            {
                distin_same.append(filename);
                same_strlist.append(filename);
            }

        }
    }
    foreach (const QString& filename, distin_same)//遍历
    {
        distin_diff.removeAt(distin_diff.indexOf(filename));
    }
    //判断不同项目中各项是来至于

    foreach (const QString& filename, distin_diff)//遍历
    {
        if(cmp_listA.indexOf(filename) != -1)//说明在A
        {
            diffA_list.append(filename);
        }
        else
        {
            diffB_list.append(filename);
        }
    }

}
QStringList Qstring_cmp::CMP_get_same_list()
{
    return same_strlist;
}
QStringList Qstring_cmp::CMP_get_diffA_list()
{
    return diffA_list;
}
QStringList Qstring_cmp::CMP_get_diffB_list()
{
    return diffB_list;
}
