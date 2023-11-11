#ifndef QSTRING_CMP_H
#define QSTRING_CMP_H

#include <QObject>
#include <QStringList>
#include <QString>
#include <QDebug>
#include <QRegExp>
#include <QCollator>
#include <QLatin1Char>
class Qstring_cmp : public QObject
{
    Q_OBJECT
public:
    explicit Qstring_cmp(QObject *parent = nullptr);
    void CMP_set_srting(const QString str_cmpA,const QString str_cmpB);
    void String_Cmp();
    QString CMP_get_same();
    QString CMP_get_diff_A();
    QString CMP_get_diff_B();
    QString same_str;
    QString diff_A;
    QString diff_B;

    void CMP_set_srtlist(const QStringList str_listA,const QStringList str_listB);
    void String_Cmp_list();
    QStringList CMP_get_same_list();
    QStringList CMP_get_diffA_list();
    QStringList CMP_get_diffB_list();
    QStringList same_strlist;
    QStringList diffA_list;
    QStringList diffB_list;


private:
    QString cmpA;
    QString cmpB;
    QStringList cmp_listA,cmp_listB;
signals:

};

#endif // QSTRING_CMP_H
