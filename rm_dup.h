#ifndef RM_DUP_H
#define RM_DUP_H

#include <QWidget>
#include <QStringList>
#include <QString>
#include <QDebug>
#include <QRegExp>
#include <QCollator>
#include <QLatin1Char>
namespace Ui {
class rm_dup;
}

class rm_dup : public QWidget
{
    Q_OBJECT

public:
    explicit rm_dup(QWidget *parent = nullptr);
    ~rm_dup();

private slots:
    void on_pushButton_CLR_clicked();

    void on_pushButton_dup_clicked();

private:
    Ui::rm_dup *ui;
};

#endif // RM_DUP_H
