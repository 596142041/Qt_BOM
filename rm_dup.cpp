#include "rm_dup.h"
#include "ui_rm_dup.h"

rm_dup::rm_dup(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::rm_dup)
{
    ui->setupUi(this);
    this->setWindowFlags(Qt::WindowMinimizeButtonHint|Qt::WindowCloseButtonHint); // 设置禁止最大化
}

rm_dup::~rm_dup()
{
    delete ui;
}

void rm_dup::on_pushButton_CLR_clicked()
{
    ui->textEdit__input->clear();
    ui->textBrowser_dup->clear();
}


void rm_dup::on_pushButton_dup_clicked()
{
    QString repeat_string = ui->textEdit__input->toPlainText().toUpper().replace(QString("\n"), QString(",")).remove(QRegExp("\\s"));
    QStringList repeat_list = repeat_string.split(QLatin1Char(','), Qt::SkipEmptyParts,Qt::CaseSensitive);
    QStringList distin_repeat;//表示重复项
    QStringList distin;
    QCollator collator;
    ui->textBrowser_dup->setReadOnly(false);
    ui->textBrowser_dup->setTextColor(QColor(255, 0, 0));
    foreach (const QString& filename, repeat_list)//遍历
    {
        if(!distin.contains(filename))
        {
            distin.append(filename);
        }
        else
        {
            if(!distin_repeat.contains(filename))
            {
                distin_repeat.append(filename);
            }
        }
    }
    std::sort(distin_repeat.begin(), distin_repeat.end(), collator);
    QString repeat_str=distin_repeat.join(", ");
    ui->textBrowser_dup->setText(repeat_str.left(repeat_str.length()));
    ui->textBrowser_dup->setReadOnly(true);
}

