#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    str_cmp = new Qstring_cmp();
    cell_Item = new QTableWidgetItem;
    ui->setupUi(this);
    // 每次选中一个单元格
    ui->tableWidgetdiff->setSelectionBehavior(QAbstractItemView::SelectItems);

    // 隐藏列表头
    ui->tableWidgetdiff->verticalHeader()->setVisible(false);

    // 设置隔行变色
    ui->tableWidgetdiff->setAlternatingRowColors(true);

    // 垂直方向的头不可点击
    ui->tableWidgetdiff->verticalHeader()->setSectionsClickable(false);

    // 设置固定列宽
    ui->tableWidgetdiff->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);
    // 设置固定行高
    ui->tableWidgetdiff->verticalHeader()->setSectionResizeMode(QHeaderView::Stretch);

    // 设置行表头背景颜色样式为浅黄色
    //ui->tableWidgetdiff->horizontalHeader()->setStyleSheet("QHeaderView::section{background:#ffff9b;}");
}

MainWindow::~MainWindow()
{
    delete ui;
}

/*
读取完整的列
*/
QString MainWindow::Read_colum(const QString Name)
{
    // qDebug()<<Name;
    QXlsx::Document xlsx(Name);
   // qDebug()<<"sheetNames"<<xlsx.sheetNames();
    // 获取工作簿
    QXlsx::Workbook *workBook = xlsx.workbook();
    // 获取当前工作簿的第一张sheet工作表
    QXlsx::Worksheet *workSheet = static_cast<QXlsx::Worksheet*>(workBook->sheet(0));
    // 获取当前sheet表所使用到的行数
    int row = workSheet->dimension().rowCount();
    // 获取当前sheet表所使用到的列数
   // int colum = workSheet->dimension().columnCount();
    //遍历MPN1
    QString lsit_mpn;
    for (int i = 2; i < row; i++)
    {
            QXlsx::Cell *cell = workSheet->cellAt(i, 5);    // 读取单元格
            if (cell)
            {
                lsit_mpn.append (cell->value().toString().trimmed()+',');
            }
    }
    return lsit_mpn;
}
QString MainWindow::Read_cell(const QString Name,int row, int col)
{
    // qDebug()<<Name;
    QXlsx::Document xlsx(Name);
    QXlsx::Workbook *workBook = xlsx.workbook();
    // 获取当前工作簿的第一张sheet工作表
    QXlsx::Worksheet *workSheet = static_cast<QXlsx::Worksheet*>(workBook->sheet(0));
    QString ret=NULL;
    QXlsx::Cell *cell = workSheet->cellAt(row, col);    // 读取单元格
    if (cell)
    {
        ret.append (cell->value().toString().trimmed()+',');
    }
    return ret;
}

void MainWindow::on_pushButton_open_clicked()
{
    File_Name_New = QFileDialog::getOpenFileName(this,
                                                 tr("Open files"),
                                                 "",
                                                 "Excel97(*.xlsx);;Excel(*.xls)");
    if(File_Name_New.isNull())
    {
            return;
    }
    QStringList path_list = File_Name_New.split(QLatin1Char('/'), Qt::SkipEmptyParts);
    ui->lineEdit_FileName->setText (path_list.at (path_list.count ()-1));
}
void MainWindow::on_pushButton_open_old_clicked()
{
    File_Name_Old = QFileDialog::getOpenFileName(this,
                                                 tr("Open files"),
                                                 "",
                                                 "Excel97(*.xlsx);;Excel(*.xls)");
    if(File_Name_Old.isNull())
    {
            return;
    }
    QStringList path_list = File_Name_Old.split(QLatin1Char('/'), Qt::SkipEmptyParts);
    ui->lineEdit_FileName_old->setText (path_list.at (path_list.count ()-1));
}
/*
注意命名:A表示新的BOM里面的数据

*/

void MainWindow::on_pushButton_open_cmp_clicked()
{

    if(File_Name_New.isNull()||File_Name_Old.isNull())
    {
            return;
    }
    QFont font = cell_Item->font();
    font.setBold(true);			// 设置粗体
    font.setPointSize(10);		// 设置字体大小
    cell_Item->setFont(font);	// 设置字体
    //-------------保存不同项目----------
    QDateTime current_date_time =QDateTime::currentDateTime();
    QString diff_name =current_date_time.toString("MMdd_mmsszzz").append (".xlsx");
    qDebug()<<"diff_name:"<<diff_name;
    QXlsx::Document diff_xlsx(diff_name);//用于保存不同项
    //QXlsx::Document diff_xlsx("demo.xlsx");//用于保存不同项
    diff_xlsx.setColumnWidth(Model_Name_A, Model_Name_With);
    diff_xlsx.setColumnWidth(Model_Name_B, Model_Name_With);
    diff_xlsx.setColumnWidth(Factory_A, Factory_With);
    diff_xlsx.setColumnWidth(Factory_B, Factory_With);
    diff_xlsx.setColumnWidth(Description_A, Description_With);
    diff_xlsx.setColumnWidth(Description_B, Description_With);
    diff_xlsx.setColumnWidth(Point_A, Point_With);
    diff_xlsx.setColumnWidth(Point_B, Point_With);
    diff_xlsx.setColumnWidth(Quantity_A, Quantity_With);
    diff_xlsx.setColumnWidth(Quantity_B, Quantity_With);
    diff_xlsx.setColumnWidth(Change_type, Change_type_With);
    // 设置单元格格式
    QXlsx::Format format;
    QXlsx::Format Format_same;     // 设置字体颜色
    QXlsx::Format Format_diff_A;     // 设置字体颜色
    QXlsx::Format Format_diff_B;     // 设置字体颜色
    // format2.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
    //相同部分字体颜色
    Format_same.setFontColor (Qt::black);
    Format_same.setFontBold(false);       // 设置加粗
    Format_same.setFontSize(12);         // 设置字体大小
    Format_same.setFontItalic(false);     // 设置倾斜
    Format_same.setFontName("宋体");      // 设置字体
    Format_same.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
    Format_same.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
    Format_same.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
    //旧版本BOM不同部分颜色
    Format_diff_B.setFontColor (QColor(0, 176, 240));
    Format_diff_B.setFontBold(true);       // 设置加粗
    Format_diff_B.setFontSize(12);         // 设置字体大小
    Format_diff_B.setFontItalic(false);     // 设置倾斜
    Format_diff_B.setFontName("宋体");      // 设置字体
    //Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
    Format_diff_B.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
    Format_diff_B.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
    //新版BOM不同部分颜色
    Format_diff_A.setFontColor (Qt::red);
    Format_diff_A.setFontBold(true);       // 设置加粗
    Format_diff_A.setFontSize(12);         // 设置字体大小
    Format_diff_A.setFontItalic(false);     // 设置倾斜
    Format_diff_A.setFontName("宋体");      // 设置字体
    //Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
    Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
    Format_diff_A.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
    //--------------------------------
    format.setFontBold(true);       // 设置加粗
    format.setFontSize(14);         // 设置字体大小
    format.setFontItalic(false);     // 设置倾斜
    format.setFontName("楷体");      // 设置字体
    format.setFontColor(QColor(0, 176, 240));   // 设置红色
    format.setPatternBackgroundColor(QColor(255, 255, 0));    // 设置单元格背景颜色
    format.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
    format.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
    int diff_row = 1;
    diff_xlsx.mergeCells("A1:E1");
    diff_xlsx.write ("A1","变更前Before",format);
    diff_xlsx.mergeCells("F1:J1");
    format.setFontColor(Qt::red);   // 设置红色
    format.setPatternBackgroundColor(QColor(91, 155, 213));    // 设置单元格背景颜色
    diff_xlsx.write ("F1","变更后After",format);
    diff_row++;
    format.setFontColor(Qt::black);   // 设置红色
    format.setPatternBackgroundColor(Qt::white);    // 设置单元格背景颜色
    format.setFontBold(false);       // 设置加粗
    format.setFontSize(12);         // 设置字体大小
    format.setFontItalic(false);     // 设置倾斜
    format.setFontName("宋体");      // 设置字体
    format.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
    diff_xlsx.write ("A2","物料型号",format);
    diff_xlsx.write ("B2","厂家",format);
    diff_xlsx.write ("C2","物料描述",format);
    diff_xlsx.write ("D2","位号",format);
    diff_xlsx.write ("E2","用量",format);
    //-----------------------------
    diff_xlsx.write ("F2","物料型号",format);
    diff_xlsx.write ("G2","厂家",format);
    diff_xlsx.write ("H2","物料描述",format);
    diff_xlsx.write ("I2","位号",format);
    diff_xlsx.write ("J2","用量",format);
    diff_xlsx.write ("K2","更改类型",format);
    diff_row++;
    //------------------------------------------------


    //--
    QString cmp_A = Read_colum(File_Name_New);
    QString cmp_B = Read_colum(File_Name_Old);
    QString Read_cell_A;
    QString Read_cell_B;
    //qDebug()<<"cmp_A:"<<cmp_A<<'\n';
    //qDebug()<<"cmp_B:"<<cmp_B<<'\n';
    str_cmp->CMP_set_srting (cmp_A,cmp_B);
    str_cmp->String_Cmp ();
    //qDebug()<<"same_str:"<<str_cmp->same_str<<'\n';
    //qDebug()<<"diff_A:"<<str_cmp->diff_A<<'\n';
    QStringList mpnA_list  = cmp_A.split(QLatin1Char(','), Qt::SkipEmptyParts);
    QStringList mpnB_list  = cmp_B.split(QLatin1Char(','), Qt::SkipEmptyParts);
    //
    QStringList same_list  = str_cmp->same_str.split(QLatin1Char(','), Qt::SkipEmptyParts);
    QStringList diffA_list = str_cmp->diff_A.split(QLatin1Char(','), Qt::SkipEmptyParts);
    QStringList diffB_list = str_cmp->diff_B.split(QLatin1Char(','), Qt::SkipEmptyParts);
    //先查找相同型号的变更
    //qDebug()<<"mpnA_list:"<<mpnA_list<<'\n';
    //qDebug()<<"mpnB_list:"<<mpnB_list<<'\n';
    foreach (const QString& filename, same_list)//遍历
    {
        //获取每一个所在的行;
        int row_A = mpnA_list.indexOf(filename);
        int row_B = mpnB_list.indexOf(filename);
        //qDebug()<<"MPN:"<<filename<<"row_A:"<<row_A<<"row_B:"<<row_B;
        Read_cell_A = Read_cell(File_Name_New,row_A+2,4);
        Read_cell_B = Read_cell(File_Name_Old,row_B+2,4);
        //qDebug()<<"Read_cell_A:"<<Read_cell_A;
       // qDebug()<<"Read_cell_B:"<<Read_cell_B;
        //已经获取到到位号了,下位号比较
        str_cmp->CMP_set_srting (Read_cell_A,Read_cell_B);
        str_cmp->String_Cmp ();
        //Designator 位号
        qDebug()<<"MPN:"<<filename<<"Designator same:"<<str_cmp->same_str;
        qDebug()<<"Designator diff_A:"<<str_cmp->diff_A<<"length "<<str_cmp->diff_A.length();
        qDebug()<<"Designator diff_B:"<<str_cmp->diff_B<<"length "<<str_cmp->diff_B.length();
        // 获得行尾
        int row_cnt = ui->tableWidgetdiff->rowCount();
        qDebug()<<"row_cnt"<<row_cnt;
        if((str_cmp->diff_A.length()+str_cmp->diff_B.length ()) !=0)
        {
            // 插入一行
            ui->tableWidgetdiff->insertRow(row_cnt);
            ui->tableWidgetdiff->setItem(row_cnt, COLUMN_HEAD_INDEX::Model_Name_A-1, new QTableWidgetItem(filename));
            ui->tableWidgetdiff->setItem(row_cnt, COLUMN_HEAD_INDEX::Model_Name_B-1, new QTableWidgetItem(filename));
            ui->tableWidgetdiff->setItem(row_cnt, COLUMN_HEAD_INDEX::Point_A-1, new QTableWidgetItem(str_cmp->same_str+"DF-"+str_cmp->diff_A));
            ui->tableWidgetdiff->setItem(row_cnt, COLUMN_HEAD_INDEX::Point_B-1, new QTableWidgetItem(str_cmp->same_str+"DF-"+str_cmp->diff_B));
            //diff_row
            //diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Model_Name_A-1);
            QXlsx::RichString rich_diff_B(str_cmp->same_str);
            QXlsx::RichString rich_diff_A(str_cmp->same_str);

            //rich_diff_A.addFragment (str_cmp->same_str,Format_same);
            Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中

            int Quantity_A = str_cmp->same_str.count (",")+str_cmp->diff_A.count (",");
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Quantity_A,Quantity_A,Format_diff_A);

            rich_diff_A.addFragment (str_cmp->diff_A,Format_diff_A);
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Model_Name_A,filename,Format_same);

            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Point_A,rich_diff_A);

            //rich_diff_B.addFragment (str_cmp->same_str,Format_same);
            int Quantity_B = str_cmp->same_str.count (",")+str_cmp->diff_B.count (",");
            rich_diff_B.addFragment (str_cmp->diff_B,Format_diff_B);
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Point_B,rich_diff_B);
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Model_Name_B,filename,Format_same);
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Quantity_B,Quantity_B,Format_diff_B);
            diff_row++;
        }


        //ui->tableWidgetdiff->setItem(row_cnt, COLUMN_HEAD_INDEX::Model_A-1, cell_Item);
        //cell_Item->setText (filename);
        //ui->tableWidgetdiff->setItem(row_cnt, COLUMN_HEAD_INDEX::Model_B-1, cell_Item);
        //ui->tableWidgetdiff->setItem(row, COLUMN_HEAD_INDEX::B-1, new QTableWidgetItem(QString::fromStdString(info[_B])));
    }
    diff_xlsx.save();//保存Excel
}


void MainWindow::on_pushButton_tst_clicked()
{
#if 0
    QString cmpa="C31, C33, C35, C39, C44, C45, C87,C1, C411, C414, C417, C420, C423, C426, C428, C2217, C2218, C2219";
    QString cmpb="C31, C33, C35, C39, C44, C45, C87, C106, C116, C123, C408, C411, C414, C417, C420, C423, C426, C428, C430, C2042, C2077, C2114, C2142, C2143, C2144, C2145, C2146, C2147, C2214, C2215, C2216, C2217, C2218, C2219,";
    str_cmp->CMP_set_srting (cmpa,cmpb);
    str_cmp->String_Cmp ();
    //Designator 位号
    qDebug()<<"Designator diff_A:"<<str_cmp->diff_A<<""<<str_cmp->diff_A.length();
    qDebug()<<"Designator diff_B:"<<str_cmp->diff_B<<""<<str_cmp->diff_B.length();
    qDebug()<<"Designator same:"<<str_cmp->same_str;
    QString str = "hello, world";
    int count = cmpa.count(",");
    qDebug() << count;  // 输出：3

#else
    QXlsx::Document xlsx("Text.xlsx");
    QXlsx::Format blue;     // 设置字体颜色
    blue.setFontColor(Qt::blue);
    QXlsx::Format red;
    red.setFontColor(Qt::red);
    red.setFontSize(20);    // 设置字体大小
    QXlsx::Format bold;
    bold.setFontBold(true); // 设置字体加粗

    QXlsx::RichString rich;
    rich.addFragment("test", blue);
    rich.addFragment("QT", red);
    rich.addFragment("中文", bold);
    xlsx.write(3,3, rich);
    QXlsx::Format format2;
    format2.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
    xlsx.write("C10", "测试边框", format2);
    xlsx.save();
#endif
#if 0
    /* 使用公式 */
    xlsx.write(11, 1, "=SUM(Cell_1)");  // 计算A1-A10数据总和，并写入(11,1)单元格中
    xlsx.write(11, 2, "=SUM(Cell_2)");  // 计算B1-B10数据总和，并写入(11,2)单元格中
    //=IF(F12="","",LEN(F12)-LEN(SUBSTITUTE(F12,",",""))+1)
    xlsx.write(12, 1, "=SUM(Cell_1)*Factor");   // 计算A1-A10数据总和再乘以0.5，并写入(12,1)单元格中
    xlsx.write(16, 1, "=SUM($A$1:$A$10)*Factor");
    xlsx.write(12, 2, "=SUM(Cell_2)*Factor");   // 计算B1-B10数据总和再乘以0.5，并写入(12,2)单元格中
    xlsx.save();
#endif
}

