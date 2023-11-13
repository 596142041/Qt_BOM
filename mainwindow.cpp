#include "mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    str_cmp = new Qstring_cmp();
    json  = new Json_resolve();//开始解析
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
    ui->progressBar->setValue (0);
    ui->pushButton_tst->setEnabled (false);
    ui->lineEdit_savepath->setReadOnly (true);
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
    // 获取工作簿
    QXlsx::Workbook *workBook = xlsx.workbook();
    // 获取当前工作簿的第一张sheet工作表
    QXlsx::Worksheet *workSheet = static_cast<QXlsx::Worksheet*>(workBook->sheet(0));
    // 获取当前sheet表所使用到的行数
    int row = workSheet->dimension().rowCount();
    // 获取当前sheet表所使用到的列数
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
 QString MainWindow::Read_colum_All(const QString File_Name,int start_row,int column)
{
    QString ret=NULL;
    QXlsx::Document xlsx(File_Name);
    // 获取工作簿
    QXlsx::Workbook *workBook = xlsx.workbook();
    // 获取当前工作簿的第一张sheet工作表
    QXlsx::Worksheet *workSheet = static_cast<QXlsx::Worksheet*>(workBook->sheet(0));
    // 获取当前sheet表所使用到的行数
    int row = workSheet->dimension().rowCount();
    //遍历MPN1
    for (int i = start_row; i < row; i++)
    {
            QXlsx::Cell *cell = workSheet->cellAt(i, column);    // 读取单元格
            if (cell)
            {
                ret.append (cell->value().toString().trimmed().replace(QString(","), QString("/"))+',');
            }
    }
    return ret;
}
QStringList MainWindow::Read_colum_List(const QString File_Name,int start_row,int column)
{
    QStringList ret;
    QXlsx::Document xlsx(File_Name);
    // 获取工作簿
    QXlsx::Workbook *workBook = xlsx.workbook();
    // 获取当前工作簿的第一张sheet工作表
    QXlsx::Worksheet *workSheet = static_cast<QXlsx::Worksheet*>(workBook->sheet(0));
    // 获取当前sheet表所使用到的行数
    int row = workSheet->dimension().rowCount();
    //遍历MPN1
    for (int i = start_row; i < row+1; i++)
    {
            QXlsx::Cell *cell = workSheet->cellAt(i, column);    // 读取MPN单元格
            if (cell)
            {
                if(cell->value().toString().trimmed() != 0)
                {
                    ret.append (cell->value().toString().trimmed());
                }
                else
                {
                    cell = workSheet->cellAt(i, column+Excel_Column_INDEX::Column_OFFSET);    // 读取MPN1_Column单元格
                    ret.append (cell->value().toString().trimmed());
                }

            }
    }
    return ret;
}
QString MainWindow::Read_cell(const QString Name,int row, int column)
{
    // qDebug()<<Name; column
    QXlsx::Document xlsx(Name);
    QXlsx::Workbook *workBook = xlsx.workbook();
    // 获取当前工作簿的第一张sheet工作表
    QXlsx::Worksheet *workSheet = static_cast<QXlsx::Worksheet*>(workBook->sheet(0));
    QString ret=NULL;
    QXlsx::Cell *cell = workSheet->cellAt(row, column);    // 读取单元格
    if (cell)
    {
        ret.append (cell->value().toString().trimmed()+',');
    }
    return ret;
}
void MainWindow::Excel_SetTitle(QXlsx::Document *pDocument)
{
    //pDocument->write ()
}
void MainWindow::Excel_update()
{
    json->Json_update ("config.json");
    if(json->BOM_excel_column.Column_OFFSET ==0)
    {
        json->BOM_excel_column.Column_OFFSET = Excel_Column_INDEX::Column_OFFSET;
    }

}
int MainWindow::Get_Row(const QString File_Name,const QString str,int column)
{
    // qDebug()<<Name;
    QXlsx::Document xlsx(File_Name);
    // qDebug()<<"sheetNames"<<xlsx.sheetNames();
    // 获取工作簿
    QXlsx::Workbook *workBook = xlsx.workbook();
    // 获取当前工作簿的第一张sheet工作表
    QXlsx::Worksheet *workSheet = static_cast<QXlsx::Worksheet*>(workBook->sheet(0));
    // 获取当前sheet表所使用到的行数
    int row = workSheet->dimension().rowCount();
    int ret = 0;
    for (int i = 2; i < row; i++)
    {
        QXlsx::Cell *cell = workSheet->cellAt(i, column);    // 读取单元格
        if (cell)
        {

                if(str.compare (cell->value().toString().trimmed(),Qt::CaseSensitive) == 0)
                {
                    ret = i;
                    break;
                }
        }
    }
    return ret;
}
void MainWindow::on_pushButton_open_clicked()
{
    QString path = json->Json_Get_KeyValue("config.json","After_file_history");//.replace("\\","/");
    qDebug()<<"File_Name_New path"<<path;
    File_Name_New = QFileDialog::getOpenFileName(this,
                                                 tr("Open files"),
                                                 path,
                                                 "Excel97(*.xlsx);;Excel(*.xls)");
    if(File_Name_New.isNull())
    {
            return;
    }
    QFileInfo fileInfo(File_Name_New);
    json->Json_Set_KeyValue("config.json","After_file_history",fileInfo.absoluteFilePath ());
    ui->lineEdit_FileName->setText (fileInfo.fileName ());
}
void MainWindow::on_pushButton_open_old_clicked()
{
    QString path = json->Json_Get_KeyValue("config.json","Befor_file_history");
    qDebug()<<"File_Name_Old path"<<path;
    File_Name_Old = QFileDialog::getOpenFileName(this,
                                                 tr("Open files"),
                                                 path,
                                                 "Excel97(*.xlsx);;Excel(*.xls)");
    if(File_Name_Old.isNull())
    {
            return;
    }
    QFileInfo fileInfo(File_Name_Old);
    json->Json_Set_KeyValue("config.json","Befor_file_history",fileInfo.absoluteFilePath ());
    ui->lineEdit_FileName_old->setText (fileInfo.fileName ());
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
    //-------------保存不同项目----------
    QDateTime current_date_time =QDateTime::currentDateTime();
    QString diff_name =current_date_time.toString("MMdd-hh-mm-ss-zzz").append (".xlsx");
    QXlsx::Document diff_xlsx(diff_name);//用于保存不同项
    diff_xlsx.setColumnWidth(COLUMN_HEAD_INDEX::Model_Name_A, COLUMN_With::Model_Name_With);
    diff_xlsx.setColumnWidth(COLUMN_HEAD_INDEX::Model_Name_B, COLUMN_With::Model_Name_With);
    diff_xlsx.setColumnWidth(COLUMN_HEAD_INDEX::Factory_A, COLUMN_With::Factory_With);
    diff_xlsx.setColumnWidth(COLUMN_HEAD_INDEX::Factory_B, COLUMN_With::Factory_With);
    diff_xlsx.setColumnWidth(COLUMN_HEAD_INDEX::Description_A, COLUMN_With::Description_With);
    diff_xlsx.setColumnWidth(COLUMN_HEAD_INDEX::Description_B, COLUMN_With::Description_With);
    diff_xlsx.setColumnWidth(COLUMN_HEAD_INDEX::Point_A, COLUMN_With::Point_With);
    diff_xlsx.setColumnWidth(COLUMN_HEAD_INDEX::Point_B, COLUMN_With::Point_With);
    diff_xlsx.setColumnWidth(COLUMN_HEAD_INDEX::Quantity_A, COLUMN_With::Quantity_With);
    diff_xlsx.setColumnWidth(COLUMN_HEAD_INDEX::Quantity_B, COLUMN_With::Quantity_With);
    diff_xlsx.setColumnWidth(COLUMN_HEAD_INDEX::Change_type, COLUMN_With::Change_type_With);
    diff_xlsx.setRowHeight(1,20);
    // 设置单元格格式
    QXlsx::Format format;
    QXlsx::Format Format_same;// 设置字体颜色
    QXlsx::Format Format_diff_A;// 设置字体颜色
    QXlsx::Format Format_diff_B;// 设置字体颜色
    QXlsx::Format Format_cell;//单元格格式,此处需要注意
    Format_cell.setBorderStyle (QXlsx::Format::BorderThin);
    Format_cell.setHorizontalAlignment(QXlsx::Format::AlignLeft);// 设置水平左对齐
    Format_cell.setVerticalAlignment(QXlsx::Format::AlignVCenter);// 设置垂直居中
    //设置自动换行
    Format_same.setTextWarp(true);
    Format_diff_A.setTextWarp(true);
    Format_diff_B.setTextWarp(true);
    Format_cell.setTextWarp(true);
    //表头
    diff_xlsx.mergeCells("A1:E1");//合并单元格
    diff_xlsx.mergeCells("F1:J1");
    int diff_row = 1;
    //--------------------------------
    format.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
    format.setFontBold(true);       // 设置加粗
    format.setFontSize(14);         // 设置字体大小
    format.setFontItalic(false);     // 设置倾斜
    format.setFontName("楷体");      // 设置字体
    format.setFontColor(QColor(0, 176, 240));   // 设置蓝色
    format.setPatternBackgroundColor(QColor(255, 255, 255));    // 设置单元格背景颜色
    format.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
    format.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
    diff_xlsx.write ("A1","变更前Before",format);
    format.setFontColor(Qt::red);   // 设置红色
    format.setPatternBackgroundColor(QColor(255, 255, 255));    // 设置单元格背景颜色
    diff_xlsx.write ("F1","变更后After",format);
    diff_row++;
    format.setFontColor(Qt::black);   // 设置红色
    format.setPatternBackgroundColor(Qt::white);    // 设置单元格背景颜色
    format.setFontBold(false);       // 设置加粗
    format.setFontSize(12);         // 设置字体大小
    format.setFontName("宋体");      // 设置字体
    diff_xlsx.write ("A2","物料型号",format);
    diff_xlsx.write ("B2","厂家",format);
    diff_xlsx.write ("C2","物料描述",format);
    diff_xlsx.write ("D2","位号",format);
    diff_xlsx.write ("E2","用量",format);
    diff_xlsx.write ("F2","物料型号",format);
    diff_xlsx.write ("G2","厂家",format);
    diff_xlsx.write ("H2","物料描述",format);
    diff_xlsx.write ("I2","位号",format);
    diff_xlsx.write ("J2","用量",format);
    diff_xlsx.write ("K2","更改类型",format);
    diff_row++;
    //------------------------------------------------
    //直接使用list来查找不同
    //读取型号列
    QStringList mpnA_list = Read_colum_List (File_Name_New,2,Excel_Column_INDEX::MPN_Column);
    QStringList mpnB_list = Read_colum_List (File_Name_Old,2,Excel_Column_INDEX::MPN_Column);
    str_cmp->CMP_set_srtlist (mpnA_list,mpnB_list);
    str_cmp->String_Cmp_list ();
    QStringList same_list = str_cmp->same_strlist;
    QStringList diffA_list = str_cmp->diffA_list;
    QStringList diffB_list = str_cmp->diffB_list;
    //distinguish
    QStringList *dis_diffA_list  = new QStringList;
    QStringList *dis_diffA_Factory_list  = new QStringList;
    int dis_start = 0;
    int dis_cnt = 0;
    //先查找相同型号的变更
    int pros_range = same_list.length ()+diffA_list.length ()+diffB_list.length ();
    ui->progressBar->setRange (0,pros_range);
#if 0
    foreach (const QString& filename, same_list)//遍历
    {
            //获取每一个所在的行;
            int row_A = Get_Row(File_Name_New,filename,Excel_Column_INDEX::MPN_Column);
            int row_B = Get_Row(File_Name_Old,filename,Excel_Column_INDEX::MPN_Column);
            int MPN_A = mpnA_list.indexOf(filename)+2;
            int MPN_B = mpnB_list.indexOf(filename)+2;
            qDebug()<<"MPN,"<<filename<<",row_A:,"<<row_A<<",MPN_A,"<<MPN_A<<",row_B:,"<<row_B<<",MPN_B,"<<MPN_B;
    }
#endif
    QString Read_cell_A;
    QString Read_cell_B;
    QString Factory_Cell;
    QString Factory_Cell_A;
    int pros_cnt = 0;
#if 1  //先查找相同型号的变更
    //遍历相同型号的的位号差异
    foreach (const QString& filename, same_list)//遍历
    {
        //获取每一个所在的行;
        int row_A = mpnA_list.indexOf(filename)+2;
        int row_B = mpnB_list.indexOf(filename)+2;
        Read_cell_A.clear ();
        Read_cell_B.clear ();
        Read_cell_A = Read_cell(File_Name_New,row_A,Excel_Column_INDEX::Point_Column);//位号A
        Read_cell_B = Read_cell(File_Name_Old,row_B,Excel_Column_INDEX::Point_Column);//位号B
        Read_cell_A.remove(QRegExp("\\s"));//位号A,移除空格;
        Read_cell_B.remove(QRegExp("\\s"));//位号B,移除空格;
        Factory_Cell = Read_cell(File_Name_Old,row_B,Excel_Column_INDEX::Factory_Column);//厂家
        if(Factory_Cell.length () == 1)
        {
            Factory_Cell = Read_cell(File_Name_Old,row_B,Excel_Column_INDEX::Factory_Column+Excel_Column_INDEX::Column_OFFSET);//厂家
        }
        Factory_Cell.remove (Factory_Cell.size ()-1,1);

        Factory_Cell_A = Read_cell(File_Name_New,row_A,Excel_Column_INDEX::Factory_Column);//厂家
        if(Factory_Cell_A.length () == 1)
        {
            Factory_Cell_A = Read_cell(File_Name_New,row_A,Excel_Column_INDEX::Factory_Column+Excel_Column_INDEX::Column_OFFSET);//厂家
        }
        Factory_Cell_A.remove (Factory_Cell_A.size ()-1,1);
//        Factory_Cell.compare (dis_diffA_Factory_list->at (dis_cnt),Qt::CaseSensitive) == 0
        //已经获取到位号,下一步位号比较
        str_cmp->CMP_set_srting (Read_cell_A,Read_cell_B);
        str_cmp->String_Cmp ();
        // 获得行尾
        int row_cnt = ui->tableWidgetdiff->rowCount();
        if((str_cmp->diff_A.length()+str_cmp->diff_B.length ()) !=0)
        {
            // 插入一行
            ui->tableWidgetdiff->insertRow(row_cnt);
            ui->tableWidgetdiff->setItem(row_cnt, COLUMN_HEAD_INDEX::Model_Name_A-2, new QTableWidgetItem(filename));
            ui->tableWidgetdiff->setItem(row_cnt, COLUMN_HEAD_INDEX::Model_Name_B-1, new QTableWidgetItem(filename));
            ui->tableWidgetdiff->setItem(row_cnt, COLUMN_HEAD_INDEX::Point_A-3, new QTableWidgetItem(str_cmp->same_str+str_cmp->diff_A));
            ui->tableWidgetdiff->setItem(row_cnt, COLUMN_HEAD_INDEX::Point_B-2, new QTableWidgetItem(str_cmp->same_str+str_cmp->diff_B));
            //相同部分字体颜色
            Format_same.setFontColor (Qt::black);
            Format_same.setFontBold(false);       // 设置加粗
            Format_same.setFontSize(12);         // 设置字体大小
            Format_same.setFontItalic(false);     // 设置倾斜
            Format_same.setFontName("宋体");      // 设置字体
            //旧版本BOM不同部分颜色
            Format_diff_B.setFontColor (QColor(0, 176, 240));
            Format_diff_B.setFontBold(true);       // 设置加粗
            Format_diff_B.setFontSize(12);         // 设置字体大小
            Format_diff_B.setFontItalic(false);     // 设置倾斜
            Format_diff_B.setFontName("宋体");      // 设置字体
            //新版BOM不同部分颜色
            Format_diff_A.setFontColor (Qt::red);
            Format_diff_A.setFontBold(true);       // 设置加粗
            Format_diff_A.setFontSize(12);         // 设置字体大小
            Format_diff_A.setFontItalic(false);     // 设置倾斜
            Format_diff_A.setFontName("宋体");      // 设置字体
            //型号写入和厂家
            Format_same.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_same.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            Format_same.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Model_Name_A,filename,Format_same);
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Model_Name_B,filename,Format_same);
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Factory_A,Factory_Cell,Format_same);//写入厂家
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Factory_B,Factory_Cell,Format_same);//写入厂家
            //---------------写入位号----------------
            QXlsx::RichString *rich_diffA = new QXlsx::RichString(); //此处是第一个大的问题点
            QXlsx::RichString *rich_diffB = new QXlsx::RichString();

            Format_same.setFontBold (false);
            Format_same.setHorizontalAlignment(QXlsx::Format::AlignLeft); //设置左对齐
            Format_same.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            Format_same.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
            //写入相同部分
            //如果相同项目,写入相同项
            if(str_cmp->same_str.length () !=0)//存在相同项目正常写入
            {
                rich_diffA->addFragment(str_cmp->same_str,Format_same);
                rich_diffB->addFragment(str_cmp->same_str,Format_same);
            }
            else
            {
                str_cmp->diff_A.remove (0,2);
                str_cmp->diff_B.remove (0,2);
            }
            //写入不同项
            if(str_cmp->diff_A.endsWith (","))
            {
                str_cmp->diff_A.remove(str_cmp->diff_A.size()-1,1);
            }
            if(str_cmp->diff_B.endsWith (","))
            {
                str_cmp->diff_B.remove(str_cmp->diff_B.size()-1,1);
            }
            Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignLeft); //设置左对齐
            Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter); // 设置水平居中
            rich_diffA->addFragment (str_cmp->diff_A,Format_diff_A);
            Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignLeft); //设置左对齐
            Format_diff_B.setVerticalAlignment(QXlsx::Format::AlignVCenter); // 设置水平居中

            rich_diffB->addFragment (str_cmp->diff_B,Format_diff_B);

            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Point_A,*rich_diffA,Format_cell);
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Point_B,*rich_diffB,Format_cell);
            //写入描述信息
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Description_A,"",Format_cell);
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Description_B,"",Format_cell);
            //写入数量
            Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            Format_diff_A.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
            int Quantity_A = str_cmp->same_str.count (",")+str_cmp->diff_A.count (",")+1;
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Quantity_A,Quantity_A,Format_diff_A);

            Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_B.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            Format_diff_B.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
            int Quantity_B = str_cmp->same_str.count (",")+str_cmp->diff_B.count (",")+1;
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Quantity_B,Quantity_B,Format_diff_B);

            delete rich_diffB;
            delete rich_diffA;
            diff_row++;
        }
        else//相同的位号,对比厂家
        {
            //如果无差异项目,比对厂家是否相同
            if(Factory_Cell_A.compare (Factory_Cell,Qt::CaseInsensitive) != 0)//厂家不同
            {
                //相同部分字体颜色
                Format_same.setFontColor (Qt::black);
                Format_same.setFontBold(false);       // 设置加粗
                Format_same.setFontSize(12);         // 设置字体大小
                Format_same.setFontItalic(false);     // 设置倾斜
                Format_same.setFontName("宋体");      // 设置字体
                //旧版本BOM不同部分颜色
                Format_diff_B.setFontColor (QColor(0, 176, 240));
                Format_diff_B.setFontBold(true);       // 设置加粗
                Format_diff_B.setFontSize(12);         // 设置字体大小
                Format_diff_B.setFontItalic(false);     // 设置倾斜
                Format_diff_B.setFontName("宋体");      // 设置字体
                Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
                Format_diff_B.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
                Format_diff_B.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
                //新版BOM不同部分颜色
                Format_diff_A.setFontColor (Qt::red);
                Format_diff_A.setFontBold(true);       // 设置加粗
                Format_diff_A.setFontSize(12);         // 设置字体大小
                Format_diff_A.setFontItalic(false);     // 设置倾斜
                Format_diff_A.setFontName("宋体");      // 设置字体
                Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
                Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
                Format_diff_A.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
                //型号写入和厂家
                Format_same.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
                Format_same.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
                Format_same.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框

                diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Model_Name_A,filename,Format_same);
                diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Model_Name_B,filename,Format_same);

                diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Factory_A,Factory_Cell_A,Format_diff_A);//写入厂家
                diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Factory_B,Factory_Cell,Format_diff_B);//写入厂家

                int Quantity = str_cmp->same_str.count (",")+1;
                diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Quantity_A,Quantity,Format_same);
                diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Quantity_B,Quantity,Format_same);
                //写入描述信息
                diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Description_A,"",Format_cell);
                diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Description_B,"",Format_cell);
                //写位号
                Format_same.setHorizontalAlignment(QXlsx::Format::AlignLeft); // 设置左对齐
                diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Point_A,str_cmp->same_str,Format_same);
                diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Point_B,str_cmp->same_str,Format_same);
                diff_row++;
            }
        }
        pros_cnt++;
        ui->progressBar->setValue (pros_cnt);
    }

    //遍历A中不同项目
    dis_start = diff_row;
    foreach (const QString& filename, diffA_list)//遍历
    {
            int new_diff_row = mpnA_list.indexOf(filename)+2;
            Factory_Cell = Read_cell(File_Name_New,new_diff_row,Excel_Column_INDEX::Factory_Column);//厂家
            if(Factory_Cell.length () == 1)
            {
                Factory_Cell = Read_cell(File_Name_New,new_diff_row,Excel_Column_INDEX::Factory_Column+Excel_Column_INDEX::Column_OFFSET);//厂家
            }
            Read_cell_A = Read_cell(File_Name_New,new_diff_row,Point_Column);
            Read_cell_A.remove (Read_cell_A.size ()-1,1);
            Factory_Cell.remove (Factory_Cell.size ()-1,1);//移除多余的 ","
            //旧版本BOM不同部分颜色
            Format_diff_B.setFontColor (QColor(0, 176, 240));
            Format_diff_B.setFontBold(true);       // 设置加粗
            Format_diff_B.setFontSize(12);         // 设置字体大小
            Format_diff_B.setFontItalic(false);     // 设置倾斜
            Format_diff_B.setFontName("宋体");      // 设置字体
            //新版BOM不同部分颜色
            Format_diff_A.setFontColor (Qt::red);
            Format_diff_A.setFontBold(true);       // 设置加粗
            Format_diff_A.setFontSize(12);         // 设置字体大小
            Format_diff_A.setFontItalic(false);     // 设置倾斜
            Format_diff_A.setFontName("宋体");      // 设置字体


            Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            Format_diff_A.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框

            Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_B.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            Format_diff_B.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
            //型号写入
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Model_Name_A,filename,Format_diff_A);
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Model_Name_B,filename,Format_diff_B);

            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Factory_A,Factory_Cell,Format_diff_A);//写入厂家
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Factory_B,Factory_Cell,Format_diff_B);//写入厂家
            //写入数量
            int Quantity_A = Read_cell_A.count (",")+1;
            int Quantity_B = 0;

            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Quantity_A,Quantity_A,Format_diff_A);
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Quantity_B,Quantity_B,Format_diff_B);
            //---------------写入位号----------------
            QXlsx::RichString *rich_diffA = new QXlsx::RichString();
            QXlsx::RichString *rich_diffB = new QXlsx::RichString();

            Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignLeft); // 设置水平居中
            rich_diffA->addFragment(Read_cell_A,Format_diff_A);//.replace(QString(","), QString(", "))

            Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignLeft); // 设置水平居中
            rich_diffB->addFragment ("",Format_diff_B);

            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Point_A,*rich_diffA);
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Point_B,*rich_diffB);
            //写入描述信息
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Description_A,"",Format_cell);
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Description_B,"",Format_cell);
            delete rich_diffB;
            delete rich_diffA;
            diff_row++;
            pros_cnt++;
            ui->progressBar->setValue (pros_cnt);
            dis_diffA_list->append (Read_cell_A.remove(QRegExp("\\s")));
            dis_diffA_Factory_list->append (Factory_Cell);
    }
    //遍历B中不同项目
    foreach (const QString& filename, diffB_list)//遍历
    {
        QXlsx::RichString *rich_diffA = new QXlsx::RichString();
        QXlsx::RichString *rich_diffB = new QXlsx::RichString();
        int old_diff_row = mpnB_list.indexOf(filename)+2;
        Read_cell_B = Read_cell(File_Name_Old,old_diff_row,Point_Column);
        Factory_Cell = Read_cell(File_Name_Old,old_diff_row,Excel_Column_INDEX::Factory_Column);//厂家
        if(Factory_Cell.length () == 1)
        {
            Factory_Cell = Read_cell(File_Name_Old,old_diff_row,Excel_Column_INDEX::Factory_Column+Excel_Column_INDEX::Column_OFFSET);//厂家
        }
        Factory_Cell.remove (Factory_Cell.size ()-1,1);//移除多余的 ","
        Read_cell_B.remove (Read_cell_B.size ()-1,1);
        dis_cnt = dis_diffA_list->indexOf (Read_cell_B.remove(QRegExp("\\s")));
        int flag =dis_cnt+1;//为0即为新增物料和位号,非0即为改变型号,位号无变化

        Read_cell_B.replace(QString(","), QString(", "));
        //------------------------------------------------------
        //旧版本BOM不同部分颜色
        Format_diff_B.setFontColor (QColor(0, 176, 240));

        Format_diff_B.setFontSize(12);         // 设置字体大小
        Format_diff_B.setFontItalic(false);     // 设置倾斜
        Format_diff_B.setFontName("宋体");      // 设置字体
        //新版BOM不同部分颜色
        Format_diff_A.setFontColor (Qt::red);

        Format_diff_A.setFontSize(12);         // 设置字体大小
        Format_diff_A.setFontItalic(false);     // 设置倾斜
        Format_diff_A.setFontName("宋体");      // 设置字体
        //---------------写入位号----------------
        if(!flag)//表示该部分是新增加的型号和位号
        {
            Format_diff_B.setFontBold(true);       // 设置加粗
            Format_diff_A.setFontBold(true);       // 设置加粗
            Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            Format_diff_A.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框

            Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_B.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            Format_diff_B.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
            //型号写入
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Model_Name_A,filename,Format_diff_A);
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Model_Name_B,filename,Format_diff_B);

            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Factory_A,Factory_Cell,Format_diff_A);//写入厂家
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Factory_B,Factory_Cell,Format_diff_B);//写入厂家
            //写入数量
            Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            Format_diff_A.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
            int Quantity_A = 0;
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Quantity_A,Quantity_A,Format_diff_A);

            Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_B.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            Format_diff_B.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
            int Quantity_B = Read_cell_B.count (",")+1;
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Quantity_B,Quantity_B,Format_diff_B);
            Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignLeft); // 设置水平居中
            Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            Format_diff_A.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框

            rich_diffA->addFragment("",Format_diff_A);

            Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignLeft); // 设置水平居中
            Format_diff_B.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            Format_diff_B.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
            rich_diffB->addFragment (Read_cell_B,Format_diff_B);

            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Point_A,*rich_diffA);
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Point_B,*rich_diffB);
            //写入描述信息
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Description_A,"",Format_cell);
            diff_xlsx.write (diff_row,COLUMN_HEAD_INDEX::Description_B,"",Format_cell);
            diff_row++;
        }
        else
        {
            Format_diff_B.setFontBold(false);       // 取消加粗
            Format_diff_A.setFontBold(false);       // 取消加粗
            Format_same.setFontBold (false);// 取消加粗
            Format_same.setHorizontalAlignment(QXlsx::Format::AlignLeft); //设置左对齐
            Format_same.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            Format_same.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
            Format_same.setFontColor (Qt::black);
            rich_diffB->addFragment(Read_cell_B,Format_same);
            //位号写入
            diff_xlsx.write (dis_start+dis_cnt,COLUMN_HEAD_INDEX::Point_A,*rich_diffB,Format_cell);
            diff_xlsx.write (dis_start+dis_cnt,COLUMN_HEAD_INDEX::Point_B,*rich_diffB,Format_cell);

            //型号写入
            Format_diff_B.setFontBold(true);       // 设定加粗
            Format_diff_A.setFontBold(true);       // 设定加粗
            Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            Format_diff_A.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框

            Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_B.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            Format_diff_B.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
            //型号写入
            //diff_xlsx.write (dis_start+dis_cnt,COLUMN_HEAD_INDEX::Model_Name_A,filename,Format_diff_A);
            diff_xlsx.write (dis_start+dis_cnt,COLUMN_HEAD_INDEX::Model_Name_B,filename,Format_diff_B);
            //写入厂家
            //判断厂家
            if(Factory_Cell.compare (dis_diffA_Factory_list->at (dis_cnt),Qt::CaseSensitive) == 0)//相同厂家
            {
                Format_same.setHorizontalAlignment(QXlsx::Format::AlignHCenter); //设置左对齐
                diff_xlsx.write (dis_start+dis_cnt,COLUMN_HEAD_INDEX::Factory_A,Factory_Cell,Format_same);//写入厂家
                diff_xlsx.write (dis_start+dis_cnt,COLUMN_HEAD_INDEX::Factory_B,Factory_Cell,Format_same);//写入厂家
            }
            else
            {
                diff_xlsx.write (dis_start+dis_cnt,COLUMN_HEAD_INDEX::Factory_B,Factory_Cell,Format_diff_B);//写入厂家
            }
            //写入数量
            int Quantity = Read_cell_B.count (",")+1;
            Format_same.setHorizontalAlignment(QXlsx::Format::AlignHCenter); //设置左对齐
            diff_xlsx.write (dis_start+dis_cnt,COLUMN_HEAD_INDEX::Quantity_B,Quantity,Format_same);
            diff_xlsx.write (dis_start+dis_cnt,COLUMN_HEAD_INDEX::Quantity_A,Quantity,Format_same);
        }
        delete rich_diffB;
        delete rich_diffA;
        pros_cnt++;
        ui->progressBar->setValue (pros_cnt);
    }
#endif
    diff_xlsx.save();//保存Excel
    ui->lineEdit_savepath->setAlignment(Qt::AlignHCenter|Qt::AlignVCenter);
    ui->lineEdit_savepath->setText (diff_name);
    ui->lineEdit_savepath->setReadOnly (true);
    delete dis_diffA_list;
    delete dis_diffA_Factory_list;
}

void MainWindow::on_pushButton_tst_clicked()
{
    //json->Json_update ("config.json");
// QString path = json->Json_Get_KeyValue("config.json","After_file_history").replace("\\","/");
// qDebug()<<"path:"<<path;
//    QStringList mpnA_list = Read_colum_List (File_Name_New,2,Excel_Column_INDEX::MPN_Column);
//    qDebug()<<"\n mpnA_list"<<mpnA_list<<"\n";
#if 0
//    QString list_A =  Read_cell(File_Name_New,46,Excel_Column_INDEX::Point_Column);//.remove(QRegExp("\\s"));//位号A
//    QString list_B =  Read_cell(File_Name_New,100,Excel_Column_INDEX::Point_Column);//.remove(QRegExp("\\s"));//位号A;
//    qDebug()<<"list_A"<<list_A<<"\n";
//    qDebug()<<"list_B"<<list_B<<"\n";
//    str_cmp->CMP_set_srting (list_A,list_B);
//    str_cmp->String_Cmp ();
//    qDebug()<<"same_str"<<str_cmp->same_str<<"\n";
//    qDebug()<<"diff_A"<<str_cmp->diff_A<<"\n";
//    qDebug()<<"diff_B"<<str_cmp->diff_B<<"\n";

//    list_A.remove(QRegExp("\\s"));//位号A;
//    list_B.remove(QRegExp("\\s"));//位号A;
//    qDebug()<<"list_A"<<list_A<<"\n";
//    qDebug()<<"list_B"<<list_B<<"\n";
//    str_cmp->CMP_set_srting (list_A,list_B);
//    str_cmp->String_Cmp ();
//    qDebug()<<"same_str"<<str_cmp->same_str<<"\n";
//    qDebug()<<"diff_A"<<str_cmp->diff_A<<"\n";
//    qDebug()<<"diff_B"<<str_cmp->diff_B<<"\n";


    QStringList list_A;
    QStringList list_B;
    list_A.append ("C33");
    list_B.append ("C33");

    list_A.append ("C35");
    list_B.append ("C37");

    list_A.append ("C315");
    list_B.append ("C137");

    list_A.append ("C59");
    list_B.append (" C59 ");

    list_A.append ("C3 1");
    list_B.append ("C3,7");
    list_B.append ("C1/7");
    qDebug()<<"list_A:"<<list_A;
    qDebug()<<"list_B:"<<list_B;
    str_cmp->CMP_set_srtlist (list_A,list_B);
    str_cmp->String_Cmp_list ();
    QStringList same_strlist = str_cmp->same_strlist;
    QStringList diffA_list = str_cmp->diffA_list;
    QStringList diffB_list = str_cmp->diffB_list;
    qDebug()<<"same_strlist:"<<same_strlist;
    qDebug()<<"diffA_list:"<<diffA_list;
    qDebug()<<"diffB_list:"<<diffB_list;
#endif


#if 0
    QString cmpa="C31, C33, C35, C39, C44, C4 5, C87,C1, C411, C414, C417";
    QString cmpb="C31, C33, C43 35, C39, C 44, C45, C87, C106, C116, C123, C408 ";
    str_cmp->CMP_set_srting (cmpa,cmpb);
    str_cmp->String_Cmp ();
    QString string_all = cmpa.replace(QString("\n"), QString(","));//.remove(QRegExp("\\s"));
    QString  stringA   = cmpa.replace(QString("\n"), QString(","));//.remove(QRegExp("\\s"));
    QString  stringB   = cmpb.replace(QString("\n"), QString(","));//.remove(QRegExp("\\s"));

    string_all = string_all.append(',');
    string_all = string_all.append(stringB);


    QStringList listA = stringA.split(QLatin1Char(','), Qt::SkipEmptyParts);
    QStringList listB = stringB.split(QLatin1Char(','), Qt::SkipEmptyParts);
    QStringList list_all = string_all.split(QLatin1Char(','), Qt::SkipEmptyParts);
    foreach (const QString& str, listA)//遍历
    {
        qDebug()<<"listA:"<<str.simplified();
    }
    foreach (const QString& str, listB)//遍历
    {
        qDebug()<<"listB:"<<str.simplified();
    }
    foreach (const QString& str, list_all)//遍历
    {
        qDebug()<<"list_all:"<<str.simplified();
    }
    //Designator 位号
    qDebug()<<"Designator diff_A:"<<str_cmp->diff_A<<""<<str_cmp->diff_A.length();
    qDebug()<<"Designator diff_B:"<<str_cmp->diff_B<<""<<str_cmp->diff_B.length();
    qDebug()<<"Designator same:"<<str_cmp->same_str;
    int count = cmpa.count(",");
    qDebug() << count;  // 输出：3
#endif

#if 0
    QXlsx::Document xlsx("Demo.xlsx");
    QXlsx::Format blue;     // 设置字体颜色
    blue.setFontColor(Qt::blue);
    blue.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
    QXlsx::Format red;
    red.setFontColor(Qt::red);
    red.setFontSize(20);    // 设置字体大小
    red.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
    QXlsx::Format bold;
    bold.setFontBold(true); // 设置字体加粗
    bold.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
    QXlsx::Format cexx;
    cexx.setBorderStyle (QXlsx::Format::BorderThin);
    cexx.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
    cexx.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中

    QXlsx::RichString rich;
    blue.setFontColor(Qt::blue);
    blue.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
    rich.addFragment("test", blue);

    blue.setFontColor(Qt::red);
    blue.setFontSize(20);    // 设置字体大小
    blue.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
    rich.addFragment("QT", blue);


    blue.setFontSize(14);    // 设置字体大小
    blue.setFontBold(true); // 设置字体加粗
    blue.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
    rich.addFragment("tst", blue);
    xlsx.write(3,3, rich,cexx);
    QXlsx::Format format2;
    format2.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
    xlsx.write("A10", "测试\n边框", format2);
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



