#include "mainwindow.h"
#include "ui_mainwindow.h"
#ifdef _WIN32
#include <windows.h>
#endif
MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    str_cmp = new Qstring_cmp();
    json  = new Json_resolve();//开始解析
    cell_Item = new QTableWidgetItem;
    Write_xlsx = new QXlsx::Document;
    ui->setupUi(this);
    setWindowFlags(Qt::WindowMinimizeButtonHint|Qt::WindowCloseButtonHint); // 设置禁止最大化
    QSize win_size = Wind_Info();
    qDebug()<<"win = "<<win_size;
    this->setMinimumSize(win_size);//设置最小尺寸，数字可以随情况更改
    this->setMaximumSize(win_size);//设置最大尺寸，数字可以随情况更改
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
    ui->progressBar->setAlignment (Qt::AlignCenter);
    //测试按钮使能
    tst_btn_enable = false;
    log_enable = false;
    write_row = 1;
    tst_btn_enable = json->Json_Get_Bool(CONFIG_NAME,"测试按钮使能");
    log_enable = json->Json_Get_Bool(CONFIG_NAME,"日志记录使能");
    default_open = false;
    if(tst_btn_enable == true)
    {
        ui->pushButton_tst->setEnabled (true);
    }
    else
    {
        ui->pushButton_tst->setEnabled (false);
    }
    if(log_enable == true)//开启日志记录功能
    {
        LogHandler::Get().installMessageHandler();
    }
    ui->lineEdit_savepath->setAlignment(Qt::AlignHCenter|Qt::AlignVCenter);
    ui->lineEdit_savepath->setReadOnly (true);
    //处理状态栏
    ui->statusbar->setSizeGripEnabled(false);//去掉状态栏右下角的三角
    // 新增标签栏
    QLabel *label_info = new QLabel(this);
    // 配置显示信息
    label_info->setFrameStyle(QFrame::Box | QFrame::Sunken);
    label_info->setText(tr("作者:皇甫仁和,本工具仅限个人使用"));
    label_info->setAlignment (Qt::AlignHCenter|Qt::AlignVCenter);
    label_info->setOpenExternalLinks(false);
    // 将信息增加到底部（永久添加）
    ui->statusbar->addPermanentWidget(label_info);
}

MainWindow::~MainWindow()
{
    if(log_enable == true)
    {
        //[4] 程序结束时释放 LogHandler 的资源，例如刷新并关闭日志文件
        LogHandler::Get().uninstallMessageHandler();
    }
    delete ui;
}
 QSize MainWindow::Wind_Info()
{
    // QSize win(1265,850);
    int width = 0;
    int height = 0;
    width  = json->Json_Get_Int(CONFIG_NAME,"wind_width");
    height = json->Json_Get_Int(CONFIG_NAME,"wind_height");
    if(width < 1265)
    {
        width = 1265;
    }
    if(height < 820)
    {
        height = 820;
    }
    return QSize(width, height);
 }
QStringList MainWindow::Read_colum(QXlsx::Document *pDocument,int start_row,int column)
{
    QStringList ret;
    int row = pDocument->dimension().rowCount();
    for (int i = start_row; i < row+1; i++)
    {
        if(pDocument->cellAt(i,column)->value().toString().trimmed() !=0)
        {
            ret.append(pDocument->cellAt(i,column)->value().toString().trimmed());
        }
        else
        {
            ret.append(pDocument->cellAt(i,column+Excel_Column_INDEX::Column_OFFSET)->value().toString().trimmed());
        }
    }
    qDebug()<<"列如下:\n"<<ret;
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

void MainWindow::Excel_update()
{
    //-------------保存不同项目----------
    QFileInfo File_Info;
    QXlsx::Format format;//仅用于表头的字体格式
    format.setFontName("宋体");
    format.setFontSize(14);         // 设置字体大小
    format.setFontBold(true);       // 设置加粗
    format.setFontItalic(true);     // 设置倾斜
    format.setBorderStyle(QXlsx::Format::BorderThin);      // 设置边框
    format.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
    format.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
    format.setPatternBackgroundColor(QColor(255, 255, 255));    // 设置单元格背景颜色
    File_Info.setFile (File_Name_New);
    Write_xlsx_name = QDateTime::currentDateTime().toString("_变更记录-MMdd_hms").append (".xlsx").prepend(File_Info.path()+"/"+File_Info.baseName ());
    qDebug()<<"Write_xlsx_name is :"<<Write_xlsx_name;
    Write_xlsx->addSheet("变更履历",QXlsx::AbstractSheet::ST_WorkSheet);//对工作簿中的表格进行命名
    //-------------------------------------------------------------

    Write_xlsx->setColumnWidth(json->Write_Column_index.Change_date, json->Wirte_Column_width.Date_Width);
    Write_xlsx->setColumnWidth(json->Write_Column_index.Indx_cnt,  json->Wirte_Column_width.Indx_Width);
    Write_xlsx->setColumnWidth(json->Write_Column_index.Model_Name_A, json->Wirte_Column_width.MPN_width);
    Write_xlsx->setColumnWidth(json->Write_Column_index.Model_Name_B, json->Wirte_Column_width.MPN_width);
    Write_xlsx->setColumnWidth(json->Write_Column_index.Factory_A, json->Wirte_Column_width.Factory_width);
    Write_xlsx->setColumnWidth(json->Write_Column_index.Factory_B, json->Wirte_Column_width.Factory_width);
    Write_xlsx->setColumnWidth(json->Write_Column_index.Description_A, json->Wirte_Column_width.Description_width);
    Write_xlsx->setColumnWidth(json->Write_Column_index.Description_B, json->Wirte_Column_width.Description_width);
    Write_xlsx->setColumnWidth(json->Write_Column_index.Point_A, json->Wirte_Column_width.Point_width);
    Write_xlsx->setColumnWidth(json->Write_Column_index.Point_B, json->Wirte_Column_width.Point_width);
    Write_xlsx->setColumnWidth(json->Write_Column_index.Quantity_A, json->Wirte_Column_width.Quantity_width);
    Write_xlsx->setColumnWidth(json->Write_Column_index.Quantity_B, json->Wirte_Column_width.Quantity_width);
    Write_xlsx->setColumnWidth(json->Write_Column_index.Change_type, json->Wirte_Column_width.Change_type_width);
    Write_xlsx->setRowHeight(1,32);
    //合并单元格
    QXlsx::CellRange cellRange;//合并单元格
    //修改原有的合并单元格的方式,方便后续格式修改
    cellRange.setFirstRow(1);
    cellRange.setLastRow(1);

    cellRange.setFirstColumn(1);
    cellRange.setLastColumn(json->Write_Column_index.Quantity_B);
    Write_xlsx->mergeCells(cellRange,format);

    cellRange.setFirstColumn(json->Write_Column_index.Model_Name_A);
    cellRange.setLastColumn(json->Write_Column_index.Change_type);
    Write_xlsx->mergeCells(cellRange,format);
    //开始写表头
    format.setFontColor(Qt::red);   // 设置红色
    Write_xlsx->write (write_row,json->Write_Column_index.Model_Name_A,"变更后After\n"+File_Info.baseName (),format);

    format.setFontColor(QColor(0, 176, 240));   // 设置蓝色
    File_Info.setFile(File_Name_Old);
    Write_xlsx->write (write_row,1,"变更前Before\n"+File_Info.baseName (),format);
    write_row++;
    //
    format.setFontSize(12);         // 设置字体大小
    format.setFontBold(false);       // 设置加粗
    format.setFontItalic(false);     // 设置倾斜
    format.setFontColor(Qt::black);   // 设置黑色
    Write_xlsx->write (write_row,json->Write_Column_index.Change_date,"变更日期",format);
    Write_xlsx->write (write_row,json->Write_Column_index.Indx_cnt,"No",format);
    Write_xlsx->write (write_row,json->Write_Column_index.Model_Name_B,"物料型号",format);
    Write_xlsx->write (write_row,json->Write_Column_index.Factory_B,"厂家",format);
    Write_xlsx->write (write_row,json->Write_Column_index.Description_B,"物料描述",format);
    Write_xlsx->write (write_row,json->Write_Column_index.Point_B,"位号",format);
    Write_xlsx->write (write_row,json->Write_Column_index.Quantity_B,"用量",format);
    Write_xlsx->write (write_row,json->Write_Column_index.Model_Name_A,"物料型号",format);
    Write_xlsx->write (write_row,json->Write_Column_index.Factory_A,"厂家",format);
    Write_xlsx->write (write_row,json->Write_Column_index.Description_A,"物料描述",format);
    Write_xlsx->write (write_row,json->Write_Column_index.Point_A,"位号",format);
    Write_xlsx->write (write_row,json->Write_Column_index.Quantity_A,"用量",format);
    Write_xlsx->write (write_row,json->Write_Column_index.Change_type,"更改类型",format);
    write_row++;
    Write_xlsx->saveAs (Write_xlsx_name);
}

void MainWindow::on_pushButton_open_clicked()
{
    QString path = json->Json_Get_KeyValue(CONFIG_NAME,"变更后文件路径");
    File_Name_New = QFileDialog::getOpenFileName(this,tr("打开新版BOM文件"),path,"Excel工作簿(*.xlsx);;Excel 97-2003工作簿(*.xls)");
    if(File_Name_New.isNull())
    {
        return;
    }
    QFileInfo fileInfo(File_Name_New);
    json->Json_Set_KeyValue(CONFIG_NAME,"变更后文件路径",fileInfo.absoluteFilePath ());
    ui->lineEdit_FileName->setText (fileInfo.fileName ());
}

void MainWindow::on_pushButton_open_old_clicked()
{
    QString path = json->Json_Get_KeyValue(CONFIG_NAME,"变更前文件路径");
    File_Name_Old = QFileDialog::getOpenFileName(this,tr("打开旧版BOM文件"),path,"Excel工作簿(*.xlsx);;Excel 97-2003工作簿(*.xls)");
    if(File_Name_Old.isNull())
    {
        return;
    }
    //Read_Old_BOM(File_Name_Old);
    QFileInfo fileInfo(File_Name_Old);
    json->Json_Set_KeyValue(CONFIG_NAME,"变更前文件路径",fileInfo.absoluteFilePath ());
    ui->lineEdit_FileName_old->setText (fileInfo.fileName ());
}
/*****
注意命名:A表示新的BOM里面的数据
*/
void MainWindow::on_pushButton_open_cmp_clicked()
{
    if(File_Name_New.isNull()||File_Name_Old.isNull())
    {
        return;
    }
    read_start_row = json->Json_Get_Int(CONFIG_NAME,"start_row");
    if(read_start_row == 0 || read_start_row > 4)
    {
        read_start_row = 2;
    }
    json->Json_update (CONFIG_NAME);
    Read_New_BOM = new QXlsx::Document(File_Name_New);
    Read_Old_BOM = new QXlsx::Document(File_Name_Old);
    //-------------保存不同项目----------
    QFileInfo File_Info;
    File_Info.setFile (File_Name_New);
    QString diff_name = QDateTime::currentDateTime().toString("_变更记录-MMdd_hms").append (".xlsx").prepend(File_Info.path()+"/"+File_Info.baseName ());
    QXlsx::Document diff_xlsx(diff_name);//用于保存不同项
    diff_xlsx.addSheet("变更履历",QXlsx::AbstractSheet::ST_WorkSheet);//对工作簿中的表格进行命名

    //-------------------------------------------------------------
    diff_xlsx.setColumnWidth(json->Write_Column_index.Change_date, json->Wirte_Column_width.Date_Width);
    diff_xlsx.setColumnWidth(json->Write_Column_index.Indx_cnt, json->Wirte_Column_width.Indx_Width);
    diff_xlsx.setColumnWidth(json->Write_Column_index.Model_Name_A, json->Wirte_Column_width.MPN_width);
    diff_xlsx.setColumnWidth(json->Write_Column_index.Model_Name_B, json->Wirte_Column_width.MPN_width);
    diff_xlsx.setColumnWidth(json->Write_Column_index.Factory_A, json->Wirte_Column_width.Factory_width);
    diff_xlsx.setColumnWidth(json->Write_Column_index.Factory_B, json->Wirte_Column_width.Factory_width);
    diff_xlsx.setColumnWidth(json->Write_Column_index.Description_A, json->Wirte_Column_width.Description_width);
    diff_xlsx.setColumnWidth(json->Write_Column_index.Description_B, json->Wirte_Column_width.Description_width);
    diff_xlsx.setColumnWidth(json->Write_Column_index.Point_A,json->Wirte_Column_width.Point_width);
    diff_xlsx.setColumnWidth(json->Write_Column_index.Point_B,json->Wirte_Column_width.Point_width);
    diff_xlsx.setColumnWidth(json->Write_Column_index.Quantity_A,json->Wirte_Column_width.Quantity_width);
    diff_xlsx.setColumnWidth(json->Write_Column_index.Quantity_B,json->Wirte_Column_width.Quantity_width);
    diff_xlsx.setColumnWidth(json->Write_Column_index.Change_type,json->Wirte_Column_width.Change_type_width);
    diff_xlsx.setRowHeight(1,32);
    // 设置单元格格式
    QXlsx::Format format;//仅用于表头的字体格式
    QXlsx::Format Format_same;// 设置字体格式
    QXlsx::Format Format_diff_A;// 设置字体格式
    QXlsx::Format Format_diff_B;// 设置字体格式
    QXlsx::Format Format_cell;//单元格格式,此处需要注意,主要是用于位号的单元格格式设定
    QXlsx::CellRange cellRange;//合并单元格

    Format_same.setFontColor (Qt::black);
    Format_same.setFontBold(false);       // 设置加粗
    
    //设置自动换行
    Format_same.setTextWrap(true);
    format.setTextWrap(true);
    Format_diff_A.setTextWrap(true);
    Format_diff_B.setTextWrap(true);
    Format_cell.setTextWrap(true);
    // 设置字体
    Format_cell.setFontName ("宋体");
    format.setFontName("宋体");
    Format_same.setFontName("宋体");
    Format_diff_A.setFontName("宋体");
    Format_diff_B.setFontName("宋体");
    //设定字体大小
    Format_same.setFontSize(12);         // 设置字体大小
    Format_cell.setFontSize (12);
    Format_diff_B.setFontSize(12);         // 设置字体大小
    Format_diff_A.setFontSize(12);         // 设置字体大小
    format.setFontSize(14);         // 设置字体大小
    // 设置倾斜
    Format_diff_B.setFontItalic(false);     
    Format_diff_A.setFontItalic(false);     // 设置倾斜
    Format_same.setFontItalic(false);     // 设置倾斜
    //设置单元格边框
    Format_diff_A.setBorderStyle(QXlsx::Format::BorderThin);
    Format_diff_B.setBorderStyle(QXlsx::Format::BorderThin);
    Format_same.setBorderStyle(QXlsx::Format::BorderThin);
    format.setBorderStyle(QXlsx::Format::BorderThin);
    Format_cell.setBorderStyle (QXlsx::Format::BorderThin);

    //合并单元格中的格式
    Format_cell.setHorizontalAlignment(QXlsx::Format::AlignHCenter);// 设置水平居中
    Format_cell.setVerticalAlignment(QXlsx::Format::AlignVCenter);// 设置垂直居中
    //修改原有的合并单元格的方式,方便后续格式修改
    cellRange.setFirstRow(1);
    cellRange.setLastRow(1);

    cellRange.setFirstColumn(1);
    cellRange.setLastColumn(json->Write_Column_index.Quantity_B);
    diff_xlsx.mergeCells(cellRange,Format_cell);

    cellRange.setFirstColumn(json->Write_Column_index.Model_Name_A);
    cellRange.setLastColumn(json->Write_Column_index.Change_type);
    diff_xlsx.mergeCells(cellRange,Format_cell);
    //--------------------------------
    Format_cell.setHorizontalAlignment(QXlsx::Format::AlignLeft);// 设置水平左对齐
    format.setFontBold(true);       // 设置加粗
    format.setFontItalic(true);     // 设置倾斜
    format.setFontColor(QColor(0, 176, 240));   // 设置蓝色
    format.setPatternBackgroundColor(QColor(255, 255, 255));    // 设置单元格背景颜色
    format.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
    format.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中

    File_Info.setFile(File_Name_Old);
    diff_xlsx.write (write_row,1,"变更前Before\n"+File_Info.fileName (),format);

    File_Info.setFile(File_Name_New);
    format.setFontColor(Qt::red);   // 设置红色
    diff_xlsx.write (write_row,json->Write_Column_index.Model_Name_A,"变更后After\n"+File_Info.fileName (),format);
    write_row++;
    format.setFontItalic(false);     // 设置倾斜
    format.setFontColor(Qt::black);   // 设置红色
    format.setPatternBackgroundColor(Qt::white);    // 设置单元格背景颜色
    format.setFontBold(false);       // 设置加粗
    format.setFontSize(12);         // 设置字体大小

    diff_xlsx.write (write_row,json->Write_Column_index.Change_date,"变更日期",format);
    diff_xlsx.write (write_row,json->Write_Column_index.Indx_cnt,"No",format);
    diff_xlsx.write (write_row,json->Write_Column_index.Model_Name_B,"物料型号",format);
    diff_xlsx.write (write_row,json->Write_Column_index.Factory_B,"厂家",format);
    diff_xlsx.write (write_row,json->Write_Column_index.Description_B,"物料描述",format);
    diff_xlsx.write (write_row,json->Write_Column_index.Point_B,"位号",format);
    diff_xlsx.write (write_row,json->Write_Column_index.Quantity_B,"用量",format);
    diff_xlsx.write (write_row,json->Write_Column_index.Model_Name_A,"物料型号",format);
    diff_xlsx.write (write_row,json->Write_Column_index.Factory_A,"厂家",format);
    diff_xlsx.write (write_row,json->Write_Column_index.Description_A,"物料描述",format);
    diff_xlsx.write (write_row,json->Write_Column_index.Point_A,"位号",format);
    diff_xlsx.write (write_row,json->Write_Column_index.Quantity_A,"用量",format);
    diff_xlsx.write (write_row,json->Write_Column_index.Change_type,"更改类型",format);
    write_row++;
    //------------------------------------------------
    //直接使用list来查找不同
    //读取型号列
    QStringList mpnA_list,mpnB_list;
    //mpnA_list = Read_colum (Read_New_BOM,2,Excel_Column_INDEX::MPN_Column);
    //mpnB_list = Read_colum (Read_Old_BOM,2,Excel_Column_INDEX::MPN_Column);

    mpnA_list = Read_colum_List (File_Name_New,read_start_row,json->BOM_excel_column.MPN_Column);
    mpnB_list = Read_colum_List (File_Name_Old,read_start_row,json->BOM_excel_column.MPN_Column);
    str_cmp->CMP_set_srtlist (mpnA_list,mpnB_list);
    str_cmp->String_Cmp_list ();
    QStringList same_list = str_cmp->same_strlist;
    QStringList diffA_list = str_cmp->diffA_list;
    QStringList diffB_list = str_cmp->diffB_list;
    if(log_enable == true)//开启日志记录
    {
        qInfo()<<"  相同型号:\n"<<same_list<<"\n";
        qInfo()<<"  新增型号:\n"<<diffA_list<<"\n";
        qInfo()<<"  删除的型号:\n"<<diffB_list<<"\n";
    }
    QString Change_date_str = QDateTime::currentDateTime().toString("yyyy年MM月dd日");
    QStringList *dis_diffA_list  = new QStringList;
    QStringList *dis_diffA_Factory_list  = new QStringList;
    int dis_start = 0;
    int dis_cnt = 0;
    //先查找相同型号的变更
    int pros_range = same_list.length ()+diffA_list.length ()+diffB_list.length ();
    ui->progressBar->setRange (0,pros_range);
    QString Read_cell_A;
    QString Read_cell_B;
    QString Factory_Cell;
    QString Factory_Cell_A;
    Read_cell_A.clear ();
    Read_cell_B.clear ();
    Factory_Cell.clear ();
    Factory_Cell_A.clear ();
    int pros_cnt = 0;//处理进度
#if 1
    //先查找相同型号的变更,遍历相同型号的的位号差异
    foreach (const QString& filename, same_list)//遍历
    {
        //获取每一个所在的行;
        int row_A = mpnA_list.indexOf(filename)+read_start_row;
        int row_B = mpnB_list.indexOf(filename)+read_start_row;
        Read_cell_A = Read_New_BOM->cellAt(row_A,json->BOM_excel_column.Point_Column)->value().toString().trimmed().toUpper().remove(QRegExp("\\s"))+',';
        Read_cell_B = Read_Old_BOM->cellAt(row_B,json->BOM_excel_column.Point_Column)->value().toString().trimmed().toUpper().remove(QRegExp("\\s"))+',';
        Factory_Cell = Read_Old_BOM->cellAt(row_B,json->BOM_excel_column.Factory_Column)->value().toString().trimmed().toUpper();
        if(Factory_Cell.length () == 1||Factory_Cell.length () == 0)
        {
            Factory_Cell = Read_Old_BOM->cellAt(row_B,json->BOM_excel_column.Factory_Column+json->BOM_excel_column.Column_OFFSET)->value().toString().trimmed().toUpper();
        }

        Factory_Cell_A = Read_New_BOM->cellAt(row_A,json->BOM_excel_column.Factory_Column)->value().toString().trimmed().toUpper();
        if(Factory_Cell_A.length () == 1||Factory_Cell_A.length () == 0)
        {
            Factory_Cell_A = Read_New_BOM->cellAt(row_A,json->BOM_excel_column.Factory_Column+json->BOM_excel_column.Column_OFFSET)->value().toString().trimmed().toUpper();
        }
        //已经获取到位号,下一步位号比较
        str_cmp->CMP_set_srting (Read_cell_A,Read_cell_B);
        str_cmp->String_Cmp ();
        // 获得行尾
        if((str_cmp->diff_A.length()+str_cmp->diff_B.length ()) !=0)
        {
            //相同部分字体颜色
            Format_same.setFontColor (Qt::black);
            Format_same.setFontBold(false);       // 设置加粗
            //旧版本BOM不同部分颜色
            Format_diff_B.setFontColor (QColor(0, 176, 240));
            Format_diff_B.setFontBold(true);       // 设置加粗

            //新版BOM不同部分颜色
            Format_diff_A.setFontColor (Qt::red);
            Format_diff_A.setFontBold(true);       // 设置加粗
            //型号写入和厂家
            Format_same.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_same.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            diff_xlsx.write (write_row,json->Write_Column_index.Model_Name_A,filename,Format_same);
            diff_xlsx.write (write_row,json->Write_Column_index.Model_Name_B,filename,Format_same);
            diff_xlsx.write (write_row,json->Write_Column_index.Factory_A,Factory_Cell,Format_same);//写入厂家
            diff_xlsx.write (write_row,json->Write_Column_index.Factory_B,Factory_Cell,Format_same);//写入厂家
            //---------------写入位号----------------
            QXlsx::RichString *rich_diffA = new QXlsx::RichString(); //此处是第一个大的问题点
            QXlsx::RichString *rich_diffB = new QXlsx::RichString();
            Format_same.setFontBold (false);
            Format_same.setHorizontalAlignment(QXlsx::Format::AlignLeft); //设置左对齐
            Format_same.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
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
            diff_xlsx.write (write_row,json->Write_Column_index.Point_A,*rich_diffA,Format_cell);
            diff_xlsx.write (write_row,json->Write_Column_index.Point_B,*rich_diffB,Format_cell);

            //写入描述信息和日期
            diff_xlsx.write (write_row,json->Write_Column_index.Description_A,"",Format_cell);
            diff_xlsx.write (write_row,json->Write_Column_index.Description_B,"",Format_cell);
            diff_xlsx.write (write_row,json->Write_Column_index.Change_type,"数量变更",format);
            diff_xlsx.write (write_row,json->Write_Column_index.Change_date,Change_date_str,format);
            diff_xlsx.write (write_row,json->Write_Column_index.Indx_cnt,write_row-2,format);

            //写入数量
            Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            
            int Quantity_A = str_cmp->same_str.count (",")+str_cmp->diff_A.count (",")+1;
            diff_xlsx.write (write_row,json->Write_Column_index.Quantity_A,Quantity_A,Format_diff_A);

            Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_B.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            int Quantity_B = str_cmp->same_str.count (",")+str_cmp->diff_B.count (",")+1;
            diff_xlsx.write (write_row,json->Write_Column_index.Quantity_B,Quantity_B,Format_diff_B);

            delete rich_diffB;
            delete rich_diffA;
            write_row++;
        }
        else//相同的位号,对比厂家
        {
            //如果无差异项目,比对厂家是否相同
            if(Factory_Cell_A.compare (Factory_Cell,Qt::CaseInsensitive) != 0)//厂家不同
            {
                //相同部分字体颜色
                Format_same.setFontColor (Qt::black);
                Format_same.setFontBold(false);       // 设置加粗

                //旧版本BOM不同部分颜色
                Format_diff_B.setFontColor (QColor(0, 176, 240));
                Format_diff_B.setFontBold(true);       // 设置加粗
                Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
                Format_diff_B.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
                //新版BOM不同部分颜色
                Format_diff_A.setFontColor (Qt::red);
                Format_diff_A.setFontBold(true);       // 设置加粗
                Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
                Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
                //型号写入和厂家
                Format_same.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
                Format_same.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中

                diff_xlsx.write (write_row,json->Write_Column_index.Model_Name_A,filename,Format_same);
                diff_xlsx.write (write_row,json->Write_Column_index.Model_Name_B,filename,Format_same);

                diff_xlsx.write (write_row,json->Write_Column_index.Factory_A,Factory_Cell_A,Format_diff_A);//写入厂家
                diff_xlsx.write (write_row,json->Write_Column_index.Factory_B,Factory_Cell,Format_diff_B);//写入厂家

                int Quantity = str_cmp->same_str.count (",")+1;
                diff_xlsx.write (write_row,json->Write_Column_index.Quantity_A,Quantity,Format_same);
                diff_xlsx.write (write_row,json->Write_Column_index.Quantity_B,Quantity,Format_same);
                //写入描述信息
                diff_xlsx.write (write_row,json->Write_Column_index.Change_type,"厂家变更",format);
                diff_xlsx.write (write_row,json->Write_Column_index.Description_A,"",Format_cell);
                diff_xlsx.write (write_row,json->Write_Column_index.Description_B,"",Format_cell);
                diff_xlsx.write (write_row,json->Write_Column_index.Change_date,Change_date_str,format);
                diff_xlsx.write (write_row,json->Write_Column_index.Indx_cnt,write_row-2,format);
                //写位号
                Format_same.setHorizontalAlignment(QXlsx::Format::AlignLeft); // 设置左对齐
                diff_xlsx.write (write_row,json->Write_Column_index.Point_A,str_cmp->same_str,Format_same);
                diff_xlsx.write (write_row,json->Write_Column_index.Point_B,str_cmp->same_str,Format_same);
                write_row++;
            }
        }
        pros_cnt++;
        ui->progressBar->setValue (pros_cnt);
    }
#endif
    //遍历A(变更后)中不同项目
#if 1
    dis_start = write_row;
    foreach (const QString& filename, diffA_list)//遍历
    {
        int new_diff_row = mpnA_list.indexOf(filename)+2;
        Factory_Cell = Read_New_BOM->cellAt(new_diff_row,json->BOM_excel_column.Factory_Column)->value().toString().trimmed().toUpper();
        if(Factory_Cell.length () == 1)
        {
            Factory_Cell = Read_New_BOM->cellAt(new_diff_row,json->BOM_excel_column.Factory_Column+json->BOM_excel_column.Column_OFFSET)->value().toString().trimmed().toUpper();
        }
        Read_cell_A = Read_New_BOM->cellAt(new_diff_row,json->BOM_excel_column.Point_Column)->value().toString().trimmed().toUpper();

        //旧版本BOM不同部分颜色
        Format_diff_B.setFontColor (QColor(0, 176, 240));
        Format_diff_B.setFontBold(true);       // 设置加粗
        //新版BOM不同部分颜色
        Format_diff_A.setFontColor (Qt::red);
        Format_diff_A.setFontBold(true);       // 设置加粗

        Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
        Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中

        Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
        Format_diff_B.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
        //型号写入
        diff_xlsx.write (write_row,json->Write_Column_index.Model_Name_A,filename,Format_diff_A);
        diff_xlsx.write (write_row,json->Write_Column_index.Model_Name_B,filename,Format_diff_B);

        diff_xlsx.write (write_row,json->Write_Column_index.Factory_A,Factory_Cell,Format_diff_A);//写入厂家
        diff_xlsx.write (write_row,json->Write_Column_index.Factory_B,Factory_Cell,Format_diff_B);//写入厂家
        //写入数量
        int Quantity_A = Read_cell_A.count (",")+1;
        int Quantity_B = 0;

        diff_xlsx.write (write_row,json->Write_Column_index.Quantity_A,Quantity_A,Format_diff_A);
        diff_xlsx.write (write_row,json->Write_Column_index.Quantity_B,Quantity_B,Format_diff_B);
        //---------------写入位号----------------
        QXlsx::RichString *rich_diffA = new QXlsx::RichString();
        QXlsx::RichString *rich_diffB = new QXlsx::RichString();

        Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignLeft); // 设置水平居中
        rich_diffA->addFragment(Read_cell_A,Format_diff_A);

        Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignLeft); // 设置水平居中
        rich_diffB->addFragment ("",Format_diff_B);

        diff_xlsx.write (write_row,json->Write_Column_index.Point_A,*rich_diffA);
        diff_xlsx.write (write_row,json->Write_Column_index.Point_B,*rich_diffB);
        //写入描述信息
        diff_xlsx.write (write_row,json->Write_Column_index.Change_type,"新增物料",format);
        diff_xlsx.write (write_row,json->Write_Column_index.Description_A,"",Format_cell);
        diff_xlsx.write (write_row,json->Write_Column_index.Description_B,"",Format_cell);
        diff_xlsx.write (write_row,json->Write_Column_index.Change_date,Change_date_str,format);
        diff_xlsx.write (write_row,json->Write_Column_index.Indx_cnt,write_row-2,format);

        delete rich_diffB;
        delete rich_diffA;
        write_row++;
        pros_cnt++;
        ui->progressBar->setValue (pros_cnt);
        dis_diffA_list->append (Read_cell_A.remove(QRegExp("\\s")));
        dis_diffA_Factory_list->append (Factory_Cell);
    }
#endif
//遍历B中不同项目
#if 1  //先查找相同型号的变更
    foreach (const QString& filename, diffB_list)//遍历
    {
        QXlsx::RichString *rich_diffA = new QXlsx::RichString();
        QXlsx::RichString *rich_diffB = new QXlsx::RichString();
        int old_diff_row = mpnB_list.indexOf(filename)+2;
        Read_cell_B  = Read_Old_BOM->cellAt(old_diff_row,json->BOM_excel_column.Point_Column)->value().toString().trimmed().toUpper().remove(QRegExp("\\s"));
        Factory_Cell = Read_Old_BOM->cellAt(old_diff_row,json->BOM_excel_column.Factory_Column)->value().toString().trimmed().toUpper();
        if(Factory_Cell.length () == 1)
        {
            Factory_Cell = Read_Old_BOM->cellAt(old_diff_row,json->BOM_excel_column.Factory_Column+json->BOM_excel_column.Column_OFFSET)->value().toString().trimmed().toUpper();
        }

        dis_cnt = dis_diffA_list->indexOf (Read_cell_B.remove(QRegExp("\\s")));
        int flag =dis_cnt+1;//为0即为新增物料和位号,非0即为改变型号,位号无变化

        Read_cell_B.replace(QString(","), QString(", "));
        //------------------------------------------------------
        //旧版本BOM不同部分颜色
        Format_diff_B.setFontColor (QColor(0, 176, 240));
        //新版BOM不同部分颜色
        Format_diff_A.setFontColor (Qt::red);
        //---------------写入位号----------------
        if(!flag)//表示该部分是新增加的型号和位号
        {
            Format_diff_B.setFontBold(true);       // 设置加粗
            Format_diff_A.setFontBold(true);       // 设置加粗
            Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中

            Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_B.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            //型号写入
            diff_xlsx.write (write_row,json->Write_Column_index.Model_Name_A,filename,Format_diff_A);
            diff_xlsx.write (write_row,json->Write_Column_index.Model_Name_B,filename,Format_diff_B);

            diff_xlsx.write (write_row,json->Write_Column_index.Factory_A,Factory_Cell,Format_diff_A);//写入厂家
            diff_xlsx.write (write_row,json->Write_Column_index.Factory_B,Factory_Cell,Format_diff_B);//写入厂家
            //写入数量
            Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            int Quantity_A = 0;
            diff_xlsx.write (write_row,json->Write_Column_index.Quantity_A,Quantity_A,Format_diff_A);

            Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_B.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            int Quantity_B = Read_cell_B.count (",")+1;
            diff_xlsx.write (write_row,json->Write_Column_index.Quantity_B,Quantity_B,Format_diff_B);

            Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignLeft); // 设置水平居中
            Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            rich_diffA->addFragment("",Format_diff_A);

            Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignLeft); // 设置水平居中
            Format_diff_B.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            rich_diffB->addFragment (Read_cell_B,Format_diff_B);

            diff_xlsx.write (write_row,json->Write_Column_index.Point_A,*rich_diffA);
            diff_xlsx.write (write_row,json->Write_Column_index.Point_B,*rich_diffB);
            //写入描述信息
            diff_xlsx.write (write_row,json->Write_Column_index.Change_type,"删除物料",format);
            diff_xlsx.write (write_row,json->Write_Column_index.Description_A,"",Format_cell);
            diff_xlsx.write (write_row,json->Write_Column_index.Description_B,"",Format_cell);
            diff_xlsx.write (write_row,json->Write_Column_index.Change_date,Change_date_str,format);
            diff_xlsx.write (write_row,json->Write_Column_index.Indx_cnt,write_row-2,format);
            write_row++;
        }
        else
        {
            Format_diff_B.setFontBold(false);       // 取消加粗
            Format_diff_A.setFontBold(false);       // 取消加粗
            Format_same.setFontBold (false);// 取消加粗
            Format_same.setHorizontalAlignment(QXlsx::Format::AlignLeft); //设置左对齐
            Format_same.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            Format_same.setFontColor (Qt::black);
            rich_diffB->addFragment(Read_cell_B,Format_same);
            //位号写入
            diff_xlsx.write (dis_start+dis_cnt,json->Write_Column_index.Point_A,*rich_diffB,Format_cell);
            diff_xlsx.write (dis_start+dis_cnt,json->Write_Column_index.Point_B,*rich_diffB,Format_cell);

            //型号写入
            Format_diff_B.setFontBold(true);       // 设定加粗
            Format_diff_A.setFontBold(true);       // 设定加粗
            Format_diff_A.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_A.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中

            Format_diff_B.setHorizontalAlignment(QXlsx::Format::AlignHCenter); // 设置水平居中
            Format_diff_B.setVerticalAlignment(QXlsx::Format::AlignVCenter);   // 设置垂直居中
            //型号写入
            diff_xlsx.write (dis_start+dis_cnt,json->Write_Column_index.Model_Name_B,filename,Format_diff_B);
            //写入厂家
            //判断厂家
            if(Factory_Cell.compare (dis_diffA_Factory_list->at (dis_cnt),Qt::CaseSensitive) == 0)//相同厂家
            {
                Format_same.setHorizontalAlignment(QXlsx::Format::AlignHCenter); //设置左对齐
                diff_xlsx.write (dis_start+dis_cnt,json->Write_Column_index.Factory_A,Factory_Cell,Format_same);//写入厂家
                diff_xlsx.write (dis_start+dis_cnt,json->Write_Column_index.Factory_B,Factory_Cell,Format_same);//写入厂家
            }
            else
            {
                diff_xlsx.write (dis_start+dis_cnt,json->Write_Column_index.Factory_B,Factory_Cell,Format_diff_B);//写入厂家
            }
            //写入数量
            int Quantity = Read_cell_B.count (",")+1;
            Format_same.setHorizontalAlignment(QXlsx::Format::AlignHCenter); //设置左对齐
            diff_xlsx.write (dis_start+dis_cnt,json->Write_Column_index.Quantity_B,Quantity,Format_same);
            diff_xlsx.write (dis_start+dis_cnt,json->Write_Column_index.Quantity_A,Quantity,Format_same);
            diff_xlsx.write (dis_start+dis_cnt,json->Write_Column_index.Change_type,"更改型号",format);
        }
        delete rich_diffB;
        delete rich_diffA;
        pros_cnt++;
        ui->progressBar->setValue (pros_cnt);
    }
#endif
    //bool setRowHidden(int row, bool hidden);
    //bool setColumnHidden(int colFirst, int colLast, bool hidden);
    bool Hidden_status = json->Json_Get_Bool(CONFIG_NAME,"是否隐藏");
    if(Hidden_status == true)
    {
        diff_xlsx.setColumnHidden (json->Write_Column_index.Quantity_A,true);
        diff_xlsx.setColumnHidden (json->Write_Column_index.Quantity_B,true);
        diff_xlsx.setColumnHidden (json->Write_Column_index.Description_A,true);
        diff_xlsx.setColumnHidden (json->Write_Column_index.Description_B,true);
        diff_xlsx.setColumnHidden(json->Write_Column_index.Change_date,json->Write_Column_index.Indx_cnt,true);
    }
    diff_xlsx.save();//保存Excel
    write_row = 1;
    json->Json_Set_KeyValue(CONFIG_NAME,"比较结果文件路径",diff_name);
    //替换文件名中"/"
    diff_name.replace("/","\\");
    default_open = json->Json_Get_Bool(CONFIG_NAME,"默认文件打开使能");
    if(ui->checkBox_Autoopen->checkState ()==Qt::Checked)//如果使能默认打开文件,比较完成之后直接打开文件,可以通过json文件配置
    {
       ShellExecuteW(NULL,QString("open").toStdWString().c_str(),diff_name.toStdWString().c_str(),NULL,NULL,SW_SHOW);
    }
    ui->lineEdit_savepath->setText (diff_name);
    diff_name.clear();
    delete dis_diffA_list;
    delete dis_diffA_Factory_list;
}

void MainWindow::on_pushButton_tst_clicked()
{
    int test_state = json->Json_Get_Int(CONFIG_NAME,"test_status");
    qDebug()<<"测试项目为:"<<test_state;
    switch (test_state)
    {
        case 1:
        {
            json->Json_update (CONFIG_NAME);
        }
        break;
        case 2:
        {
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
        }
        break;
        case 3:
        {
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
            qDebug() <<"count is :"<< count;  // 输出：3
        }
        break;
        case 4:
        {
            QXlsx::Document xlsx("Demo4.xlsx");
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
        }
        break;
        case 5:
        {
            QXlsx::Document xlsx("Demo5.xlsx");//测试项目待定
            /* 使用公式 */
            xlsx.write(11, 1, "=SUM(Cell_1)");  // 计算A1-A10数据总和，并写入(11,1)单元格中
            xlsx.write(11, 2, "=SUM(Cell_2)");  // 计算B1-B10数据总和，并写入(11,2)单元格中
            //=IF(F12="","",LEN(F12)-LEN(SUBSTITUTE(F12,",",""))+1)
            xlsx.write(12, 1, "=SUM(Cell_1)*Factor");   // 计算A1-A10数据总和再乘以0.5，并写入(12,1)单元格中
            xlsx.write(16, 1, "=SUM($A$1:$A$10)*Factor");
            xlsx.write(12, 2, "=SUM(Cell_2)*Factor");   // 计算B1-B10数据总和再乘以0.5，并写入(12,2)单元格中
            xlsx.save();
        }
        break;
        case 6:
        {
            quint8 key_16[16] =  {0x2b, 0x7e, 0x15, 0x16, 0x28, 0xae, 0xd2, 0xa6, 0xab, 0xf7, 0x15, 0x88, 0x09, 0xcf, 0x4f, 0x3c};
            for (int i=0; i<16; i++)
            key16.append(key_16[i]);

            quint8 key_24[24] = { 0x8e, 0x73, 0xb0, 0xf7, 0xda, 0x0e, 0x64, 0x52, 0xc8, 0x10, 0xf3, 0x2b, 0x80, 0x90, 0x79, 0xe5, 0x62, 0xf8,
                                 0xea, 0xd2, 0x52, 0x2c, 0x6b, 0x7b};
            for (int i=0; i<24; i++)
            key24.append(key_24[i]);

            quint8 key_32[32]= { 0x60, 0x3d, 0xeb, 0x10, 0x15, 0xca, 0x71, 0xbe, 0x2b, 0x73, 0xae, 0xf0, 0x85, 0x7d, 0x77, 0x81,
                                 0x1f, 0x35, 0x2c, 0x07, 0x3b, 0x61, 0x08, 0xd7, 0x2d, 0x98, 0x10, 0xa3, 0x09, 0x14, 0xdf, 0xf4 };
            for (int i=0; i<32; i++)
            key32.append(key_32[i]);

            quint8 iv_16[16]     = {0x00, 0x01, 0x02, 0x03, 0x04, 0x05, 0x06, 0x07, 0x08, 0x09, 0x0a, 0x0b, 0x0c, 0x0d, 0x0e, 0x0f};
            for (int i=0; i<16; i++)
            iv.append(iv_16[i]);

            quint8 in_text[16]    = { 0x6b, 0xc1, 0xbe, 0xe2, 0x2e, 0x40, 0x9f, 0x96, 0xe9, 0x3d, 0x7e, 0x11, 0x73, 0x93, 0x17, 0x2a };
            quint8 out_text[16]   = { 0x3a, 0xd7, 0x7b, 0xb4, 0x0d, 0x7a, 0x36, 0x60, 0xa8, 0x9e, 0xca, 0xf3, 0x24, 0x66, 0xef, 0x97 };
            quint8 out_text_2[16] = { 0xbd, 0x33, 0x4f, 0x1d, 0x6e, 0x45, 0xf2, 0x5f, 0xf7, 0x12, 0xa2, 0x14, 0x57, 0x1f, 0xa5, 0xcc };
            quint8 out_text_3[16] = { 0xf3, 0xee, 0xd1, 0xbd, 0xb5, 0xd2, 0xa0, 0x3c, 0x06, 0x4b, 0x5a, 0x7e, 0x3d, 0xb1, 0x81, 0xf8 };
            quint8 out_text_4[16] = { 0x3b, 0x3f, 0xd9, 0x2e, 0xb7, 0x2d, 0xad, 0x20, 0x33, 0x34, 0x49, 0xf8, 0xe8, 0x3c, 0xfb, 0x4a };

            for (int i=0; i<16; i++)
            {
            in.append(in_text[i]);
            outECB128.append(out_text[i]);
            outECB192.append(out_text_2[i]);
            outECB256.append(out_text_3[i]);
            outOFB128.append(out_text_4[i]);
            }
            quint8 text_cbc[64]   = { 0x6b, 0xc1, 0xbe, 0xe2, 0x2e, 0x40, 0x9f, 0x96, 0xe9, 0x3d, 0x7e, 0x11, 0x73, 0x93, 0x17, 0x2a,
                                   0xae, 0x2d, 0x8a, 0x57, 0x1e, 0x03, 0xac, 0x9c, 0x9e, 0xb7, 0x6f, 0xac, 0x45, 0xaf, 0x8e, 0x51,
                                   0x30, 0xc8, 0x1c, 0x46, 0xa3, 0x5c, 0xe4, 0x11, 0xe5, 0xfb, 0xc1, 0x19, 0x1a, 0x0a, 0x52, 0xef,
                                   0xf6, 0x9f, 0x24, 0x45, 0xdf, 0x4f, 0x9b, 0x17, 0xad, 0x2b, 0x41, 0x7b, 0xe6, 0x6c, 0x37, 0x10 };

            quint8 output_cbc[64] = { 0x76, 0x49, 0xab, 0xac, 0x81, 0x19, 0xb2, 0x46, 0xce, 0xe9, 0x8e, 0x9b, 0x12, 0xe9, 0x19, 0x7d,
                                     0x50, 0x86, 0xcb, 0x9b, 0x50, 0x72, 0x19, 0xee, 0x95, 0xdb, 0x11, 0x3a, 0x91, 0x76, 0x78, 0xb2,
                                     0x73, 0xbe, 0xd6, 0xb8, 0xe3, 0xc1, 0x74, 0x3b, 0x71, 0x16, 0xe6, 0x9e, 0x22, 0x22, 0x95, 0x16,
                                     0x3f, 0xf1, 0xca, 0xa1, 0x68, 0x1f, 0xac, 0x09, 0x12, 0x0e, 0xca, 0x30, 0x75, 0x86, 0xe1, 0xa7 };

            for (int i=0; i<64; i++){
            inCBC128.append(text_cbc[i]);
            outCBC128.append(output_cbc[i]);
            }
            QAESEncryption encryption(QAESEncryption::AES_128, QAESEncryption::ECB);
            qDebug()<<"befor encode in"<<in;
            QByteArray encode_ret = encryption.encode(in, key16);
            qDebug()<<"after encode encode_ret"<<encode_ret;
            QAESEncryption decode(QAESEncryption::AES_128, QAESEncryption::ECB);
            qDebug()<<"befor encode outECB128"<<outECB128;
            QByteArray decode_ret = decode.decode(outECB128, key16);
            qDebug()<<"after decode decode_ret"<<decode_ret;
        }
        break;
        case 7:
        {
            read_start_row = json->Json_Get_Int(CONFIG_NAME,"start_row");
            qDebug()<<"json测试,read_start_row:"<<read_start_row;
        }
        break;
        default:
        {
            qDebug()<<"测试项目为空";
        }
        break;
    }
}


void MainWindow::on_actionrm_dup_triggered()
{
    rm_dup *RM_dup = new rm_dup();
    RM_dup->setWindowFlags (Qt::Window);
    RM_dup->show ();
}

