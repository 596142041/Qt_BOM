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
int MainWindow::Get_Row(const QString File_Name,const QString str,int column)
{
    QXlsx::Document xlsx(File_Name);
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
QString MainWindow::Read_cell(const QString Name,int row, int column)
{
    QXlsx::Document xlsx(Name);
    QXlsx::Workbook *workBook = xlsx.workbook();
    // 获取当前工作簿的第一张sheet工作表
    QXlsx::Worksheet *workSheet = static_cast<QXlsx::Worksheet*>(workBook->sheet(0));
    QString ret=NULL;
    QXlsx::Cell *cell = workSheet->cellAt(row, column);    // 读取单元格
    if (cell)
    {
        ret.append (cell->value().toString().trimmed().toUpper()+',');
    }
    return ret;
}
Read_New_BOM = new QXlsx::Document(File_Name_New);
qDebug()<<"Read_New_BOM->cellAt(11,4):"<<Read_New_BOM->cellAt(11,4)->value().toString();
Read_New_BOM = new QXlsx::Document(File_Name_New);
int row_tmp =  json->Json_Get_Int("config.json","row");
int col_tmp =  json->Json_Get_Int("config.json","col");

Read_New_BOM->cellAt(row_A,Excel_Column_INDEX::Point_Column)->value().toString().trimmed().toUpper().remove(QRegExp("\\s"))+',';
Read_Old_BOM->cellAt(row_B,Excel_Column_INDEX::Point_Column)->value().toString().trimmed().toUpper().remove(QRegExp("\\s"))+',';
qDebug()<<"row_tmp"<<row_tmp<<"col_tmp"<<col_tmp;
qDebug()<<"Read_New_BOM cellAt:"<<Read_New_BOM->cellAt(row_tmp,col_tmp)->value().toString();
qDebug()<<"Read_New_BOM cell,结尾添加逗号"<<Read_New_BOM->cellAt(row_tmp,col_tmp)->value().toString().trimmed().toUpper()+',';
qDebug()<<"Read_New_BOM cell,位号A,移除空格"<<Read_New_BOM->cellAt(row_tmp,col_tmp)->value().toString().trimmed().toUpper().remove(QRegExp("\\s"))+',';
//remove(QRegExp("\\s"));//位号A,移除空格;

Read_cell_A = Read_New_BOM->cellAt(row_A,Excel_Column_INDEX::Point_Column)->value().toString().trimmed().toUpper().remove(QRegExp("\\s"))+',';
Read_cell_B = Read_Old_BOM->cellAt(row_B,Excel_Column_INDEX::Point_Column)->value().toString().trimmed().toUpper().remove(QRegExp("\\s"))+',';


Read_cell_A = Read_New_BOM->cellAt(row_A,Excel_Column_INDEX::Point_Column)->value().toString().trimmed().toUpper().remove(QRegExp("\\s"))+',';
Read_cell_B = Read_Old_BOM->cellAt(row_B,Excel_Column_INDEX::Point_Column)->value().toString().trimmed().toUpper().remove(QRegExp("\\s"))+',';
//Factory_Cell = Read_cell(File_Name_Old,row_B,Excel_Column_INDEX::Factory_Column);//厂家
Factory_Cell = Read_Old_BOM->cellAt(row_B,Excel_Column_INDEX::Factory_Column)->value().toString().trimmed().toUpper();
if(Factory_Cell.length () == 1)
{
    //Factory_Cell = Read_cell(File_Name_Old,row_B,Excel_Column_INDEX::Factory_Column+Excel_Column_INDEX::Column_OFFSET);//厂家
    Factory_Cell = Read_Old_BOM->cellAt(row_B,Excel_Column_INDEX::Factory_Column+Excel_Column_INDEX::Column_OFFSET)->value().toString().trimmed().toUpper();
}

Factory_Cell_A = Read_New_BOM->cellAt(row_B,Excel_Column_INDEX::Factory_Column)->value().toString().trimmed().toUpper();
if(Factory_Cell_A.length () == 1)
{
    Factory_Cell_A = Read_New_BOM->cellAt(row_B,Excel_Column_INDEX::Factory_Column+Excel_Column_INDEX::Column_OFFSET)->value().toString().trimmed().toUpper();
}

QString MainWindow::Read_cell(const QString Name,int row, int column)
{
    QXlsx::Document xlsx(Name);
    QXlsx::Workbook *workBook = xlsx.workbook();
    // 获取当前工作簿的第一张sheet工作表
    QXlsx::Worksheet *workSheet = static_cast<QXlsx::Worksheet*>(workBook->sheet(0));
    QString ret=NULL;
    QXlsx::Cell *cell = workSheet->cellAt(row, column);    // 读取单元格
    if (cell)
    {
        ret.append (cell->value().toString().trimmed().toUpper()+',');
    }
    return ret;
}