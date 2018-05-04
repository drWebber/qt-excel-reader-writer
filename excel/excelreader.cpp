#include "excelreader.h"

#include <qregularexpression.h>

ExcelReader::ExcelReader(const QString &path)
{
    excel = new QAxObject("Excel.Application", NULL);
    workbooks = excel->querySubObject("WorkBooks");
    workbooks->dynamicCall("Open (const QString&)", path);
    workbook = excel->querySubObject("ActiveWorkBook");
    sheets = workbook->querySubObject("Worksheets");
    sheet = sheets->querySubObject("Item(int)", 1);
    usedRange = sheet->querySubObject("UsedRange");
}

ExcelReader::~ExcelReader()
{
    excel->setProperty("DisplayAlerts", false);
    workbook->dynamicCall("Close (Boolean)", true);
    excel->dynamicCall("Quit (void)");

    delete usedRange;
    delete workbook;
    delete workbooks;
    delete excel;
}

void ExcelReader::setPath(const QString &value)
{
    path = value;
}

int ExcelReader::sheetsCount()
{
    return sheets->property("Count").toInt();
}

int ExcelReader::rowCount()
{
    QAxObject *rows = usedRange->querySubObject("Rows");
    int rowsCount = rows->property("Count").toInt();
    delete rows;
    return rowsCount;
}

int ExcelReader::columnCount()
{
    QAxObject *columns = usedRange->querySubObject("Columns");
    int colsCount = columns->property("Count").toInt();
    delete columns;
    return colsCount;
}

QVariant ExcelReader::readCell(int row, int column)
{
    QAxObject *cell = sheet->querySubObject("Cells(int,int)",
                                            row, column);
    QVariant value = cell->property("Value");
    delete cell;
    return value;
}

int ExcelReader::match(const QString &range, const QString &lookupValue)
{
    QString str = "Range(" + range + ")";
    char *test = (char*)str.toUtf8().data();
    QAxObject *rng = sheet->querySubObject(test);
    QAxObject *set = rng->querySubObject("Find(const QString&)", lookupValue);

    int res = -1;
    if (set != nullptr) {
        QString str = set->dynamicCall("Address").toString();

        QRegularExpression re("\\$[A-Z]{1,10}\\$(\\d{1,10})");
        QRegularExpressionMatch match = re.match(str);
        if (match.hasMatch()) {
            res = match.captured(1).toInt();
        }
        delete set;
    }
    delete rng;
    return res;
}

QString ExcelReader::getPath() const
{
    return path;
}
