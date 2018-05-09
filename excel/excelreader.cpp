#include "excelreader.h"

#include <QFileInfo>
#include <qregularexpression.h>

ExcelReader::ExcelReader(const QFile &file)
{
    excel = new QAxObject("Excel.Application", NULL);
    workbooks = excel->querySubObject("WorkBooks");
    workbooks->dynamicCall("Open (const QString&)",
                           QFileInfo(file).absoluteFilePath());
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

int ExcelReader::sheetsCount() const
{
    return sheets->property("Count").toInt();
}

int ExcelReader::rowCount() const
{
    QAxObject *rows = usedRange->querySubObject("Rows");
    int rowsCount = rows->property("Count").toInt();
    delete rows;
    return rowsCount;
}

int ExcelReader::columnCount() const
{
    QAxObject *columns = usedRange->querySubObject("Columns");
    int colsCount = columns->property("Count").toInt();
    delete columns;
    return colsCount;
}

QVariant ExcelReader::readCell(int row, int column) const
{
    QAxObject *cell = sheet->querySubObject("Cells(int,int)",
                                            row, column);
    QVariant value = cell->property("Value");
    delete cell;
    return value;
}

void ExcelReader::writeCell(int row, int column, const QVariant &value) const
{
    QAxObject *cell = sheet->querySubObject("Cells(int,int)",
                                            row, column);
    cell->setProperty("Value", QVariant(value));
    delete cell;
}

void ExcelReader::save() const
{
    workbook->dynamicCall("Save()");
}

int ExcelReader::match(const QString &range, const QString &lookupValue) const
{
    QString str = "Range(" + range + ")";
    char *val = (char*)str.toUtf8().data();
    QAxObject *rng = sheet->querySubObject(val);
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
