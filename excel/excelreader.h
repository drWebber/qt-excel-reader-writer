#ifndef EXCELREADER_H
#define EXCELREADER_H

#include <qaxobject.h>
#include <qfile.h>

class ExcelReader
{
private:
    QAxObject *excel;
    QAxObject *workbooks;
    QAxObject *workbook;
    QAxObject *sheets;
    QAxObject *sheet;
    QAxObject *usedRange;

    QString path;

public:
    ExcelReader(const QFile &file);
    ~ExcelReader();

    int sheetsCount(); /* кол-во листов в документе */
    int rowCount(); /* кол-во строк активного листа */
    int columnCount(); /* кол-во столбцов активного листа */
    QVariant readCell(int row, int column); /* чтение содержимого ячейки */

    /*
     * возвращает номер строки по первому найденному соответствию lookup_value
     * в установленном диапазоне, напр. Range(A1:A10)
     * если соответсвий не найдено - возвращает -1
     */
    int match(const QString &range, const QString &lookupValue);
};

#endif // EXCELREADER_H
