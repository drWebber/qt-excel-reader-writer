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

    int sheetsCount() const; /* кол-во листов в документе */
    int rowCount() const; /* кол-во строк активного листа */
    int columnCount() const; /* кол-во столбцов активного листа */
    /* чтение содержимого ячейки */
    QVariant readCell(int row, int column) const;
    /* запись содержимого в ячейку */
    void writeCell(int row, int column,
                   const QVariant &value) const;
    void save() const;

    /*
     * возвращает номер строки по первому найденному соответствию lookup_value
     * в установленном диапазоне, напр. Range(A1:A10)
     * если соответсвий не найдено - возвращает -1
     */
    int match(const QString &range, const QString &lookupValue) const;
};

#endif // EXCELREADER_H
