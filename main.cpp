#include <QCoreApplication>
#include <qdebug.h>

#include <excel/excelreader.h>
#include <dshow.h>

int main(int argc, char *argv[])
{
    QCoreApplication a(argc, argv);

    CoInitialize(0);

    QFile file("test.xlsx");
    if (file.exists()) {
        ExcelReader er(file);
        qDebug() << "sheets count:" << er.sheetsCount();
        qDebug() << "rowCount:" << er.rowCount();
        qDebug() << "colsCount:" <<  er.columnCount();
        int match = er.match( "A1:A3", "CA");
        qDebug() << "lookupValue row:" << match;

        const int COLUMN = 2;
        qDebug() << "cell value:" << er.readCell(match, COLUMN);
    } else {
        qDebug() << "file not found";
    }


    return a.exec();
}
