#ifndef EXCELPROCESSOR_H
#define EXCELPROCESSOR_H

#include <QString>
#include <QVector>
#include <QAxObject>
#include "corrector.h"
#include <condition_variable>
#include <mutex>


class ExcelProcessor : public QObject
{
    Q_OBJECT

public:
    ExcelProcessor(int startRow, int endRow, QString col, QString path, int typeFix, bool onLogger);
    ~ExcelProcessor();

    bool openFile(const QString& filePath);
    void closeFile();

    QVector<QString> readColumn(const QString columnLetter, int startRow, int endRow);
    bool writeColumn(const QString& columnLetter, int startRow, const QVector<QString>& values);

    int applyCorrectionsToColumn(const QString& columnLetter, int startRow, int endRow, Corrector::CorrectionType type);

    bool isOpen() const { return excel != nullptr && workbook != nullptr; }

    void SetStatusWork(bool status);
    bool GetStatusWork();
    void ContinueProcess();

    //структура для данных с формы
    struct Data
    {
        int startRow;
        int endRow;
        QString col;
        QString path;
        Corrector::CorrectionType typeFix;
        bool onLogger;
    };

    Data parameters;

private:
    QAxObject* excel;
    QAxObject* workbooks;
    QAxObject* workbook;
    QAxObject* worksheet;

    QString getColumnLetter(int col);
    int getColumnNumber(const QString& letter);
    QAxObject* getCell(int row, int col);

    std::ofstream CreateFileLogs();
    void AddFileLog(QString log, std::ofstream& logger);
    void CloseFileLog(std::ofstream& logger);

    std::condition_variable cv;
    std::mutex mx;

    bool statusWork = true;

signals:
    void sendLog(const QString log);
    void ControlButtonStart(const bool status);
};

#endif // EXCELPROCESSOR_H
