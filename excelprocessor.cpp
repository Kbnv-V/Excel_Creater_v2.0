#include "excelprocessor.h"

#include <QRegularExpression>
#include <QDir>
#include <QDebug>
#include <QStringList>
#include <QCoreApplication>
#include <QEventLoop>
#include <QMainWindow>
#include <fstream>
#include <QDate>
#include <chrono>
#include <thread>

ExcelProcessor::ExcelProcessor(int startRow, int endRow, QString col, QString path, int typeFix, bool onLogger)
{
    excel = nullptr;
    workbooks = nullptr;
    workbook = nullptr;
    worksheet = nullptr;

    //заполнение структуры данных
    parameters.startRow = startRow;
    parameters.endRow = endRow;
    parameters.col = col;
    parameters.path = path;
    parameters.typeFix = (typeFix == 0) ? Corrector::CorrectionType::EMAIL : Corrector::CorrectionType::PHONE;
    parameters.onLogger = onLogger;
}

ExcelProcessor::~ExcelProcessor()
{
    closeFile();
}

//открываем файл excel
bool ExcelProcessor::openFile(const QString& filePath)
{
    closeFile();

    ControlButtonStart(false);

    sendLog("- Открытие файла...");

    excel = new QAxObject("Excel.Application"); //запускаем приложение excel
    if (!excel || excel->isNull())
    {
        sendLog("- ОШИБКА: Не удалось запустить Excel. Убедитесь, что Excel установлен.");
        delete excel;
        excel = nullptr;
        return false;
    }

    excel->setProperty("Visible", false); //делаем окно невидимым
    excel->setProperty("DisplayAlerts", false); //отключаем показ предупреждений и диалоговых окон

    //получаем открытые файлы в приложении, которое мы открыли через new QAxObject("Excel.Application") (должен быть пустой список)
    workbooks = excel->querySubObject("Workbooks"); //возвращает COM-объект для управления открытыми документами
    if (!workbooks)
    {
        sendLog("- ОШИБКА: Не удалось получить доступ к рабочим книгам!");
        closeFile();
        return false;
    }

    //преобразуем путь к файлу в понятный формат для WIndows
    QString nativePath = QDir::toNativeSeparators(filePath);

    //открываем нужный файл. Создаем еще один COM-объект для управления конкретным файлом
    workbook = workbooks->querySubObject("Open(const QString&)", nativePath);
    if (!workbook)
    {
        sendLog("- ОШИБКА: Не удалось открыть файл: %1" + filePath);
        closeFile();
        return false;
    }

    //создаем COM-объект для управления листом в открытом документе
    worksheet = workbook->querySubObject("Worksheets(int)", 1);
    if (!worksheet)
    {
        worksheet = workbook->querySubObject("ActiveSheet");
    }
    if (!worksheet)
    {
        sendLog("- ОШИБКА: Не удалось открыть лист Excel!");
        closeFile();
        return false;
    }

    return true;
}

//закрытие файла
void ExcelProcessor::closeFile()
{
    if(worksheet)
    {
        delete worksheet;
        worksheet = nullptr;
    }

    if(workbook)
    {
        workbook->dynamicCall("Close(Bool)", false);
        delete workbook;
        workbook = nullptr;
    }

    if (workbooks)
    {
        delete workbooks;
        workbooks = nullptr;
    }

    if (excel)
    {
        excel->dynamicCall("Quit()");
        delete excel;
        excel = nullptr;
    }

    ControlButtonStart(true);
}

//чтение строк указанного столбца в массив
QVector<QString> ExcelProcessor::readColumn(const QString columnLetter, int startRow, int endRow)
{
    QVector<QString> result; //объявляем пустой массив

    //проверка открытия файла в текущем объекте
    if (!isOpen())
    {
        sendLog("- ОШИБКА: Excel файл не открыт!");
        return result;
    }

    QString normalizedColumn = columnLetter.trimmed().toUpper(); //приводим значение к нужному виду
    int col = getColumnNumber(normalizedColumn); //преобразуем букву в цифру

    //проверяем корректность данных
    if(col <= 0 || startRow <= 0 || endRow <= 0 || startRow > endRow)
    {
        sendLog("- ОШИБКА: Неверно указан столбец или диапазон строк!");
        return result;
    }

    for(int row = startRow; row <= endRow; row++)
    {
        if(!GetStatusWork()) //остановка работы программы
        {
            std::unique_lock<std::mutex> lc(mx);
            cv.wait(lc, [this] () { return statusWork; } );
        }
        QAxObject* cell = getCell(row, col);

        //добавляем полученные значения из файла в массив
        if (cell)
        {
            result.append(cell->property("Value").toString());
            delete cell;
        }
        else
        {
            result.append("null");
        }
    }

    return result;
};

//обработка полученных данных
int ExcelProcessor::applyCorrectionsToColumn(const QString& columnLetter, int startRow, int endRow, Corrector::CorrectionType type)
{
    sendLog("- Начало обработки...");

    //проверка открытия файла
    if (!isOpen())
    {
        sendLog("- ОШИБКА: Excel файл не открыт!");
        return -1;
    }

    QString normalizedColumn = columnLetter.trimmed().toUpper(); //приводим значение к нужному виду

    //читаем столбец и записываем в массив
    QVector<QString> values = readColumn(normalizedColumn, startRow, endRow);

    //проверяем заполнение массива
    if (values.isEmpty())
    {
        sendLog("- ОШИБКА: Не удалось прочитать данные из столбца!");
        return -1;
    }

    Corrector corrector; //объявляем класс
    int modifiedCells = 0;

    std::ofstream logger;

    //если логирование включено, то создается файл для записи логов
    if(parameters.onLogger)
    {
        QDate dateNow = dateNow.currentDate();
        logger = CreateFileLogs();
        AddFileLog("--------- " + dateNow.toString("dd.MM.yyyy") + " Редактирование файла по адресу " + parameters.path + " ---------", logger);
    }

    for(int i = 0; i < values.size(); ++i)
    {
        if(!GetStatusWork()) //остановка работы программы
        {
            std::unique_lock<std::mutex> lc(mx);
            cv.wait(lc, [this] () { return statusWork; } );
        }

        QString oldValue = values[i];

        if (oldValue == "null")
        {
            continue;
        }

        Corrector::CorrectionResult result = corrector.correct(oldValue, type); //исправляем значение

        if(result.correctionsCount > 0 && result.CorrectedText != oldValue)
        {
            values[i] = result.CorrectedText; //записываем исправленный вариант
            modifiedCells++;

            int row = startRow + i;
            QString cellAddress = normalizedColumn + QString::number(row); //формируем запись типа "А1"

            std::this_thread::sleep_for(std::chrono::milliseconds(30));

            //добавляем в лог
            //changeLog->append(QString("%1: \"%2\"").arg(cellAddress, result.CorrectedText));
            sendLog("[" + cellAddress + "] " + oldValue + " --> " + result.CorrectedText + "\n\n");

            if(parameters.onLogger)
            {
                AddFileLog("[" + cellAddress + "] " + oldValue + " --> " + result.CorrectedText + "\n\n", logger);
            }
            QCoreApplication::processEvents(QEventLoop::ExcludeUserInputEvents);
        }
    }

    if (modifiedCells > 0)
    {
        if (!writeColumn(normalizedColumn, startRow, values))
        {
            return -1;
        }

        workbook->dynamicCall("Save()");
    }

    CloseFileLog(logger); //закрываем файл с логами

    return modifiedCells;
};


//запись исправленных значений в файл
 bool ExcelProcessor::writeColumn(const QString& columnLetter, int startRow, const QVector<QString>& values)
{
     sendLog("- Обновление данных в файле...");

    //проверка открытия файла
    if (!isOpen())
    {
        sendLog("- ОШИБКА: Excel файл не открыт!");
        return false;
    }

    QString normalizedColumn = columnLetter.trimmed().toUpper(); //приводим значение к нужному виду
    int col = getColumnNumber(normalizedColumn); //преобразуем букву в цифру

    //проверяем корректность данных
    if(col <= 0 || startRow <= 0)
    {
        sendLog("- ОШИБКА: Неверно указан столбец или диапазон строк!");
        return false;
    }

    for(int i = 0; i < values.size(); ++i)
    {
        if(!GetStatusWork()) //остановка работы программы
        {
            std::unique_lock<std::mutex> lc(mx);
            cv.wait(lc, [this] () { return statusWork; } );
        }

        int row = startRow + i;
        QAxObject* cell = getCell(row, col);

        if(!cell)
        {
            sendLog("- ОШИБКА: Не удалось получить ячейку " + normalizedColumn + QString::number(row));
            return false;
        }

        //запись значения в ячейку файла
        cell->setProperty("Value", values[i]);
    }

    sendLog("- Обработка файла завершена.");
    return true;
}

//преобразует номер колонки (1,2,3) в буквенный адрес (A,B,C)
QString ExcelProcessor::getColumnLetter(int col)
{
    QString result;
    while (col > 0)
    {
        col--;  // Переходим к 0-индексации (A=0, B=1, ...)
        result.prepend(QChar('A' + (col % 26)));  // Добавляем букву в начало строки
        col /= 26;  // Переходим к следующему разряду (для AA, AB и т.д.)
    }
    return result;
}

//преобразует буквенный адрес колонки (A, B) в номер (1, 2, 3)
int ExcelProcessor::getColumnNumber(const QString& letter)
{
    int result = 0;
    for (QChar ch : letter)
    {
        // Каждая буква дает вклад: A=1, B=2, ... Z=26
        // Для "AB": сначала A=1, потом B = 1*26 + 2 = 28
        result = result * 26 + (ch.toUpper().toLatin1() - 'A' + 1);
    }
    return result;
}

//получает COM-объект ячейки по ее координатам
QAxObject* ExcelProcessor::getCell(int row, int col)
{
    if (!worksheet)
    {
        return nullptr;
    }

    //формируем адрес ячейки "A1", "B15"
    QString cellRef = getColumnLetter(col) + QString::number(row);
    //запрашиваем у Excel объект этой ячейки
    return worksheet->querySubObject("Range(const QString&)", cellRef);
}

//Сеттер. изменение статуса работы
void ExcelProcessor::SetStatusWork(bool status)
{
    statusWork = status;
}

//Геттер. получение статуса работы
bool ExcelProcessor::GetStatusWork()
{
    return statusWork;
}

//метод для возобновления работы
void ExcelProcessor::ContinueProcess()
{
    cv.notify_all();
}

//метод для создания файла для записи логов
std::ofstream ExcelProcessor::CreateFileLogs()
{
    std::ofstream logger("excel_creater_logger.txt", std::ofstream::app);
    return logger;
}

//метод для записи лога в созданный файл
void ExcelProcessor::AddFileLog(QString log, std::ofstream& logger)
{
    if(logger.is_open())
    {
        logger << log.toStdString() << std::endl;
    }
}

//метод для закрытия файла для логов
void ExcelProcessor::CloseFileLog(std::ofstream& logger)
{
    logger.close();
}