#include "mainwindow.h"
#include "ui_mainwindow.h"
#include <QFileDialog>
#include <QDebug>
#include <QRegularExpression>
#include "ExcelProcessor.h"
//#include "Corrector.h"
#include <QCoreApplication>
#include <thread>

/*
void MainWindow::addLogs(QString message)
{
    ui->logs->append(message);
    QCoreApplication::processEvents(QEventLoop::ExcludeUserInputEvents);
}
*/

MainWindow::MainWindow(QWidget *parent)
    : QMainWindow(parent)
    , ui(new Ui::MainWindow)
{
    ui->setupUi(this);

    //блокируем кнопки остановки и возобновления до старта обработки
    ControlButtonContinueWork(false);
    ControlButtonStop(false);
}

MainWindow::~MainWindow()
{
    delete ui;
}

//метод для добавления лога
void MainWindow::AddLog(QString log)
{
    ui->logs->append(log);
}

// вкл/выкл кнопки для старта
void MainWindow::ControlButtonStart(bool status)
{
    if(status)
    {
        ui->start->setEnabled(true);
    }
    else
    {
        ui->start->setEnabled(false);
    }
}

// вкл/выкл кнопки для остановки
void MainWindow::ControlButtonStop(bool status)
{
    if(status)
    {
        ui->stop->setEnabled(true);
    }
    else
    {
        ui->stop->setEnabled(false);
    }
}

// вкл/выкл кнопки для возобновления обработки
void MainWindow::ControlButtonContinueWork(bool status)
{
    if(status)
    {
        ui->continueWork->setEnabled(true);
    }
    else
    {
        ui->continueWork->setEnabled(false);
    }
}

// *** Нажатие на кнопку открытия окна для выбора файлов ***
void MainWindow::on_select_file_clicked()
{
    //открываем диалог выбора файла
    QString fileName = QFileDialog::getOpenFileName(this, "Выберите Excel файл", "", "Excel файлы (*.xlsx *.xls);;Все файлы (*)");

    //если файл выбран (не пустая строка)
    if (!fileName.isEmpty())
    {
        //записываем путь в поле lineEdit
        ui->patch_file->setText(fileName);

        //для отладки
        qDebug() << "Выбран файл:" << fileName;
    }
}

//метод для валидации данных
bool MainWindow::ValidationData()
{
    //проверяем выбран ли файл
    if(ui->patch_file->text().isEmpty())
    {
        AddLog("- ОШИБКА: Не выбран файл для редактирования!");
        return false;
    }

    //проверяем заполнение диапазона строк
    if(ui->start_row->value() == 0 || ui->end_row->value() == 0 || ui->start_row->value() > ui->end_row->value())
    {
        AddLog("- ОШИБКА: Некорректно указан диапазон строк!");
        AddLog("- ПОДСКАЗКА: начальная сторка - 1 и конечная строка - 100");
        return false;
    }

    //проверяем поле для колонки на русские буквы
    QRegularExpression russianLetters("[А-Яа-я]");
    QRegularExpression digits("[0-9]");
    QString col = ui->column->text().trimmed().toUpper();
    if(col.isEmpty())
    {
        AddLog("- ОШИБКА: Колонка не заполнена!");
        AddLog("- ПОДСКАЗКА: Заполните значение для колонки и повторите запуск");
        return false;
    }
    else if(col.contains(russianLetters))
    {
        AddLog("- ОШИБКА: В указанном диапазоне используются символы русского регистра!");
        AddLog("- ПОДСКАЗКА: Используйте символы английского регистра при указании колонки");
        return false;
    }
    else if(col.contains(digits))
    {
        AddLog("- ОШИБКА: Колонка не может содержать цифры!");
        AddLog("- ПОДСКАЗКА: Вводите адрес колонки в следующем формате: A, B, C");
        return false;
    }

    //проверяем заполнение типа исправляемых данных
    if(ui->typeFix->currentIndex() == -1)
    {
        AddLog("- ОШИБКА: Не выбран тип исправляемых данных!");
        return false;
    }

    return true;
}

// *** Нажание на кнопку старта ***
void MainWindow::on_start_clicked()
{
    ui->logs->clear(); //чистим поле с логами
    ControlButtonStop(true); //выключение кнопки возобновления

    if(!ValidationData())
    {
        return;
    }

    //принудительная обрабтка накопившихся событий в приложении и игнор действий пользователя.
    //QCoreApplication::processEvents(QEventLoop::ExcludeUserInputEvents);

    //создаем объект для работы с Excel
    processor = std::make_unique<ExcelProcessor>
    (
        ui->start_row->value(),
        ui->end_row->value(),
        ui->column->text().trimmed().toUpper(),
        ui->patch_file->text(),
        ui->typeFix->currentIndex(),
        ui->logger->isChecked()
    );

    connect(processor.get(), &ExcelProcessor::sendLog, this, &MainWindow::AddLog); //регистрируем сигнал для отправки логов
    connect(processor.get(), &ExcelProcessor::ControlButtonStart, this, &MainWindow::ControlButtonStart); //регистрируем сигнал для управления кнопкой старта

    //принудительно и временно запускает обработку всех ожидающих действий пользователя в текущем потоке
    //QCoreApplication::processEvents();

    std::thread thProcessor([this]()
    {
        //открываем файл
        if (!processor->openFile(processor->parameters.path))
        {
            return;
        }
        else
        {
            // Применяем исправления
            int result = processor->applyCorrectionsToColumn
            (
                processor->parameters.col,
                processor->parameters.startRow,
                processor->parameters.endRow,
                processor->parameters.typeFix
            );

            if(!result)
            {
                return;
            }
            else
            {
                processor->closeFile();
            }
        }
    });

    thProcessor.detach();

    //ControlButtonStop(false);
}

// *** остановка обработки после нажатия на кнопку "остановить" ***
void MainWindow::on_stop_clicked()
{
    AddLog("- Пауза...");
    ControlButtonContinueWork(true);
    ControlButtonStop(false);
    processor->SetStatusWork(false);
}

// *** возобновление обработки после нажатия на кнопку "возобновить" ***
void MainWindow::on_continueWork_clicked()
{
    AddLog("- Возобновление обработки файлов.");
    ControlButtonContinueWork(false);
    ControlButtonStop(true);
    processor->ContinueProcess();
    processor->SetStatusWork(true);
}
