#ifndef MAINWINDOW_H
#define MAINWINDOW_H
#include <QString>
#include <QMainWindow>
#include "ExcelProcessor.h"
#include <memory>

QT_BEGIN_NAMESPACE
namespace Ui {
class MainWindow;
}
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    explicit MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

    void AddLog(QString log);

private slots:
    void ControlButtonStart(bool status); // вкл/выкл кнопки для старта
    void ControlButtonStop(bool status); // вкл/выкл кнопки для  остановки
    void ControlButtonContinueWork(bool status); // вкл/выкл кнопки для  возобновления обработки

    void on_select_file_clicked();
    void on_start_clicked();
    void on_stop_clicked();
    void on_continueWork_clicked();

private:
    Ui::MainWindow *ui;
    bool ValidationData(); //метод для проверки данных
    std::unique_ptr<ExcelProcessor> processor = nullptr;
};

#endif // MAINWINDOW_H
