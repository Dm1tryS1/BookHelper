#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <metadatawidget.h>
#include <renamewidget.h>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void on_pb_MetaData_clicked();
    void on_pb_Rename_clicked();

private:
    Ui::MainWindow *ui;
    MetaDataWidget *metaData;
    RenameWidget *rename;
};

#endif // MAINWINDOW_H
