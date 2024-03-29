#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include "headers/metadatawidget.h"
#include "headers/renamewidget.h"

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT
    QSettings *settings;
    Ui::MainWindow *ui;
    MetaDataWidget *metaData;
    RenameWidget *rename;

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

private slots:
    void on_pb_MetaData_clicked();
    void on_pb_Rename_clicked();

    void showMenuFromMetaData();
    void showMenuFromRename();
};

#endif
