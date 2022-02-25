#include "headers/mainwindow.h"
#include "ui_mainwindow.h"
#include "windows.h"

MainWindow::MainWindow(QWidget *parent) : QMainWindow(parent), ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    setWindowTitle("Карманный помощник");
    setWindowIcon(QIcon(":/icon.ico"));

    metaData = new MetaDataWidget();
    connect(metaData, &MetaDataWidget::menuWindow, this, &MainWindow::show);

    rename = new RenameWidget();
    connect(rename, &RenameWidget::menuWindow, this, &RenameWidget::show);

    QDir workdir;
    workdir.mkpath("./workdir");
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_pb_MetaData_clicked()
{    
    this->close();
    metaData->show();
}

void MainWindow::on_pb_Rename_clicked()
{
    this->close();
    rename->show();
}

