#include "headers/mainwindow.h"
#include "ui_mainwindow.h"

MainWindow::MainWindow(QWidget *parent) : QMainWindow(parent), ui(new Ui::MainWindow)
{
    ui->setupUi(this);
    setWindowTitle("Карманный помощник");
    setWindowIcon(QIcon(":/icon.ico"));

    metaData = new MetaDataWidget();
    connect(metaData, &MetaDataWidget::menuWindow, this, &MainWindow::showMenuFromMetaData);

    rename = new RenameWidget();
    connect(rename, &RenameWidget::menuWindow, this, &MainWindow::showMenuFromRename);

    QDir workdir;
    workdir.mkpath("./workdir");
    settings = new QSettings("./workdir/settings.conf", QSettings::IniFormat);
}

MainWindow::~MainWindow()
{
    delete ui;
}

void MainWindow::on_pb_MetaData_clicked()
{
    if (!settings->value("MetaData/inputPath").toString().isEmpty())
        metaData->setInputPath(settings->value("MetaData/inputPath").toString());
    else
         metaData->setInputPath("C:/");

    this->close();
    metaData->show();

}

void MainWindow::on_pb_Rename_clicked()
{
    if (!settings->value("Rename/dirPath").toString().isEmpty())
        rename->setDirectoryPath(settings->value("Rename/dirPath").toString());
    else
        rename->setDirectoryPath("C:/");

    if (!settings->value("Rename/filePath").toString().isEmpty())
        rename->setFilePath(settings->value("Rename/filePath").toString());
    else
        rename->setFilePath("C:/");

    if (!settings->value("Rename/resultPath").toString().isEmpty())
        rename->setResultPath(settings->value("Rename/resultPath").toString());
    else
        rename->setResultPath("C:/");

    this->close();
    rename->show();
}

void MainWindow::showMenuFromRename()
{
    this->show();
    settings->setValue("Rename/dirPath", rename->getDirectoryPath());
    settings->setValue("Rename/filePath", rename->getFilePath());
    settings->setValue("Rename/resultPath", rename->getResultPath());
}

void MainWindow::showMenuFromMetaData()
{
    this->show();
    settings->setValue("MetaData/inputPath", metaData->getInputPath());
}
