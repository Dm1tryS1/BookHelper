#include "renamewidget.h"
#include "ui_renamewidget.h"

RenameWidget::RenameWidget(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::RenameWidget)
{
    ui->setupUi(this);
}

RenameWidget::~RenameWidget()
{
    delete ui;
}

void RenameWidget::on_pb_Close_clicked()
{
    this->close();
    emit menuWindow();
}

void RenameWidget::on_pb_File_clicked()
{
    FilePath = QFileDialog::getOpenFileName(this, "Выбрать файл для экспорта", FilePath, "(*.xlsx)");
    ui->lineEdit_File->setText(FilePath);
}

void RenameWidget::on_pb_Directory_clicked()
{
    DirectoryPath = QFileDialog::getExistingDirectory(this, "Выбрать файл", DirectoryPath, QFileDialog::ShowDirsOnly);
    ui->lineEdit_Directory->setText(DirectoryPath);
}

void RenameWidget::on_pb_Rename_clicked()
{

}
