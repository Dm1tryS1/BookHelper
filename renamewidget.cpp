#include "renamewidget.h"
#include "ui_renamewidget.h"
#include <private/qzipwriter_p.h>


RenameWidget::RenameWidget(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::RenameWidget)
{
    ui->setupUi(this);
    worker.moveToThread(&thread);

    connect(&thread, &QThread::started, &worker, &Worker::doWork);
    connect(&worker, &Worker::workProgress, this, &RenameWidget::setText);

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
    if (!ui->lineEdit_File->text().isEmpty())
        flag1 = true;
    if (flag1 && flag2 && flag3)
        ui->pb_Rename->setEnabled(true);
}

void RenameWidget::on_pb_Directory_clicked()
{
    DirectoryPath = QFileDialog::getExistingDirectory(this, "Выбрать файл", DirectoryPath, QFileDialog::ShowDirsOnly);
    ui->lineEdit_Directory->setText(DirectoryPath);
    if (!ui->lineEdit_Directory->text().isEmpty())
        flag2 = true;
    if (flag1 && flag2 && flag3)
        ui->pb_Rename->setEnabled(true);
}

void RenameWidget::on_pb_Result_clicked()
{
    ResultPath = QFileDialog::getExistingDirectory(this, "Выбрать файл", ResultPath, QFileDialog::ShowDirsOnly);
    ui->lineEdit_Result->setText(ResultPath);
    if (!ui->lineEdit_Result->text().isEmpty())
        flag3 = true;
    if (flag1 && flag2 && flag3)
        ui->pb_Rename->setEnabled(true);
}

void RenameWidget::setText(QString message)
{
    ui->lableLoad->setText(message);
}

void RenameWidget::on_pb_Rename_clicked()
{
    if (!FilePath.isEmpty() && !DirectoryPath.isEmpty() && !ResultPath.isEmpty())
    {
        ui->lableLoad->setText("");
        worker.setter(DirectoryPath,FilePath,ResultPath);
        thread.start();
        flag1 = false;
        flag2 = false;
        flag3 = false;
        ui->pb_Rename->setEnabled(false);
        ui->lineEdit_Directory->setText("");
        ui->lineEdit_File->setText("");
        ui->lineEdit_Result->setText("");
    }
}

void Worker::setter(QString DPath, QString FPath, QString RPath)
{
    DirectoryPath = DPath;
    FilePath = FPath;
    ResultPath = RPath;
}

void Worker::doWork()
{
    emit workProgress("");
    QAxObject* excel;
    excel = new QAxObject( "Excel.Application", this);
    excel->dynamicCall("SetVisible(bool)","false");
    QAxObject *workbooks = excel->querySubObject( "Workbooks" );
    QAxObject *workbook = workbooks->querySubObject( "Open(const QString&)", FilePath);
    QAxObject *sheets = workbook->querySubObject( "Worksheets" );
    QAxObject *sheet = sheets->querySubObject("Item( int )",1);

    int lastRow = 0;
    bool lastRowFound = false;

    while(!lastRowFound)
    {
        lastRow++;
        QAxObject *cell = sheet->querySubObject( "Cells(int,int)", lastRow,1);
        if (cell->property("Value").toString() == "")
            lastRowFound = true;
        delete cell;
    }

    QDir dir;
    dir.setPath(DirectoryPath);
    QStringList listDir = dir.entryList(QDir::Dirs | QDir::NoDot | QDir::NoDotDot);

    for (int i = 0; i<listDir.size();i++)
    {
        QString tempresult = ResultPath + "/" + "books" + "/" + getArticle(listDir.at(i), lastRow, sheet);
        QString temppath = DirectoryPath + "/" + listDir.at(i) + "/" + listDir.at(i);
        dir.mkpath(tempresult);

        dir.setPath(temppath);
        foreach (QString f, dir.entryList(QDir::Files))
            QFile::copy(temppath + "/" + f,  tempresult + "/" + f);

        tempresult = ResultPath + "/" + "image";
        temppath = DirectoryPath + "/" + listDir.at(i);
        dir.mkpath(tempresult);
        dir.setPath(temppath);
        foreach (QString f, dir.entryList(QDir::Files))
            QFile::copy(temppath + "/" + f,  tempresult + "/" + f);
        QThread::msleep(100);
        emit workProgress("Обработано каталогов: " + QString::number(i) + "/" + QString::number(listDir.size()));
    }
    QThread::msleep(100);
    emit workProgress("Архивация...");
    this->doArchive(ResultPath + "/" + "books" + "/", ResultPath + "/" "books.zip");
    this->doArchive(ResultPath + "/" + "image" + "/", ResultPath + "/" "image.zip");

    emit workProgress("Готово");

    workbook->dynamicCall("Save()");
    workbooks->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
    delete excel;
}

QString Worker::getArticle(QString isbn, int lastRow, QAxObject *sheet)
{
    QString result;
    QAxObject *cell;

    for (int i = 0; i<lastRow; i++)
    {
        cell = sheet->querySubObject( "Cells(int,int)", i+1,1);
        if (cell->property("Value").toString() == isbn)
        {
            result = sheet->querySubObject( "Cells(int,int)", i+1,2)->property("Value").toString();
            break;
        }

    }
    delete cell;
    return result;
}

void Worker::doArchive(QString path, QString zippath)
{
    {
           QZipWriter zip(zippath);

           zip.setCompressionPolicy(QZipWriter::AutoCompress);

           QDirIterator it(path, QDir::Files | QDir::Dirs | QDir::NoDotAndDotDot,
                           QDirIterator::Subdirectories);
           while (it.hasNext()) {
               QString file_path = it.next();
               if (it.fileInfo().isDir()) {
                   zip.setCreationPermissions(QFile::permissions(file_path));
                   zip.addDirectory(file_path.remove(path));
               } else if (it.fileInfo().isFile()) {
                   QFile file(file_path);
                   if (!file.open(QIODevice::ReadOnly))
                       continue;

                   zip.setCreationPermissions(QFile::permissions(file_path));
                   QByteArray ba = file.readAll();
                   zip.addFile(file_path.remove(path), ba);
               }
           }
           zip.close();
       }
}
