#include "renamewidget.h"
#include "ui_renamewidget.h"
#include <private/qzipwriter_p.h>

QString RenameWidget::filePath;
QString RenameWidget::directoryPath;
QString RenameWidget::resultPath;
bool    RenameWidget::report;

RenameWidget::RenameWidget(QWidget *parent) :
    QWidget(parent),
    ui(new Ui::RenameWidget)
{
    ui->setupUi(this);

    filePath      = "C:/Users/dvm10/Desktop/";
    directoryPath = "C:/Users/dvm10/Desktop/";
    resultPath    = "C:/Users/dvm10/Desktop/";
    report        =  true;

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
    filePath = QFileDialog::getOpenFileName(this, "Выбрать файл для экспорта", filePath, "(*.xlsx)");
    ui->lineEdit_File->setText(filePath);
    setStatusButton();
}

void RenameWidget::on_pb_Directory_clicked()
{
    directoryPath = QFileDialog::getExistingDirectory(this, "Выбрать файл", directoryPath, QFileDialog::ShowDirsOnly);
    ui->lineEdit_Directory->setText(directoryPath);
    setStatusButton();
}

void RenameWidget::on_pb_Result_clicked()
{
    resultPath = QFileDialog::getExistingDirectory(this, "Выбрать файл", resultPath, QFileDialog::ShowDirsOnly);
    ui->lineEdit_Result->setText(resultPath);
    setStatusButton();
}

void RenameWidget::setText(QString message)
{
    ui->lableLoad->setText(message);
}

void RenameWidget::on_pb_Rename_clicked()
{
    if (!filePath.isEmpty() && !directoryPath.isEmpty() && !resultPath.isEmpty())
    {
        report = ui->cb_Report->isChecked();
        thread.terminate();
        thread.start();
        ui->pb_Rename->setEnabled(false);
        ui->lineEdit_Directory->setText("");
        ui->lineEdit_File->setText("");
        ui->lineEdit_Result->setText("");
    }
}

void RenameWidget::setStatusButton()
{
    if (!ui->lineEdit_Result->text().isEmpty() && !ui->lineEdit_Directory->text().isEmpty() && !ui->lineEdit_File->text().isEmpty())
         ui->pb_Rename->setEnabled(true);
}

void Worker::doWork()
{
        emit workProgress("");

        std::unique_ptr<QAxObject> excel (new QAxObject( "Excel.Application", this));
        excel->dynamicCall("SetVisible(bool)","false");
        std::unique_ptr<QAxObject> workbooks (excel->querySubObject( "Workbooks"));
        std::unique_ptr<QAxObject> workbook  (workbooks->querySubObject( "Open(const QString&)", RenameWidget::getFilePath()));
        std::unique_ptr<QAxObject> sheets    (workbook->querySubObject( "Worksheets" ));
        std::unique_ptr<QAxObject> sheet     (sheets->querySubObject("Item( int )",1));


        int lastRow = 0;

        while(true)
        {
            lastRow++;
            {
                std::unique_ptr<QAxObject> cell (sheet->querySubObject( "Cells(int,int)", lastRow,1));
                if (cell->property("Value").toString() == "")
                    break;
            }
        }

        QDir dir;
        QDir dirtemp;
        dir.setPath(RenameWidget::getDirectoryPath());
        QStringList listDir = dir.entryList(QDir::Dirs | QDir::NoDot | QDir::NoDotDot);
        QString article;
        QStringList mistakes;

        for (int i = 0; i<listDir.size();i++)
        {
            dirtemp.setPath(RenameWidget::getDirectoryPath() + "/" + listDir.at(i));
            QStringList tempstr = dirtemp.entryList(QDir::Dirs | QDir::NoDot | QDir::NoDotDot);
            article = getArticle(tempstr.at(0), lastRow, sheet.get());

            if (!article.isEmpty())
            {

                QString tempresult = RenameWidget::getResultPath() + "/" + "books" + "/" + article;
                QString temppath = RenameWidget::getDirectoryPath() + "/" + listDir.at(i) + "/" + tempstr .at(0);
                dir.mkpath(tempresult);
                dir.setPath(temppath);
                foreach (QString f, dir.entryList(QDir::Files))
                    QFile::copy(temppath + "/" + f,  tempresult + "/" + f);

                tempresult = RenameWidget::getResultPath() + "/" + "image";
                temppath = RenameWidget::getDirectoryPath() + "/" + listDir.at(i);
                dir.mkpath(tempresult);
                dir.setPath(temppath);
                foreach (QString f, dir.entryList(QDir::Files))
                    QFile::copy(temppath + "/" + f,  tempresult + "/" + f);

                QThread::msleep(100);
                doArchive(RenameWidget::getResultPath() + "/" + "books" + "/" + article + "/", RenameWidget::getResultPath() + "/" + "books" + "/" + article + "/" + article + ".zip", article);
                emit workProgress("Обработано каталогов: " + QString::number(i+1) + "/" + QString::number(listDir.size()));
            }
            else
                mistakes << RenameWidget::getDirectoryPath() + "/" + listDir.at(i);
        }
        if (RenameWidget::getReport())
        {
            QFile file(RenameWidget::getResultPath() + "/" + "report.txt");
            if (file.open(QIODevice::WriteOnly))
            {
                QTextStream out(&file);
                if (mistakes.size())
                {
                    out << "Broken ISBN:" << endl;
                    for (int i = 0;i < mistakes.size(); i++)
                        out << mistakes.at(i) << endl;
                    file.close();
                }
                else
                    out << "No mistakes!" << endl;
            }
        }

        QThread::msleep(100);
        emit workProgress("Архивация...");
        this->doArchive(RenameWidget::getResultPath() + "/" + "image" + "/", RenameWidget::getResultPath() + "/" "image.zip", article);

        emit workProgress("Готово");

        workbook->dynamicCall("Save()");
        workbooks->dynamicCall("Close()");
        excel->dynamicCall("Quit()");
}

QString Worker::getArticle(QString isbn, int lastRow, QAxObject *sheet)
{
    std::unique_ptr<QAxObject>myCell;
    for (int i = 0; i<lastRow; i++)
    {
        myCell.reset(sheet->querySubObject( "Cells(int,int)", i+1,2));
        QString check = myCell->property("Value").toString();
        if (myCell->property("Value").toString() == isbn)
            return sheet->querySubObject( "Cells(int,int)", i+1,1)->property("Value").toString();
    }
    return "";
}

void Worker::doArchive(QString path, QString zippath, QString article)
{
    QZipWriter zip(zippath);

    zip.setCompressionPolicy(QZipWriter::AutoCompress);

    QDirIterator it(path, QDir::Files | QDir::Dirs | QDir::NoDotAndDotDot, QDirIterator::Subdirectories);
    while (it.hasNext())
    {

        QString file_path = it.next();
        if (!(it.fileInfo().baseName()== article))
        {
            if (it.fileInfo().isDir())
            {
                zip.setCreationPermissions(QFile::permissions(file_path));
                zip.addDirectory(file_path.remove(path));
            }
            else if (it.fileInfo().isFile())
            {
                QFile file(file_path);
                if (!file.open(QIODevice::ReadOnly))
                    continue;
                zip.setCreationPermissions(QFile::permissions(file_path));
                QByteArray ba = file.readAll();
                zip.addFile(file_path.remove(path), ba);
            }
        }
    }
    zip.close();
}
