#include "headers/renamewidget.h"
#include "ui_renamewidget.h"
#include <private/qzipwriter_p.h>

RenameWidget::RenameWidget(QWidget *parent) :QWidget(parent), ui(new Ui::RenameWidget)
{
    ui->setupUi(this);

    worker.moveToThread(&thread);

    connect(this, &RenameWidget::doWork, &worker, &Worker::doWork);
    connect(&worker, &Worker::workProgress, this, &RenameWidget::setText);

    connect(&worker, &Worker::workStart, this, &RenameWidget::workStart);
    connect(&worker, &Worker::workEnd, this, &RenameWidget::workEnd); 

    thread.start();
}

RenameWidget::~RenameWidget()
{
    delete ui;
}

void RenameWidget::on_pb_Close_clicked()
{
    emit menuWindow();
    this->close();
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

        QFile excelFile(resultPath + "/" + "excel.txt");
        excelFile.open(QIODevice::WriteOnly);
        QTextStream outExcel(&excelFile);

        int lastRow = 0;
        {
            std::unique_ptr<QAxObject>excel(new QAxObject( "Excel.Application", this));
            excel->dynamicCall("SetVisible(bool)","false");
            std::unique_ptr<QAxObject> workbooks (excel->querySubObject( "Workbooks"));
            std::unique_ptr<QAxObject> workbook  (workbooks->querySubObject( "Open(const QString&)", filePath));
            std::unique_ptr<QAxObject> sheets    (workbook->querySubObject( "Worksheets" ));
            std::unique_ptr<QAxObject> sheet     (sheets->querySubObject("Item( int )",1));

            while(true)
            {
                lastRow++;
                if (sheet->querySubObject( "Cells(int,int)", lastRow,1)->property("Value").toString() == "")
                    break;
                QString key = sheet->querySubObject( "Cells(int,int)", lastRow,2)->property("Value").toString();
                QString value = sheet->querySubObject( "Cells(int,int)", lastRow,1)->property("Value").toString();
                outExcel << key << " " << value << endl;
            }

            excelFile.close();
            workbook->dynamicCall("Save()");
            workbooks->dynamicCall("Close()");
            excel->dynamicCall("Quit()");
        }

        emit doWork(lastRow, directoryPath, resultPath, report);
    }
}

void RenameWidget::setStatusButton()
{
    if (!ui->lineEdit_Result->text().isEmpty() && !ui->lineEdit_Directory->text().isEmpty() && !ui->lineEdit_File->text().isEmpty())
         ui->pb_Rename->setEnabled(true);
}

void RenameWidget::workStart()
{
    this->setEnabled(false);
}

void RenameWidget::workEnd()
{
    this->setEnabled(true);
    ui->lineEdit_Directory->clear();
    ui->lineEdit_File->clear();
    ui->lineEdit_Result->clear();
}


void Worker::doWork(int size, QString directoryPath, QString resultPath, bool report)
{
        emit workStart();
        emit workProgress("");

        QStringList mistakes;
        QDir dir;
        dir.setPath(directoryPath);
        QStringList listDir = dir.entryList(QDir::Dirs | QDir::NoDot | QDir::NoDotDot);

        QFile excelFile(resultPath + "/" + "excel.txt");
        QTextStream outExcel(&excelFile);

        for (int i = 0; i<listDir.size();i++)
        {
            dir.setPath(directoryPath + "/" + listDir.at(i));

            QStringList tempstr = dir.entryList(QDir::Dirs | QDir::NoDot | QDir::NoDotDot);

            QString article = "";
            excelFile.open(QIODevice::ReadOnly);
            for (int j = 0; j<size-1; j++)
            {
                QString key;
                QString value;
                outExcel >> key >> value;
                if (key == tempstr.at(0))
                {
                    article = value;
                    break;
                }
            }
            excelFile.close();
            emit workProgress("Обработано каталогов: " + QString::number(i) + "/" + QString::number(listDir.size()));


            if (!article.isEmpty())
            {
                QString temppath = directoryPath + "/" + listDir.at(i) + "/" + tempstr .at(0);
                QString tempres  = resultPath + "/" + "archives" + "/" + article;

                dir.mkpath(tempres);
                doArchive(temppath + "/", tempres + "/" + article + ".zip");

                temppath = directoryPath + "/" + listDir.at(i);
                dir.setPath(temppath);

                foreach (QString f, dir.entryList(QDir::Files))   
                    QFile::copy(temppath + "/" + f,  tempres + "/" + article + ".jpg");

            }
            else
                mistakes << directoryPath + "/" + listDir.at(i);
        }

        if (report)
        {
            QFile file(resultPath + "/" + "report.txt");
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
        excelFile.remove();
        emit workProgress("Готово");
        emit workEnd();
}

void Worker::doArchive(QString path, QString zippath)
{
    QZipWriter zip(zippath);

    zip.setCompressionPolicy(QZipWriter::AutoCompress);

    QDirIterator it(path, QDir::Files);
    while (it.hasNext())
    {
        QString file_path = it.next();
        if (it.fileInfo().isFile())
        {
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

void RenameWidget::closeEvent(QCloseEvent *event)
{
    event->ignore();
    emit menuWindow();
    event->accept();
}
