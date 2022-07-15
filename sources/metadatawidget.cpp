#include "headers/metadatawidget.h"
#include "ui_metadatawidget.h"

MetaDataWidget::MetaDataWidget(QWidget *parent) : QWidget(parent), ui(new Ui::MetaDataWidget)
{
    ui->setupUi(this);
    player = new QMediaPlayer(this);

    ui->tableWidget->setEditTriggers(QAbstractItemView::NoEditTriggers);
    ui->tableWidget->horizontalHeader()->setDefaultSectionSize(30);
    ui->tableWidget->verticalHeader()->setDefaultSectionSize(15);
    ui->tableWidget->horizontalHeader()->setSectionResizeMode(QHeaderView::Stretch);

    QStringList listColumn;
    listColumn << QString("Имя файла")
               << QString("Название")
               << QString("Субтитры")
               << QString("Комментарий")
               << QString("Участвующие исполнители")
               << QString("Исполнитель альбома")
               << QString("Альбом")
               << QString("Год")
               << QString("№")
               << QString("Жанр")
               << QString("Продолжительность")
               << QString("Общая продолжительность")
               << QString("Композитор")
               << QString("Дирижер")
               << QString("Настроение")
               << QString("Скорость потока")
               << QString("Кодек")
               << QString("Количество канало")
               << QString("Частота дискретизации");


    ui->tableWidget->setColumnCount(listColumn.size());
    ui->tableWidget->setHorizontalHeaderLabels(listColumn);

    connect(player, SIGNAL(mediaStatusChanged(QMediaPlayer::MediaStatus)), this, SLOT(onMediaStatusChanged(QMediaPlayer::MediaStatus)));

    ui->progressBar->setVisible(false);
}

MetaDataWidget::~MetaDataWidget()
{
    delete ui;
}

void MetaDataWidget::on_pb_Close_clicked()
{
    this->close();
    emit menuWindow();
}

void MetaDataWidget::makeList(QString path)
{
    QStringList filters;
    filters << "*.mp3";

    QDir dir;

    dir.setPath(path);
    QStringList listDir = dir.entryList(QDir::Dirs | QDir::NoDot | QDir::NoDotDot);

    if (listDir.empty())
    {
        QStringList files = dir.entryList(filters, QDir::Files);
        for (int i = 0; i<files.size();i++)
        {
            listFile << path + "/" + files.at(i);
            amount++;
        }
    }
    else
    {
        for (int i = 0; i<listDir.size();i++)
            makeList(path+"/"+listDir.at(i));
    }
}

void MetaDataWidget::on_pb_Load_clicked()
{
    amount = 0;
    listFile.clear();
    ind = 0;
    inputPath = QFileDialog::getExistingDirectory(this, "Выбрать каталок", inputPath, QFileDialog::ShowDirsOnly);
    if (!inputPath.isEmpty())
    {
            this->setEnabled(false);
            makeList(inputPath);
            ui->progressBar->setRange(0,amount);
            ui->progressBar->setValue(0);
            ui->progressBar->setVisible(true);
            player->setMedia(QUrl::fromLocalFile(listFile.at(0)));
    }
    else return;
}

QString MetaDataWidget::getDuration(QString duration)
{
    int seconds  = duration.toInt() / 1000;
    int minutes  = seconds / 60;
    seconds      = seconds % 60;
    int hours    = minutes / 60;
    minutes      = minutes % 60;
    return QString::number(hours) + ":" + QString::number(minutes) + ":" + QString::number(seconds);
}


void MetaDataWidget::newRow(QStringList list)
{
    ui->tableWidget->setRowCount(ui->tableWidget->rowCount() + 1);
    for (int i = 0; i < list.size(); i++)
    {
        QTableWidgetItem* item = new QTableWidgetItem(list.at(i));
        ui->tableWidget->setItem(ui->tableWidget->rowCount()-1,i,item);

    }
}

QString getPath(QString path)
{
    return path.remove(path.lastIndexOf('/'), path.length()-path.lastIndexOf('/'));
}

void MetaDataWidget::getSumDuration()
{
    QString currBook = getPath(QString(ui->tableWidget->item(0,0)->text()));
    int index = 0;
    int sum = 0;

    for(int i=0; i<ui->tableWidget->rowCount();i++)
    {
        if (currBook != getPath(ui->tableWidget->item(i,0)->text()))
        {
            ui->tableWidget->item(index,11)->setText(getDuration(QString::number(sum)));
            sum = ui->tableWidget->item(i,10)->text().toInt();
            index = i;
            currBook = getPath(ui->tableWidget->item(i,0)->text());
            ui->tableWidget->item(index,10)->setText(getDuration(ui->tableWidget->item(i,10)->text()));
        }
        else
        {
            sum += ui->tableWidget->item(i,10)->text().toInt();
            ui->tableWidget->item(i,10)->setText(getDuration(ui->tableWidget->item(i,10)->text()));
        }
    }
    ui->tableWidget->item(index,11)->setText(getDuration(QString::number(sum)));
}

void MetaDataWidget::onMediaStatusChanged(QMediaPlayer::MediaStatus status)
{
    if (status == QMediaPlayer::LoadedMedia)
    {
        QStringList list;

        list << listFile.at(ind)
             << (player->metaData("Title")).toString()
             << (player->metaData("SubTitle")).toString()
             << (player->metaData("Comment")).toString()
             << (player->metaData("ContributingArtist")).toString()
             << (player->metaData("AlbumArtist")).toString()
             << (player->metaData("AlbumTitle")).toString()
             << (player->metaData("Year")).toString()
             << (player->metaData("TrackNumber")).toString()
             << (player->metaData("Genre")).toString()
             << (player->metaData("Duration").toString())
             << ""
             << (player->metaData("Composer")).toString()
             << (player->metaData("Conductor")).toString()
             << (player->metaData("Mood")).toString()
             << QString::number(((player->metaData("AudioBitRate")).toInt() / 1000))
             << (player->metaData("AudioCodec")).toString()
             << (player->metaData("ChannelCount")).toString()
             << (player->metaData("SampleRate")).toString();

        newRow(list);
        ind++;

        if (ind == listFile.size())
        {
            this->setEnabled(true);
            ui->progressBar->setVisible(false);
            getSumDuration();

        }
        else
        {
            player->setMedia(QUrl::fromLocalFile(listFile.at(ind)));
            ui->progressBar->setValue(ind);
        }
    }
}

void MetaDataWidget::setHeader(QAxObject *sheet)
{
            std::unique_ptr<QAxObject>cell(sheet->querySubObject( "Cells(int,int)", 1,1));
            cell->dynamicCall( "SetValue(const QVariant&)", "Путь к файлу");

            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,2));
            cell->dynamicCall( "SetValue(const QVariant&)", "Название");


            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,3));
            cell->dynamicCall( "SetValue(const QVariant&)", "Субтитры");

            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,4));
            cell->dynamicCall( "SetValue(const QVariant&)", "Комментарий");

            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,5));
            cell->dynamicCall( "SetValue(const QVariant&)", "Участвующие исполнители");


            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,6));
            cell->dynamicCall( "SetValue(const QVariant&)", "Исполнитель альбома");

            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,7));
            cell->dynamicCall( "SetValue(const QVariant&)", "Альбом");


            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,8));
            cell->dynamicCall( "SetValue(const QVariant&)", "Год");


            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,9));
            cell->dynamicCall( "SetValue(const QVariant&)", "№");


            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,10));
            cell->dynamicCall( "SetValue(const QVariant&)", "Жанр");

            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,11));
            cell->dynamicCall( "SetValue(const QVariant&)", "Продолжительность");

            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,12));
            cell->dynamicCall( "SetValue(const QVariant&)", "Общая продолжительность");


            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,13));
            cell->dynamicCall( "SetValue(const QVariant&)", "Композитор");


            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,14));
            cell->dynamicCall( "SetValue(const QVariant&)", "Дирижер");


            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,15));
            cell->dynamicCall( "SetValue(const QVariant&)", "Настроение");

            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,16));
            cell->dynamicCall( "SetValue(const QVariant&)", "Скорость потока");

            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,17));
            cell->dynamicCall( "SetValue(const QVariant&)", "Кодек");

            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,18));
            cell->dynamicCall( "SetValue(const QVariant&)", "Количество каналов");

            cell.reset(sheet->querySubObject( "Cells(int,int)", 1,19));
            cell->dynamicCall( "SetValue(const QVariant&)", "Частота дискретизации");
}

bool MetaDataWidget::setColor(QAxObject *sheet, int lastRow)
{
    std::unique_ptr<QAxObject>cell(sheet->querySubObject( "Cells(int,int)", lastRow-1,1));
    std::unique_ptr<QAxObject>interior(cell->querySubObject("Interior"));
    QString color = interior->property("Color").toString();

    if (color == "16777215")
        return true;
    else
        return false;
}

void MetaDataWidget::on_pb_Export_clicked()
{
        QString fileName = QFileDialog::getOpenFileName(this, "Выбрать файл для экспорта", inputPath, "(*.xlsx)");
        if (!fileName.isEmpty())
        {
            ui->progressBar->setRange(0,ui->tableWidget->rowCount());
            ui->progressBar->setValue(0);
            ui->progressBar->setVisible(true);

            this->setEnabled(false);

            int lastRow = 0;
            bool lastRowFound = false;
            QString currBook;
            bool flag;

            QAxObject *excel = new QAxObject( "Excel.Application", this);
            excel->dynamicCall("SetVisible(bool)","false");
            excel-> setProperty ("DisplayAlerts", false);
            QAxObject *workbooks = excel->querySubObject( "Workbooks" );
            QAxObject *workbook = workbooks->querySubObject( "Open(const QString&)", fileName);
            QAxObject *sheets = workbook->querySubObject( "Worksheets" );
            QAxObject *sheet = sheets->querySubObject("Item( int )",1);

            while(!lastRowFound)
            {
                lastRow++;
                QAxObject *cell = sheet->querySubObject( "Cells(int,int)", lastRow,1);
                if (cell->property("Value").toString() == "")
                    lastRowFound = true;
                delete cell;
            }

            if (lastRow == 1)
            {
                    setHeader(sheet);
                    lastRow++;
            }

            QAxObject *cell = sheet->querySubObject( "Cells(int,int)", lastRow-1,1);
            currBook = cell->property("Value()").toString();
            flag = setColor(sheet, lastRow);
            delete cell;

            for (int i=0;i<ui->tableWidget->rowCount();i++)
            {
                ui->progressBar->setValue(i);
                for (int j=0;j<ui->tableWidget->columnCount();j++)
                {
                    QAxObject *cell = sheet->querySubObject( "Cells(int,int)", lastRow+i,j+1);
                    QAxObject *interior = cell->querySubObject("Interior");

                    if (currBook == getPath(ui->tableWidget->item(i,0)->text()))
                    {
                        if (flag)
                            interior->setProperty("Color",QColor("lightgray"));
                        else
                            interior->setProperty("Color",QColor("white"));

                    }
                    else

                    {
                        currBook = getPath(ui->tableWidget->item(i,0)->text());
                        flag = !flag;

                        if (flag)
                            interior->setProperty("Color",QColor("lightgray"));
                        else
                            interior->setProperty("Color",QColor("white"));
                    }

                    cell->dynamicCall( "SetValue(const QVariant&)", QVariant((ui->tableWidget->item(i,j)->text())));
                    delete interior;
                    delete cell;
                }

            }
            workbook->dynamicCall("Save()");
            workbooks->dynamicCall("Close()");
            excel->dynamicCall("Quit()");

            ui->progressBar->setVisible(false);

            delete excel;

            this->setEnabled(true);
            QMessageBox msgBox;
            msgBox.setWindowFlags(Qt::WindowStaysOnTopHint);
            msgBox.setText("Выгрузка прошла успешно!");
            msgBox.exec();

        } else return;
}




void MetaDataWidget::on_pb_Clear_clicked()
{
    ui->tableWidget->setRowCount(0);
}

void MetaDataWidget::closeEvent(QCloseEvent *event)
{
    event->ignore();
    emit menuWindow();
    event->accept();
}

