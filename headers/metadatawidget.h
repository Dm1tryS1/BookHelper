#ifndef METADATAWIDGET_H
#define METADATAWIDGET_H

#include <QWidget>
#include <QtWidgets>
#include <QMediaPlayer>
#include <QAxObject>

namespace Ui {
class MetaDataWidget;
}

class MetaDataWidget : public QWidget
{
    Q_OBJECT
    QString InputPath ="C:/Users/dvm10/Desktop/";
    QStringList listFile;
    QMediaPlayer *player;
    int ind;
    int amount;
    Ui::MetaDataWidget *ui;

    void makeList(QString path);
    void newRow(QStringList list);

    QString getDuration(QString duration);
    void getSumDuration();

    void setHeader(QAxObject *sheet);
    bool setColor(QAxObject *sheet,int lastRow);

public:
    explicit MetaDataWidget(QWidget *parent = nullptr);
    ~MetaDataWidget();

signals:
    void menuWindow();

private slots:
    void on_pb_Load_clicked();
    void on_pb_Export_clicked();
    void on_pb_Clear_clicked();
    void on_pb_Close_clicked();
    void onMediaStatusChanged(QMediaPlayer::MediaStatus status);
    void closeEvent(QCloseEvent *event) override;
};

#endif // METADATAWIDGET_H
