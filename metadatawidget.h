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
    QString InputPath =":/";
    QStringList listFile;
    QMediaPlayer *player;

    int ind = 0;
    int amount = 0;

    QAxObject* excel;

public:
    explicit MetaDataWidget(QWidget *parent = nullptr);
    ~MetaDataWidget();

signals:
    void menuWindow();

private slots:
    void on_pb_Load_clicked();
    void on_pb_Export_clicked();
    void on_pb_Clear_clicked();

    QString getDuration(QString duration);
    void newRow(QStringList list);
    void setHeader(QAxObject *sheet);
    bool setColor(QAxObject *sheet,int lastRow);
    void makeList(QString path);
    void onMediaStatusChanged(QMediaPlayer::MediaStatus status);
    void getSumDuration();

    void on_pb_Close_clicked();
private:
    Ui::MetaDataWidget *ui;
};

#endif // METADATAWIDGET_H
