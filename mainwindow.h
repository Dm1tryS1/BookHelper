#ifndef MAINWINDOW_H
#define MAINWINDOW_H

#include <QMainWindow>
#include <QtWidgets>
#include <QMediaPlayer>
#include <QAxObject>

QT_BEGIN_NAMESPACE
namespace Ui { class MainWindow; }
QT_END_NAMESPACE

class MainWindow : public QMainWindow
{
    Q_OBJECT
    QString InputPath =":/";
    QStringList listFile;
    QMediaPlayer *player;

    int ind = 0;
    int amount = 0;

    QAxObject* excel;

public:
    MainWindow(QWidget *parent = nullptr);
    ~MainWindow();

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

private:
    Ui::MainWindow *ui;
};

#endif // MAINWINDOW_H
