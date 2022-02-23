#ifndef RENAMEWIDGET_H
#define RENAMEWIDGET_H

#include <QWidget>
#include <QtWidgets>
#include <QAxObject>

namespace Ui {
class RenameWidget;
}

class Worker : public QObject
{
    Q_OBJECT
    void doArchive(QString path, QString zippath);
public:
    Worker(QObject *_parent = 0):QObject(_parent){};

signals:
    void workProgress(QString message);
    void workEnd();
    void workStart();


public slots:
    void doWork(int size, QString directoryPath, QString resultPath, bool report);
};


class RenameWidget : public QWidget
{
    Q_OBJECT
    QString filePath;
    QString directoryPath;
    QString resultPath;
    bool    report;

    Ui::RenameWidget *ui;
    Worker worker;
    QThread thread;
    void setStatusButton();

public:
    RenameWidget(QWidget *parent = nullptr);
    ~RenameWidget();

signals:
    void menuWindow();
    void doWork(int size, QString directoryPath, QString resultPath, bool report);

private slots:
    void on_pb_Close_clicked();
    void on_pb_Rename_clicked();
    void on_pb_File_clicked();
    void on_pb_Directory_clicked();
    void on_pb_Result_clicked();
    void setText(QString message);
    void workEnd();
    void workStart();
    void closeEvent(QCloseEvent *event) override;

};
#endif // RENAMEWIDGET_H
