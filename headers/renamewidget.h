#ifndef RENAMEWIDGET_H
#define RENAMEWIDGET_H

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
    bool    report = true;

    Ui::RenameWidget *ui;
    Worker worker;
    QThread thread;
    void setStatusButton();
    void changeDir(int ind, QString path);

public:
    RenameWidget(QWidget *parent = nullptr);
    ~RenameWidget();
    QString getFilePath()     {return filePath;}
    QString getDirectoryPath(){return directoryPath;}
    QString getResultPath()   {return resultPath;}

    void setFilePath(QString _filePath)           {filePath = _filePath;}
    void setDirectoryPath(QString _directoryPath) { directoryPath = _directoryPath;}
    void setResultPath(QString _resultPath)       {resultPath = _resultPath;}

signals:
    void menuWindow();
    void doWork(const int &size, const QString &directoryPath, const QString &resultPath, const bool &report);

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
#endif
