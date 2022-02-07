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
public:
    Worker(QObject *_parent = 0):QObject(_parent){};

signals:
    void workProgress(QString message);

public slots:
    void doWork();

private:
    QString FilePath ="";
    QString DirectoryPath ="";
    QString ResultPath = "";

    QString getArticle(QString isbn, int lastRow, QAxObject *sheet);
    void doArchive(QString path, QString zippath);
public:
    void setter(QString DPath, QString FPath, QString RPath);
};


class RenameWidget : public QWidget
{
    Q_OBJECT
    QString FilePath ="C:/Users/dvm10/Desktop/";
    QString DirectoryPath ="C:/Users/dvm10/Desktop/";
    QString ResultPath = "C:/Users/dvm10/Desktop/";

public:
    explicit RenameWidget(QWidget *parent = nullptr);
    ~RenameWidget();

signals:
    void menuWindow();

private slots:
    void on_pb_Close_clicked();
    void on_pb_Rename_clicked();
    void on_pb_File_clicked();
    void on_pb_Directory_clicked();
    void on_pb_Result_clicked();
    void setText(QString message);

public:
    Ui::RenameWidget *ui;
    Worker worker;
    QThread thread;
    bool flag1 = false;
    bool flag2 = false;
    bool flag3 = false;
};
#endif // RENAMEWIDGET_H
