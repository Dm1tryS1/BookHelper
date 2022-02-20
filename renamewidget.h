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
    QString getArticle(QString isbn, int lastRow, QAxObject *sheet);
    void doArchive(QString path, QString zippath, QString article);

public:
    Worker(QObject *_parent = 0):QObject(_parent){};

signals:
    void workProgress(QString message);

public slots:
    void doWork();
};


class RenameWidget : public QWidget
{
    Q_OBJECT
    static QString filePath;
    static QString directoryPath;
    static QString resultPath;
    static bool    report;

    Ui::RenameWidget *ui;
    Worker worker;
    QThread thread;
    void setStatusButton();

public:
    explicit RenameWidget(QWidget *parent = nullptr);
    ~RenameWidget();
    static QString getFilePath() { return filePath; }
    static QString getDirectoryPath() { return directoryPath; }
    static QString getResultPath() { return resultPath; }
    static bool    getReport() { return report; }


signals:
    void menuWindow();

private slots:
    void on_pb_Close_clicked();
    void on_pb_Rename_clicked();
    void on_pb_File_clicked();
    void on_pb_Directory_clicked();
    void on_pb_Result_clicked();
    void setText(QString message);

};
#endif // RENAMEWIDGET_H
