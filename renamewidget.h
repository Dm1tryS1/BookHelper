#ifndef RENAMEWIDGET_H
#define RENAMEWIDGET_H

#include <QWidget>
#include <QtWidgets>
#include <QAxObject>

namespace Ui {
class RenameWidget;
}

class RenameWidget : public QWidget
{
    Q_OBJECT
    QString FilePath =":/";
    QString DirectoryPath =":/";

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

private:
    Ui::RenameWidget *ui;
};

#endif // RENAMEWIDGET_H
