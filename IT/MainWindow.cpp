// MainWindow.cpp

#include <QtGlobal>
#include <QString>
#include <QFileDialog>
#include <QMessageBox>
#include <QVBoxLayout>
#include <QPrinter>
#include <QPrintPreviewDialog>
#include <QPainter>
#include <QDebug>
#include <QVector>
#include <QList>
#include <QSharedPointer>
#include <QInputDialog>
#include <QStringList>
#include <QVarLengthArray>
#include <QDateTime>

#include "xlsxcelllocation.h"
#include "xlsxcell.h"
#include "CopycatTableModel.h"
#include "tableprinter.h"
#include "MainWindow.h"
#include "ui_MainWindow.h"

MainWindow::MainWindow(QWidget *parent) :
    QMainWindow(parent),
    ui(new Ui::MainWindow)
{
    xlsxDoc = nullptr;
    tabWidget = nullptr;

    ui->setupUi(this);

    tabWidget = new QTabWidget(this);
    setCentralWidget(tabWidget);

    this->setWindowTitle(QString("LGE IT Tool"));

    setWindowIcon(QIcon(":/excel.ico"));

    if(!createDatabase())
    {
        QMessageBox msgBox;
        QString alertMsg = QString("Failed create database");
        msgBox.setIcon( QMessageBox::Critical );
        msgBox.setText( alertMsg );
        msgBox.exec();
    }
}

MainWindow::~MainWindow()
{
    delete ui;

    if ( nullptr != xlsxDoc )
    {
        delete xlsxDoc;
    }
}

void MainWindow::on_action_Quit_triggered()
{
    // quit
    this->close();
}

void MainWindow::on_action_Open_triggered()
{
    // open file dialog
    QString fileName = QFileDialog::getOpenFileName(this,
        tr("Open Excel"), ".", tr("Excel Files (*.xlsx)"));

    if ( ! loadXlsx(fileName) ) // load xlsx
    {
        QMessageBox msgBox;
        QString alertMsg = QString("Failed to load file.\n %1").arg(fileName);
        msgBox.setIcon( QMessageBox::Critical );
        msgBox.setText( alertMsg );
        msgBox.exec();
        return;
    }

    this->setWindowTitle(fileName);

}

bool MainWindow::loadXlsx(QString fileName)
{
    // tried to load xlsx using temporary document
    QXlsx::Document xlsxTmp( fileName );
    if ( !xlsxTmp.isLoadPackage() )
    {
        return false; // failed to load
    }

    // clear xlsxDoc
    if ( nullptr != xlsxDoc )
    {
        delete xlsxDoc;
        xlsxDoc = nullptr;
    }

    // load new xlsx using new document
    xlsxDoc = new QXlsx::Document( fileName );
    xlsxDoc->isLoadPackage();

    // clear tab widget
    tabWidget->clear();
    // Removes all the pages, but does not delete them.
    // Calling this function is equivalent to calling removeTab()
    // until the tab widget is empty.

    // clear sub-items of every tab
    foreach ( XlsxTab* ptrTab, xlsxTabList )
    {
        if ( nullptr == ptrTab )
            continue;
        delete ptrTab;
    }
    xlsxTabList.clear();

    int sheetIndexNumber = 0;
    int activeSheetNumber = -1;

    AbstractSheet* activeSheet = xlsxDoc->workbook()->activeSheet();
    // NOTICE: active sheet is lastest sheet. It's not focused sheet.

    foreach( QString curretnSheetName, xlsxDoc->sheetNames() )
    {
        QXlsx::AbstractSheet* currentSheet = xlsxDoc->sheet( curretnSheetName );
        if ( nullptr == currentSheet )
            continue;

        if ( activeSheet == currentSheet )
        {
            // current sheet is active sheet.
            activeSheetNumber = sheetIndexNumber;
        }

        XlsxTab* newSheet = new XlsxTab( this, xlsxDoc, currentSheet, sheetIndexNumber ); // create new tab
        xlsxTabList.push_back( newSheet ); // append to xlsx pointer list
        tabWidget->addTab( newSheet, curretnSheetName  ); // add tab widget
        sheetIndexNumber++; // increase sheet index number
    }

    if ( (-1) != activeSheetNumber ){
        qDebug() <<" activeSheetNumber=" << activeSheetNumber;
        tabWidget->setCurrentIndex(activeSheetNumber);
    }

    return true;
}

void MainWindow::on_action_About_triggered()
{
    QString text ="LG RnD Viet Nam Hai Phong" ;

    QMessageBox::about(this, "LGE IT Tool", text);
}

void MainWindow::on_action_New_triggered()
{
    // TODO: new document
    QMessageBox msgBox;
    msgBox.setWindowTitle("LGE IT");
    msgBox.setText( "Do you want delete all data in database ?" );
    msgBox.setStandardButtons(QMessageBox::Yes);
    msgBox.addButton(QMessageBox::No);
    msgBox.setDefaultButton(QMessageBox::No);
    if(msgBox.exec() == QMessageBox::Yes){
    QSqlQuery query;
    query.exec(QString("DROP TABLE IF EXISTS RawLog;"));
    query.exec(QString("DROP TABLE IF EXISTS VT1;"));
    query.exec(QString("DROP TABLE IF EXISTS VT2;"));
    query.exec(QString("DROP TABLE IF EXISTS SD1;"));
    query.exec(QString("DROP TABLE IF EXISTS SD2;"));
    query.exec(QString("DROP TABLE IF EXISTS SD3;"));
    // recreate table
    createTables();
  }
}
bool MainWindow::createDatabase()
{
    QSqlDatabase db = QSqlDatabase::addDatabase("QSQLITE");
    db.setDatabaseName(qApp->applicationDirPath()
                        + QDir::separator()
                        + "LGEIT.sqlite" );
    if (!db.open())
    {
            QMessageBox::critical(nullptr, QObject::tr("Cannot open database"),
                QObject::tr("Unable to establish a database connection.\n"
                            "This example needs SQLite support. Please read "
                            "the Qt SQL driver documentation for information how "
                            "to build it.\n\n"
                            "Click Cancel to exit."), QMessageBox::Cancel);
            return false;
    }
    return true;
}
void MainWindow::createTables()
{
    QSqlQuery query;
    query.exec("create table RawLog "
              "(id integer primary key autoincrement unique not null, "
              "securityzone varchar(20), "
              "buildingname varchar(30), "
              "buildingfloor varchar(30), "
              "doorname varchar(30), "
              "classification varchar(5), "
              "company varchar(30), "
              "empno varchar(30), "
              "username varchar(30), "
              "designation varchar(50), "
              "department varchar(50), "
              "usertype varchar(30), "
              "entrancedate varchar(30), "
              "entrancetime varchar(30), "
              "entranceresult varchar(30), "
              "card varchar(30), "
              "mifare varchar(30), "
              "cardnumber varchar(30), "
              "email varchar(50) )");

    query.exec("create table VT1 "
              "(id integer primary key autoincrement unique not null, "
              "empno varchar(30), "
              "username varchar(30), "
              "department varchar(50), "
              "entrancedate varchar(30), "
              "entrancetimein varchar(30), "
              "entrancetimeout varchar(30), "
              "reason varchar(150))");

    query.exec("create table VT2 "
              "(id integer primary key autoincrement unique not null, "
              "empno varchar(30), "
              "username varchar(30), "
              "department varchar(50), "
              "entrancedate varchar(30), "
              "entrancetimein varchar(30), "
              "entrancetimeout varchar(30), "
              "reason varchar(150))");

    query.exec("create table SD1 "
              "(id integer primary key autoincrement unique not null, "
              "empno varchar(30), "
              "username varchar(30), "
              "department varchar(50), "
              "entrancedate varchar(30), "
              "entrancetimein varchar(30), "
              "entrancetimeout varchar(30), "
              "reason varchar(150))");

    query.exec("create table SD2 "
              "(id integer primary key autoincrement unique not null, "
              "empno varchar(30), "
              "username varchar(30), "
              "department varchar(50), "
              "entrancedate varchar(30), "
              "entrancetimein varchar(30), "
              "entrancetimeout varchar(30), "
              "reason varchar(150))");

    query.exec("create table SD3 "
              "(id integer primary key autoincrement unique not null, "
              "empno varchar(30), "
              "username varchar(30), "
              "department varchar(50), "
              "entrancedate varchar(30), "
              "entrancetimein varchar(30), "
              "entrancetimeout varchar(30), "
              "reason varchar(150))");

}
void MainWindow::on_action_Save_triggered()
{
    if ( nullptr == xlsxDoc )
    {
        QMessageBox msgBox;
        msgBox.setText( "xlsx document is empty." );
        msgBox.setIcon( QMessageBox::Warning );
        msgBox.exec();
        return;
    }

    Worksheet* wsheet = (Worksheet*) xlsxDoc->workbook()->activeSheet();
    if ( nullptr == wsheet )
    {
        QMessageBox msgBox;
        msgBox.setText( "worksheet is set." );
        msgBox.setIcon( QMessageBox::Warning );
        msgBox.exec();
        return;
    }

    QStringList items;
    foreach( QString curretnSheetName, xlsxDoc->sheetNames() )
    {
        items << curretnSheetName;
    }

    bool ok = false;
    QString selectedItem = QInputDialog::getItem( this, tr("Save Database"), tr("Select sheet:"), items, 0, false, &ok );

    if (!ok)
    {
        QMessageBox msgBox;
        msgBox.setText( "Please select one sheet for save database!" );
        msgBox.setIcon( QMessageBox::Warning );
        msgBox.exec();

        return;
    }

    int idx = 0;
    foreach (const QString& varList, items)
    {
        if ( varList == selectedItem )
        {
            break;
        }
        idx++;
    }
    qDebug() <<" setActiveSheet index=" << idx;
    xlsxDoc->workbook()->setActiveSheet( idx );
    wsheet = (Worksheet*) xlsxDoc->workbook()->activeSheet(); // set active sheet

    int maxRow = -1;
    int maxCol = -1;
    QVector<CellLocation> clList = wsheet->getFullCells( &maxRow, &maxCol );

    // Fixed for Visual C++, cause VC++ does not support C99.
    // https://docs.microsoft.com/en-us/cpp/c-language/ansi-conformance?view=vs-2017
    // QString arr[maxRow][maxCol];

    QVector <QVector <QString> > arr;

    for (int i = 0; i < maxRow; i++)
    {
        QVector<QString> tempVector;
        for (int j = 0; j < maxCol; j++)
        {
            tempVector.push_back(QString(""));
        }
        arr.push_back(tempVector);
    }

    for ( int ic = 0; ic < clList.size(); ++ic )
    {
        CellLocation cl = clList.at(ic);

        int row = cl.row - 1;
        int col = cl.col - 1;

        QSharedPointer<Cell> ptrCell = cl.cell; // cell pointer

        QString strValue = ptrCell->value().toString();

        arr[row][col] = strValue;
    }
    if(selectedItem == "VT1")
    {
        for (int ir = 3 ; ir < maxRow; ir++)
        {
            QSqlQuery query;
            query.prepare("INSERT INTO VT1 (empno, username, department,\
                            entrancedate, entrancetimein, entrancetimeout, reason) "
                         "VALUES (:empno, :username, :department,\
                                  :entrancedate, :entrancetimein, :entrancetimeout, :reason)");
            //query.bindValue(":id", ir);
            for (int ic = 0 ; ic < maxCol; ic++)
            {
                QString strValue = arr[ir][ic];
                if ( strValue.isNull() )
                {
                    strValue = QString("");
                }
                //qDebug() <<" strValue=" << strValue << "ic" << ic;

                switch(ic)
                {
                case 0:
                    query.bindValue(":empno", strValue);
                    break;
                case 1:
                    query.bindValue(":username", strValue);
                    break;
                case 2:
                    query.bindValue(":department", strValue);
                    break;
                case 3:
                    {
                    //QDate date = QDate::fromString(strValue,"yyyy-MM-dd");
                    query.bindValue(":entrancedate", strValue);
                    }
                    break;
                case 4:
                    query.bindValue(":entrancetimein", strValue);
                    break;
                case 5:
                    query.bindValue(":entrancetimeout", strValue);
                    break;
                case 6:
                    query.bindValue(":reason", strValue);
                    break;
                }
            }
            if( !query.exec() )
            {
              qDebug() << query.lastError().text();
            }
            //qDebug() <<"\n";
        }
    }
    if(selectedItem == "VT2")
    {
        qDebug() <<" VT2";
        for (int ir = 3 ; ir < maxRow; ir++)
        {
            QSqlQuery query;
            query.prepare("INSERT INTO VT2 (empno, username, department,\
                            entrancedate, entrancetimein, entrancetimeout, reason) "
                         "VALUES (:empno, :username, :department,\
                                  :entrancedate, :entrancetimein, :entrancetimeout, :reason)");
            //query.bindValue(":id", ir);
            for (int ic = 0 ; ic < maxCol; ic++)
            {
                QString strValue = arr[ir][ic];
                if ( strValue.isNull() )
                {
                    strValue = QString("");
                }
                //qDebug() <<" strValue=" << strValue << "ic" << ic;

                switch(ic)
                {
                case 0:
                    query.bindValue(":empno", strValue);
                    break;
                case 1:
                    query.bindValue(":username", strValue);
                    break;
                case 2:
                    query.bindValue(":department", strValue);
                    break;
                case 3:
                    query.bindValue(":entrancedate", strValue);
                    break;
                case 4:
                    query.bindValue(":entrancetimein", strValue);
                    break;
                case 5:
                    query.bindValue(":entrancetimeout", strValue);
                    break;
                case 6:
                    query.bindValue(":reason", strValue);
                    break;
                }
            }
            if( !query.exec() )
            {
              qDebug() << query.lastError().text();
            }
            //qDebug() <<"\n";
        }
    }
    if(selectedItem == "SD1")
    {
        qDebug() <<" SD1";
        for (int ir = 3 ; ir < maxRow; ir++)
        {
            QSqlQuery query;
            query.prepare("INSERT INTO SD1 (empno, username, department,\
                            entrancedate, entrancetimein, entrancetimeout, reason) "
                         "VALUES (:empno, :username, :department,\
                                  :entrancedate, :entrancetimein, :entrancetimeout, :reason)");
            //query.bindValue(":id", ir);
            for (int ic = 0 ; ic < maxCol; ic++)
            {
                QString strValue = arr[ir][ic];
                if ( strValue.isNull() )
                {
                    strValue = QString("");
                }
                //qDebug() <<" strValue=" << strValue << "ic" << ic;

                switch(ic)
                {
                case 0:
                    query.bindValue(":empno", strValue);
                    break;
                case 1:
                    query.bindValue(":username", strValue);
                    break;
                case 2:
                    query.bindValue(":department", strValue);
                    break;
                case 3:
                    query.bindValue(":entrancedate", strValue);
                    break;
                case 4:
                    query.bindValue(":entrancetimein", strValue);
                    break;
                case 5:
                    query.bindValue(":entrancetimeout", strValue);
                    break;
                case 6:
                    query.bindValue(":reason", strValue);
                    break;
                }
            }
            if( !query.exec() )
            {
              qDebug() << query.lastError().text();
            }
            //qDebug() <<"\n";
        }
    }
    if(selectedItem == "SD2")
    {
        qDebug() <<" SD2";
        for (int ir = 3 ; ir < maxRow; ir++)
        {
            QSqlQuery query;
            query.prepare("INSERT INTO SD2 (empno, username, department,\
                            entrancedate, entrancetimein, entrancetimeout, reason) "
                         "VALUES (:empno, :username, :department,\
                                  :entrancedate, :entrancetimein, :entrancetimeout, :reason)");
            //query.bindValue(":id", ir);
            for (int ic = 0 ; ic < maxCol; ic++)
            {
                QString strValue = arr[ir][ic];
                if ( strValue.isNull() )
                {
                    strValue = QString("");
                }
                //qDebug() <<" strValue=" << strValue << "ic" << ic;

                switch(ic)
                {
                case 0:
                    query.bindValue(":empno", strValue);
                    break;
                case 1:
                    query.bindValue(":username", strValue);
                    break;
                case 2:
                    query.bindValue(":department", strValue);
                    break;
                case 3:
                    query.bindValue(":entrancedate", strValue);
                    break;
                case 4:
                    query.bindValue(":entrancetimein", strValue);
                    break;
                case 5:
                    query.bindValue(":entrancetimeout", strValue);
                    break;
                case 6:
                    query.bindValue(":reason", strValue);
                    break;
                }
            }
            if( !query.exec() )
            {
              qDebug() << query.lastError().text();
            }
            //qDebug() <<"\n";
        }
    }
    if(selectedItem == "SD3")
    {
        qDebug() <<" SD3";
        for (int ir = 3 ; ir < maxRow; ir++)
        {
            QSqlQuery query;
            query.prepare("INSERT INTO SD3 (empno, username, department,\
                            entrancedate, entrancetimein, entrancetimeout, reason) "
                         "VALUES (:empno, :username, :department,\
                                  :entrancedate, :entrancetimein, :entrancetimeout, :reason)");
            //query.bindValue(":id", ir);
            for (int ic = 0 ; ic < maxCol; ic++)
            {
                QString strValue = arr[ir][ic];
                if ( strValue.isNull() )
                {
                    strValue = QString("");
                }
                //qDebug() <<" strValue=" << strValue << "ic" << ic;

                switch(ic)
                {
                case 0:
                    query.bindValue(":empno", strValue);
                    break;
                case 1:
                    query.bindValue(":username", strValue);
                    break;
                case 2:
                    query.bindValue(":department", strValue);
                    break;
                case 3:
                    query.bindValue(":entrancedate", strValue);
                    break;
                case 4:
                    query.bindValue(":entrancetimein", strValue);
                    break;
                case 5:
                    query.bindValue(":entrancetimeout", strValue);
                    break;
                case 6:
                    query.bindValue(":reason", strValue);
                    break;
                }
            }
            if( !query.exec() )
            {
              qDebug() << query.lastError().text();
            }
            //qDebug() <<"\n";
        }
    }
    else if(selectedItem != "VT1" && selectedItem != "VT2" && selectedItem != "SD1"&& \
            selectedItem != "SD2" && selectedItem != "SD3")
    {
        qDebug() <<" RawLog";
        for (int ir = 1 ; ir < maxRow; ir++)
        {
            QSqlQuery query;
            query.prepare("INSERT INTO RawLog (securityzone, buildingname, buildingfloor, doorname,\
                            classification, company, empno, username, designation, department, usertype,\
                            entrancedate, entrancetime, entranceresult, card, mifare, cardnumber, email) "
                         "VALUES (:securityzone, :buildingname, :buildingfloor, :doorname,\
                                  :classification, :company, :empno, :username, :designation, :department, :usertype,\
                                  :entrancedate, :entrancetime, :entranceresult, :card, :mifare, :cardnumber, :email)");
            //query.bindValue(":id", ir);
            for (int ic = 0 ; ic < maxCol; ic++)
            {
                QString strValue = arr[ir][ic];
                if ( strValue.isNull() )
                {
                    strValue = QString("");
                }
                //qDebug() <<" strValue=" << strValue << "ic" << ic;

                switch(ic)
                {
                case 0:
                    query.bindValue(":securityzone", strValue);
                    break;
                case 1:
                    query.bindValue(":buildingname", strValue);
                    break;
                case 2:
                    query.bindValue(":buildingfloor", strValue);
                    break;
                case 3:
                    query.bindValue(":doorname", strValue);
                    break;
                case 4:
                    query.bindValue(":classification", strValue);
                    break;
                case 5:
                    query.bindValue(":company", strValue);
                    break;
                case 6:
                    query.bindValue(":empno", strValue);
                    break;
                case 7:
                    query.bindValue(":username", strValue);
                    break;
                case 8:
                    query.bindValue(":designation", strValue);
                    break;
                case 9:
                    query.bindValue(":department", strValue);
                    break;
                case 10:
                    query.bindValue(":usertype", strValue);
                    break;
                case 11:
                    query.bindValue(":entrancedate", strValue);
                    break;
                case 12:
                    query.bindValue(":entrancetime", strValue);
                    break;
                case 13:
                    query.bindValue(":entranceresult", strValue);
                    break;
                case 14:
                    query.bindValue(":card", strValue);
                    break;
                case 15:
                    query.bindValue(":mifare", strValue);
                    break;
                case 16:
                    query.bindValue(":cardnumber", strValue);
                    break;
                case 17:
                    query.bindValue(":email", strValue);
                    break;
                }
            }
            if( !query.exec() )
            {
              qDebug() << query.lastError().text();
            }
            //qDebug() <<"\n";
        }
    }



}

void MainWindow::on_action_Print_triggered()
{
    QSqlQuery query;
    query.prepare("UPDATE VT1 \
                   SET entrancetimein = ( SELECT entrancetime FROM RawLog WHERE (RawLog.classification LIKE 'IN')\
                                                                                 AND (VT1.empno = RawLog.empno)\
                                                                                 AND (VT1.entrancedate = RawLog.entrancedate)\
                                                                                 ORDER BY RawLog.entrancetime ASC)");
    if( !query.exec() )
    {
        qDebug() << query.lastError().text();
    }

    query.prepare("UPDATE VT1 \
                   SET entrancetimeout = ( SELECT entrancetime FROM RawLog WHERE (RawLog.classification LIKE 'OUT')\
                                                                                 AND (VT1.empno = RawLog.empno)\
                                                                                 AND (VT1.entrancedate = RawLog.entrancedate)\
                                                                                 ORDER BY RawLog.entrancetime DESC)");
    if( !query.exec() )
    {
        qDebug() << query.lastError().text();
    }

    query.prepare("UPDATE VT2 \
                   SET entrancetimein = ( SELECT entrancetime FROM RawLog WHERE (RawLog.classification LIKE 'IN')\
                                                                                 AND (VT2.empno = RawLog.empno)\
                                                                                 AND (VT2.entrancedate = RawLog.entrancedate)\
                                                                                 ORDER BY RawLog.entrancetime ASC)");
    if( !query.exec() )
    {
        qDebug() << query.lastError().text();
    }

    query.prepare("UPDATE VT2 \
                   SET entrancetimeout = ( SELECT entrancetime FROM RawLog WHERE (RawLog.classification LIKE 'OUT')\
                                                                                 AND (VT2.empno = RawLog.empno)\
                                                                                 AND (VT2.entrancedate = RawLog.entrancedate)\
                                                                                 ORDER BY RawLog.entrancetime DESC)");
    if( !query.exec() )
    {
        qDebug() << query.lastError().text();
    }

    query.prepare("UPDATE SD1 \
                   SET entrancetimein = ( SELECT entrancetime FROM RawLog WHERE (RawLog.classification LIKE 'IN')\
                                                                                 AND (SD1.empno = RawLog.empno)\
                                                                                 AND (SD1.entrancedate = RawLog.entrancedate)\
                                                                                 ORDER BY RawLog.entrancetime ASC)");
    if( !query.exec() )
    {
        qDebug() << query.lastError().text();
    }

    query.prepare("UPDATE SD1 \
                   SET entrancetimeout = ( SELECT entrancetime FROM RawLog WHERE (RawLog.classification LIKE 'OUT')\
                                                                                 AND (SD1.empno = RawLog.empno)\
                                                                                 AND (SD1.entrancedate = RawLog.entrancedate)\
                                                                                 ORDER BY RawLog.entrancetime DESC)");
    if( !query.exec() )
    {
        qDebug() << query.lastError().text();
    }

    query.prepare("UPDATE SD2 \
           SET entrancetimein = ( SELECT entrancetime FROM RawLog WHERE (RawLog.classification LIKE 'IN')\
                                                                         AND (SD2.empno = RawLog.empno)\
                                                                         AND (SD2.entrancedate = RawLog.entrancedate)\
                                                                             ORDER BY RawLog.entrancetime ASC)");
    if( !query.exec() )
    {
        qDebug() << query.lastError().text();
    }

    query.prepare("UPDATE SD2 \
                   SET entrancetimeout = ( SELECT entrancetime FROM RawLog WHERE (RawLog.classification LIKE 'OUT')\
                                                                                 AND (SD2.empno = RawLog.empno)\
                                                                                 AND (SD2.entrancedate = RawLog.entrancedate)\
                                                                                 ORDER BY RawLog.entrancetime DESC)");
    if( !query.exec() )
    {
        qDebug() << query.lastError().text();
    }

    query.prepare("UPDATE SD3 \
           SET entrancetimein = ( SELECT entrancetime FROM RawLog WHERE (RawLog.classification LIKE 'IN')\
                                                                         AND (SD3.empno = RawLog.empno)\
                                                                         AND (SD3.entrancedate = RawLog.entrancedate)\
                                                                             ORDER BY RawLog.entrancetime ASC)");
    if( !query.exec() )
    {
        qDebug() << query.lastError().text();
    }
    query.prepare("UPDATE SD3 \
                   SET entrancetimeout = ( SELECT entrancetime FROM RawLog WHERE (RawLog.classification LIKE 'OUT')\
                                                                                 AND (SD3.empno = RawLog.empno)\
                                                                                 AND (SD3.entrancedate = RawLog.entrancedate)\
                                                                                 ORDER BY RawLog.entrancetime DESC)");
    if( !query.exec() )
    {
        qDebug() << query.lastError().text();
    }
}

//void MainWindow::on_action_Print_triggered()
//{
//    if ( nullptr == xlsxDoc )
//        return;

//    QPrintPreviewDialog dialog;
//    connect(&dialog, SIGNAL(paintRequested(QPrinter*)), this, SLOT(print(QPrinter*)));
//    dialog.exec();
//}
void MainWindow::print(QPrinter *printer)
{
    if ( nullptr == xlsxDoc )
    {
        QMessageBox msgBox;
        msgBox.setText( "xlsx document is empty." );
        msgBox.setIcon( QMessageBox::Warning );
        msgBox.exec();
        return;
    }

    Worksheet* wsheet = (Worksheet*) xlsxDoc->workbook()->activeSheet();
    if ( nullptr == wsheet )
    {
        QMessageBox msgBox;
        msgBox.setText( "worksheet is set." );
        msgBox.setIcon( QMessageBox::Warning );
        msgBox.exec();
        return;
    }

    QStringList items;
    foreach( QString curretnSheetName, xlsxDoc->sheetNames() )
    {
        items << curretnSheetName;
    }

    bool ok = false;
    QString selectedItem
        = QInputDialog::getItem( this, tr("QXlsx"), tr("Select sheet:"), items, 0, false, &ok );

    if (!ok)
    {
        QMessageBox msgBox;
        msgBox.setText( "Please select printable sheet" );
        msgBox.setIcon( QMessageBox::Warning );
        msgBox.exec();
        return;
    }

    int idx = 0;
    foreach (const QString& varList, items)
    {
        if ( varList == selectedItem )
        {
            break;
        }
        idx++;
    }
    qDebug() <<" setActiveSheet index=" << idx;
    xlsxDoc->workbook()->setActiveSheet( idx );
    wsheet = (Worksheet*) xlsxDoc->workbook()->activeSheet(); // set active sheet

    QList<QString> colTitle;
    QList<VLIST> xlsxData;
    QVector<int> printColumnStretch;
    int sheetIndexNumber = 0;

    int maxRow = -1;
    int maxCol = -1;
    QVector<CellLocation> clList = wsheet->getFullCells( &maxRow, &maxCol );

	// Fixed for Visual C++, cause VC++ does not support C99. 
	// https://docs.microsoft.com/en-us/cpp/c-language/ansi-conformance?view=vs-2017
    // QString arr[maxRow][maxCol];

	QVector <QVector <QString> > arr;

	for (int i = 0; i < maxRow; i++)
	{
		QVector<QString> tempVector;
		for (int j = 0; j < maxCol; j++)
		{
			tempVector.push_back(QString("")); 
		}
		arr.push_back(tempVector);
	}

    for ( int ic = 0; ic < clList.size(); ++ic )
    {
        CellLocation cl = clList.at(ic);

        int row = cl.row - 1;
        int col = cl.col - 1;

        QSharedPointer<Cell> ptrCell = cl.cell; // cell pointer

        QString strValue = ptrCell->value().toString();

        arr[row][col] = strValue;
    }

    for (int ir = 0 ; ir < maxRow; ir++)
    {
        VLIST vl;

        for (int ic = 0 ; ic < maxCol; ic++)
        {
            QString strValue = arr[ir][ic];
            if ( strValue.isNull() )
            {
                strValue = QString("");
            }
            qDebug() <<" strValue=" << strValue;
            vl.append( strValue );
        }

        xlsxData.append( vl );
        qDebug() <<"\n";
    }

    QVector<QString> printHeaders;

    for ( int ic = 0 ; ic < maxCol; ic++ )
    {
        std::string colString = convertFromNumberToExcelColumn( ( ic + 1 ) );
        QString strCol = QString::fromStdString( colString );
        qDebug() <<" strCol=" << strCol;
        colTitle.append( strCol );
        printHeaders.append( strCol );

        printColumnStretch.append( wsheet->columnWidth( (ic + 1) ) ); // TODO: check this code
    }

    CopycatTableModel copycatTableModel(colTitle, xlsxData); // model

    QPainter painter;
    if ( !painter.begin(printer) )
    {
        QMessageBox msgBox;
        msgBox.setText( "Can't start printer" );
        msgBox.setIcon( QMessageBox::Critical );
        msgBox.exec();
        return;
    }

    // print table
    TablePrinter tablePrinter(&painter, printer);
    if ( !tablePrinter.printTable( &copycatTableModel, printColumnStretch, printHeaders ) )
    {
        QMessageBox msgBox;
        msgBox.setText( tablePrinter.lastError() );
        msgBox.setIcon( QMessageBox::Warning );
        msgBox.exec();
        return;
    }

    // tablePrinter.setCellMargin( l, r, t, b );
    // tablePrinter.setPageMargin( l, r, t, b );

    painter.end();
}

std::string MainWindow::convertFromNumberToExcelColumn(int n)
{
    // main code from https://www.geeksforgeeks.org/find-excel-column-name-given-number/
    // Function to print Excel column name for a given column number

    std::string stdString;

    char str[1000]; // To store result (Excel column name)
    int i = 0; // To store current index in str which is result

    while ( n > 0 )
    {
        // Find remainder
        int rem = n % 26;

        // If remainder is 0, then a 'Z' must be there in output
        if ( rem == 0 )
        {
            str[i++] = 'Z';
            n = (n/26) - 1;
        }
        else // If remainder is non-zero
        {
            str[i++] = (rem-1) + 'A';
            n = n / 26;
        }
    }
    str[i] = '\0';

    // Reverse the string and print result
    std::reverse( str, str + strlen(str) );

    stdString = str;
    return stdString;
}
