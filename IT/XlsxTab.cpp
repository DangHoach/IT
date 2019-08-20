// XlsxTab.cpp

#include <QtGlobal>
#include <QObject>
#include <QString>
#include <QSharedPointer>
#include <QLayout>
#include <QVBoxLayout>
#include <QVariant>
#include <QFont>
#include <QBrush>
#include <QDebug>
#include <QTextCharFormat>

#include "XlsxTab.h"
#include "xlsxcell.h"
#include "xlsxformat.h"

XlsxTab::XlsxTab(QWidget* parent,
                 QXlsx::Document* ptrDoc,
                 QXlsx::AbstractSheet* ptrSheet,
                 int SheetIndex)
    : QWidget(parent)
{
    table = nullptr;
    sheet = nullptr;
    sheetIndex = -1;

    if ( nullptr == ptrSheet )
        return;

    table = new XlsxTableWidget( this );
    Q_ASSERT( nullptr != table );

    // set layout
    vLayout = new QVBoxLayout;
    vLayout->addWidget(table);
    this->setLayout(vLayout);

    document = ptrDoc; // set document
    sheet = ptrSheet; // set sheet data
    sheetIndex = SheetIndex; // set shett index
    if ( ! setSheet() )
    {
    }

}

XlsxTab::~XlsxTab()
{

    if ( nullptr != vLayout )
    {
        vLayout->deleteLater();
        vLayout = nullptr;
    }

    if ( nullptr != table )
    {
        table->clear();
        table->deleteLater();
        table = nullptr;
    }

}

bool XlsxTab::setSheet()
{
    if ( nullptr == sheet )
        return false;

    if ( nullptr == table )
        return false;

    // set active sheet
    sheet->workbook()->setActiveSheet( sheetIndex );
    Worksheet* wsheet = (Worksheet*) sheet->workbook()->activeSheet();
    if ( nullptr == wsheet )
        return false;

    // get full cells of sheet
    int maxRow = -1;
    int maxCol = -1;
    QVector<CellLocation> clList = wsheet->getFullCells( &maxRow, &maxCol );

    // set max count of row,col
      // NOTE: This part should be modified later.
      //  The maximum value of the sheet should be set to an appropriate value.
    table->setRowCount( maxRow  );
    table->setColumnCount( maxCol );

    for ( int ic = 0; ic < clList.size(); ++ic )
    {
          // cell location
          CellLocation cl = clList.at(ic);

          ////////////////////////////////////////////////////////////////////
          // First cell of tableWidget is 0.
          // But first cell of Qxlsx document is 1.
          int row = cl.row - 1;
          int col = cl.col - 1;

          ////////////////////////////////////////////////////////////////////
          // cell pointer
          QSharedPointer<Cell> ptrCell = cl.cell;

          ////////////////////////////////////////////////////////////////////
          // create new item of table widget
          QTableWidgetItem* newItem = new QTableWidgetItem();

          ///////////////////////////////////////////////////////////////////
          // value of cell
          QVariant var = cl.cell.data()->value();
          QString str = var.toString();

          ////////////////////////////////////////////////////////////////////
          // set text
          newItem->setText( str );

          ////////////////////////////////////////////////////////////////////
          // set item
          table->setItem( row, col, newItem );

          // TODO: set label of table widget ('A', 'B', 'C', ...)
              // QTableWidgetItem *QTableWidget::horizontalHeaderItem(int column) const
              // QTableWidgetItem *QTableWidget::verticalHeaderItem(int row) const

            ////////////////////////////////////////////////////////////////////
            // row height and column width
            double dRowHeight = wsheet->rowHeight( cl.row );
            double dColWidth  = wsheet->columnWidth( cl.col );

            // dRowHeight = dRowHeight * double(2.0);
            // dColWidth = dColWidth * double(2.0);

            // TODO: define ratio of widget col/row

            // table->setRowHeight( row, dRowHeight );
            // table->setColumnWidth( col, dColWidth );

          ////////////////////////////////////////////////////////////////////
          // font
          newItem->setFont( ptrCell->format().font() );

          ////////////////////////////////////////////////////////////////////
          // font color
          if ( ptrCell->format().fontColor().isValid() )
          {
            newItem->setTextColor( ptrCell->format().fontColor() );
          }

          ////////////////////////////////////////////////////////////////////
          // background color

          {
              QColor clrForeGround = ptrCell->format().patternForegroundColor();
              if ( clrForeGround.isValid() )
              {
                    // qDebug() << "[debug] ForeGround : " << clrForeGround;
              }

              QColor clrBackGround = ptrCell->format().patternBackgroundColor();
              if ( clrBackGround.isValid() )
              {
                    // TODO: You must use various patterns.
                    newItem->setBackground( Qt::SolidPattern );
                    newItem->setBackgroundColor( clrBackGround );
              }
          }

          ////////////////////////////////////////////////////////////////////
          // pattern
          Format::FillPattern fp = ptrCell->format().fillPattern();
          Qt::BrushStyle qbs = Qt::NoBrush;
          switch(fp)
          {
              case Format::PatternNone :       qbs = Qt::NoBrush; break;
              case Format::PatternSolid :      qbs = Qt::SolidPattern; break;
              case Format::PatternMediumGray :
              case Format::PatternDarkGray :
              case Format::PatternLightGray :
              case Format::PatternDarkHorizontal :
              case Format::PatternDarkVertical :
              case Format::PatternDarkDown :
              case Format::PatternDarkUp :
              case Format::PatternDarkGrid :
              case Format::PatternDarkTrellis :
              case Format::PatternLightHorizontal :
              case Format::PatternLightVertical :
              case Format::PatternLightDown :
              case Format::PatternLightUp :
              case Format::PatternLightTrellis :
              case Format::PatternGray125 :
              case Format::PatternGray0625 :
              case Format::PatternLightGrid :
              default:
              break;
          }

        /*
        QBrush qbr( ptrCell->format().patternForegroundColor(), qbs );
        newItem->setBackground( qbr );
        newItem->setBackgroundColor( ptrCell->format().patternBackgroundColor() );
        */

          ////////////////////////////////////////////////////////////////////
          // set alignment

          int alignment = 0;
          Format::HorizontalAlignment ha = ptrCell->format().horizontalAlignment();
          switch(ha)
          {
            case Format::AlignHCenter :
            case Format::AlignHFill :
            case Format::AlignHMerge :
            case Format::AlignHDistributed :
                alignment = alignment | Qt::AlignHCenter;
            break;

            case Format::AlignRight :
                alignment = alignment | Qt::AlignRight;
            break;

            case Format::AlignHJustify :
                alignment = alignment | Qt::AlignJustify;
            break;

            case Format::AlignLeft :
            case Format::AlignHGeneral :
            default:
                alignment = alignment | Qt::AlignLeft;
            break;
          }

          Format::VerticalAlignment va = ptrCell->format().verticalAlignment();
          switch(va)
          {
              case Format::AlignTop :
                  alignment = alignment |  Qt::AlignTop;
              break;

              case Format::AlignVCenter :
                  alignment = alignment |  Qt::AlignVCenter;
              break;

              case Format::AlignBottom :
                  alignment = alignment |  Qt::AlignBottom;
              break;

              case Format::AlignVJustify :
              case Format::AlignVDistributed :
              default:
                // alignment = alignment | (int)(Qt::AlignBaseline);
                alignment = alignment | QTextCharFormat::AlignBaseline;
              break;
          }

          newItem->setTextAlignment( alignment );

          ////////////////////////////////////////////////////////////////////

    }

    return true;
}

std::string XlsxTab::convertFromNumberToExcelColumn(int n)
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

