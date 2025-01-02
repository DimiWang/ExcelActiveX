/**
 * @file:exceldata.cpp   -
 * @description: This is Excel communication class.
 * @project: BENCH OnSemiconductor
 * @date: 2014\11\27 13-51-10
 *
 */


#include "excel.h"
#include <QMessageBox>
#include <QFile>
#include <QDebug>
#include <QRgb>
#include "axobject.h"
#include "excelenums.h"
#include <QLocale>






/****************************************************************************
 * @function name: Constructor
 *
 * @param:
 *          const QString filename  - file name *.xlsx
 * @description:
 ****************************************************************************/
Excel::Excel(const QString filename, bool use_thread, bool autosave) :
    QObject(0)
{
    m_autosave = autosave;
    m_opened = false;
    m_saved = false;
    m_filename = filename;
    mp_exlObject = new AxObject( "Excel.Application",use_thread,false);
    if(mp_exlObject->isValid()){
        mp_exlObject->setProperty( 0,"DisplayAlerts", 1);
        mp_exlObject->setProperty(0,"DisplayStatusBar",0);
        mp_exlObject->setProperty(0,"EnableEvents",0);
    }
    mp_currentSheet =0;
    mp_currentWorkBook =0;
    m_badFile = false;
    m_updatesOn = true;
}

Excel::~Excel()
{    
    if(!m_badFile){
        close();
        mp_exlObject->dynamicCall(0,"Workbooks.Close");
        mp_exlObject->method_run(mp_exlObject->id(),"Quit");
    }
    delete mp_exlObject;
}

bool Excel::validName(const QString &name)
{
    //allowed symbols
    //~    !    @    #    $    %    ^    &    (    )    -    _    =    +    {    }    |    ;    :    ,    <    .    >
    // but we use some of them
    QRegExp rx("[a-zA-Z0-9\\_\\<\\>\\$\\@\\!\\.\\#]+");
    if(name.size()<31 && name.toLower()!="history" && rx.exactMatch(name)){
        return true;
    }
    return false;
}

void Excel::setErrorSlot(QObject *pobj, const char *slot)
{
    mp_exlObject->setErrorSlot(pobj,slot);
}

QString Excel::fileName()
{
    QVariant v;
    if(mp_exlObject->property(this->mp_currentWorkBook, "FullName", &v))
        return v.toString();
    return QString();

}



/****************************************************************************
 * @function name: ExcelData::open()
 *
 * @param:
 *
 *       char mode  - file mode 'r' -read 'w'- write
 * @description: opens excel file for read or write
 * @return: ( bool )  success = true
 ****************************************************************************/
bool Excel::open()
{
    bool result = false;
    QVariant var;
    /* check if opened and file name is empty*/
    if (mp_exlObject != NULL && !m_opened ) {

        /* try open workbooks*/

        if (QFile::exists(m_filename) )
        {
            result = mp_exlObject->dynamicCall(0,"Workbooks.Open", &var, m_filename );
            mp_currentWorkBook = var.toInt();
            result = mp_exlObject->method_run(mp_currentWorkBook,"Activate");
        }
        else
        {
            result = mp_exlObject->dynamicCall(0,"Workbooks.Add",&var);
            mp_currentWorkBook = var.toInt();
            mp_exlObject->method_run(mp_currentWorkBook,"Activate");

        }
        //result = mp_workbook != 0;
        m_opened = result;
    }

    return result;
}

bool Excel::activate()
{
    return mp_exlObject->dynamicCall(0,"ActiveWindow.Activate");
}

void Excel::release()
{
    mp_exlObject->release();
}

void Excel::clearAbort()
{
    mp_exlObject->clearAbort();
}

bool Excel::isAborted()
{
    if(mp_exlObject && mp_exlObject->state() & AxObject::Abort) return true;
    return false;
}

bool Excel::isReadOnly()
{
    QVariant value;
    mp_exlObject->property(0,"ActiveWorkbook.ReadOnly",&value);
    return value.toBool();
}


/****************************************************************************
 * @function name: ExcelData::isOpen()
 *
 * @param:
 *             void
 * @description: Check if file is opened
 * @return: ( bool ) is opened = true
 ****************************************************************************/
bool Excel::isOpen() const
{
    return m_opened && mp_exlObject;
}

bool Excel::setZoom(int val)
{
    return mp_exlObject->setProperty(0,"ActiveWindow.Zoom",val);
}

int Excel::zoom()
{
    return 0;//mp_object->property("ActiveWindow.Zoom").toIn;
}


/****************************************************************************
 * @function name: ExcelData::addSheet()
 * @param:
 *       const QString sheetname  - sheet name
 * @description: adds sheet to workbook
 * @return: ( bool ) - success = true
 ****************************************************************************/
bool Excel::addSheet(const QString &sheetname)
{
    bool result =false;
    if (m_opened)
    {
        mp_exlObject->clearBag();
        do{
            QVariant var;

            if(!mp_exlObject->dynamicCall(mp_currentWorkBook,"Worksheets.Add", &var))  break;
            mp_currentSheet = var.toInt();

            mp_exlObject->assignObject("ActiveSheet", mp_currentSheet);
            if(!mp_exlObject->setProperty(mp_currentSheet, "Name", sheetname)) break;

            result = true;
        }while(0);
    }
    return result;
}


/****************************************************************************
 * @function name: ExcelData::setCellHint - ---
 * @param:
 *   qint32 row
 *   qint32 col
 *   const QString &text
 * @description:
 * @return: ( bool ) - success = true
 ****************************************************************************/
bool Excel::setCellHint(qint32 row, qint32 col, const QString &text)
{
    mp_exlObject->clearBag();
    Excel::Cell cell((int)row-1,(int)col-1);
    QVariant v;
    if(mp_exlObject->dynamicCall(mp_currentSheet, QString("Range(\"%1\").AddComment").arg(cell.toRange()),&v)){
        AxObject::Class comment = v.toInt();
        mp_exlObject->setProperty(comment,"Visible",false);
        mp_exlObject->dynamicCall(comment,"Text",0,text);
        return true;
    }
    return false;
}


/****************************************************************************
 * @function name: ExcelData::setCurrentSheet - ---
 * @param:
 *   const QString sheetname  - sheetn name
 * @description: Sets active sheet
 * @return: ( bool ) - success = true
 ****************************************************************************/
bool Excel::setCurrentSheet(const QString &sheetname)
{
    bool result = false;
    if (m_opened)
    {
        mp_exlObject->clearBag();
        result = mp_exlObject->dynamicCall(0,QString("Sheets(\"%1\").Activate").arg(sheetname));
        mp_currentSheet = mp_exlObject->object(QString("Sheets(\"%1\")").arg(sheetname));
    }
    return result;
}

AxObject::Class Excel::currentSheet()
{
    if (m_opened)
    {
        if(mp_currentSheet ==0)
        {
            QVariant v;
            if(mp_exlObject->property(0,"ActiveSheet",&v))
            {
                mp_currentSheet = v.toInt();
            }
        }
    }
    return mp_currentSheet;
}

AxObject::Class Excel::currentWorkBook()
{
    if (m_opened)
    {
        if(mp_currentWorkBook ==0)
        {
            QVariant v;
            if(mp_exlObject->property(0,"ActiveWorkbook",&v))
            {
                mp_currentWorkBook = v.toInt();
            }
        }
    }
    return mp_currentWorkBook;
}

/****************************************************************************
 * @function name: ExcelData::removeSheet()
 * @param:
 *    const QString sheetname
 * @description: removes sheet by name
 * @return: ( bool ) success = true
 ****************************************************************************/
bool Excel::removeSheet(const QString &sheetname)
{
    bool result = false;
    if (m_opened)
    {
        mp_exlObject->clearBag();
        result = mp_exlObject->dynamicCall(0,QString("Sheets(\"%1\").Delete").arg(sheetname));
    }
    return result;
}

/****************************************************************************
 * @function name: ExcelData::removeSheet()
 * @param:
 *    const int sheetnumber
 * @description: removes sheet by number
 * @return: ( bool ) success = true
 ****************************************************************************/
bool Excel::removeSheet(int sheetnumber)
{
    bool result = false;
    if (m_opened)
    {
        mp_exlObject->clearBag();
        result = mp_exlObject->dynamicCall(0,QString("Sheets(%1).Delete").arg(sheetnumber));
    }
    return result;
}


/****************************************************************************
 * @function name: ExcelData::save - ---
 * @description:
 * @return: ( bool ) success = true
 ****************************************************************************/
bool Excel::save(void)
{
    bool ok=0;
    if(m_filename.isEmpty()) return false;
    mp_exlObject->clearBag();
    if(!m_saved  ){
        ok = mp_exlObject->dynamicCall(currentWorkBook(),"SaveAs",0, m_filename);
    }
    else ok =  mp_exlObject->dynamicCall(currentWorkBook(),"Save");
    if(ok ) m_saved =1;
    return ok;
}

bool Excel::saveAs(const QString &)
{    
    bool ok ;
    ok = mp_exlObject->dynamicCall(mp_currentWorkBook,"SaveAs",0, m_filename);
    if(ok) m_saved = 1;
    return m_saved;
}


/****************************************************************************
 * @function name: ExcelData::write - ---
 * @param:
 *    qint32 i
 *    qint32 j
 *    const QVariant data
 * @description: writes data to cell i,j
 * @return: ( bool ) success = true
 ****************************************************************************/

bool Excel::write(Excel::Cell cell, const QVariant &data, const QFont &font )
{
    return write(cell.xlRow(),cell.xlCol(),data,font);
}

bool Excel::write(const QString &range, const QStringList &l, const QFont &font)
{    
    mp_exlObject->clearBag();
    bool result = false;
    if (m_opened && Range_Is_Valid(range) )
    {
        result = mp_exlObject->setProperty(mp_currentSheet, QString("Range(\"%1\").FormulaArray").arg(range),l);


        if(result && font != QFont())
        {
            mp_exlObject->setProperty(mp_currentSheet, QString("Range(\"%1\").Font.Name").arg(range), font.family());
            if(font.pointSize()>0)
                mp_exlObject->setProperty(mp_currentSheet, QString("Range(\"%1\").Font.Size").arg(range), font.pointSize());
            mp_exlObject->setProperty(mp_currentSheet, QString("Range(\"%1\").Font.Bold").arg(range), font.bold());
            mp_exlObject->setProperty(mp_currentSheet, QString("Range(\"%1\").Font.Italic").arg(range), font.italic());
        }
    }
    return result;
}

bool Excel::mergeRange(const QString &range,bool on)
{
    if(m_opened && Range_Is_Valid(range)){
        return mp_exlObject->setProperty(mp_currentSheet, QString("Range(\"%1\").MergeCells").arg(range), on);
    }
    return false;
}


bool Excel::write(qint32 row, qint32 col, const QVariant &data, const QFont &font)
{
    mp_exlObject->clearBag();
    bool result = false;
    if (m_opened && row > 0 && col > 0 )
    {
        result = mp_exlObject->setProperty(mp_currentSheet, QString("Cells(%1,%2).Value").arg(row).arg(col),data);
        if(result && font != QFont())
        {
            mp_exlObject->setProperty(mp_currentSheet, QString("Cells(%1,%2).Font.Name").arg(row).arg(col), font.family());
            if(font.pointSize()>0)
                mp_exlObject->setProperty(mp_currentSheet, QString("Cells(%1,%2).Font.Size").arg(row).arg(col), font.pointSize());
            mp_exlObject->setProperty(mp_currentSheet, QString("Cells(%1,%2).Font.Bold").arg(row).arg(col), font.bold());
            mp_exlObject->setProperty(mp_currentSheet, QString("Cells(%1,%2).Font.Italic").arg(row).arg(col), font.italic());
        }
    }
    return result;
}


bool Excel::write(AxObject::Class range, qint32 row, qint32 col, const QVariant &data, const QFont &font)
{
    mp_exlObject->clearBag();
    bool result = false;
    if (m_opened && row > 0 && col > 0 )
    {
        int retry=5;        
        while(!result && retry--)
            result = mp_exlObject->setProperty(range, QString("Cells(%1,%2).Value").arg(row).arg(col),data);

        if(result && font != QFont())
        {
            mp_exlObject->setProperty( range,QString("Cells(%1,%2).Font.Name").arg(row).arg(col), font.family());
            if(font.pointSize()>0)
                mp_exlObject->setProperty(range, QString("Cells(%1,%2).Font.Size").arg(row).arg(col), font.pointSize());
            mp_exlObject->setProperty(range, QString("Cells(%1,%2).Font.Bold").arg(row).arg(col), font.bold());
            mp_exlObject->setProperty(range, QString("Cells(%1,%2).Font.Italic").arg(row).arg(col), font.italic());
        }
    }
    return result;
}

bool Excel::cellVisible(int row, int col)
{
    return mp_exlObject->dynamicCall(mp_currentSheet,QString("Cells(%1, %2).Select").arg(row).arg(col));
}

Excel::Table* Excel::CreateTable(const Excel::Rect &rect
                                 , const QStringList &headerData
                                 , const QStringList &tableData)
{
    Excel::Table *ptable = CreateTable(rect, new TableHeader1Line(), new TableStandardBody());

    if(!headerData.isEmpty())
        ptable->setHeaderData(headerData);
    if(!tableData.isEmpty())
        ptable->setTableData(tableData);
    return ptable;
}

Excel::Table * Excel::CreateTable(const Rect &rect
                                  , DataArea *tableHeaderArea
                                  , DataArea *tableDataArea)
{
    if(!m_opened) return 0;

    Excel::Table *ptable =  new Excel::Table(this,rect);

    if(tableHeaderArea){
        ptable->setHeaderArea(tableHeaderArea);
    }

    if(tableDataArea){
        ptable->setDataArea(tableDataArea);
    }

    //draw frame around table
    Excel::Frame frame(Excel::Frame::DrawBottom
                       |Excel::Frame::DrawLeft
                       |Excel::Frame::DrawRight
                       |Excel::Frame::DrawTop
                       );

    frame.setStyle(Excel::Frame::Bottom, Excel::Frame::LineContinuous);
    frame.setStyle(Excel::Frame::Top, Excel::Frame::LineContinuous);
    frame.setStyle(Excel::Frame::Left, Excel::Frame::LineContinuous);
    frame.setStyle(Excel::Frame::Right, Excel::Frame::LineContinuous);
    drawFrame(rect,frame);

    return ptable;
}

bool Excel::AppendRow(Table *ptable, const QStringList &tableData)
{
    return ptable->appendDataRow(tableData);
}

bool Excel::SetDataToColumn(Excel::Table *ptable, const QVariantList &data, int column)
{

    bool result = false;
    mp_exlObject->clearBag();
    setUpdatesOn(0);
    if(!data.isEmpty() && ptable)
    {
        VARIANT v;
        AxObject::QVariantList_to_2D_VARIANT(data,1,data.count(),v);
        if(mp_exlObject->setPropertyVariant(currentSheet(),
                                            QString("Range(\"%1\").Value")
                                            .arg(ptable->dataRect().column(column).toRange()), v))
        {
            result = true;
        }
    }
    setUpdatesOn(1);

    return result;
}


bool Excel::SetDataToRange(const Excel::Rect &rect,  QVariantList data)
{

    bool result = false;
    mp_exlObject->clearBag();
    if(data.size() <rect.width()*rect.height()){
        for(int i=data.size();i<rect.width()*rect.height();i++)
            data.append(QString(""));
    }
    VARIANT v;
    AxObject::QVariantList_to_2D_VARIANT(data,rect.width(),rect.height(),v);
    if(mp_exlObject->setPropertyVariant(currentSheet(),
                                        QString("Range(\"%1\").Value")
                                        .arg(rect.toRange()), v))
    {
        result = true;
    }

    return result;
}


bool Excel::SetDataToRow(Excel::Table *ptable, const QVariantList &data, int row)
{
    bool result = false;
    mp_exlObject->clearBag();
    setUpdatesOn(0);
    if(!data.isEmpty() && ptable)
    {
        VARIANT v;
        AxObject::QVariantList_to_2D_VARIANT(data,data.count(),1,v);
        if(mp_exlObject->setPropertyVariant(currentSheet(),
                                            QString("Range(\"%1\").Value")
                                            .arg(ptable->dataRect().row(row).toRange()), v))
        {
            result = true;
        }
    }
    setUpdatesOn(1);
    return result;
}

int Excel::width(const QString &range)
{
    QVariant v;
    if(mp_exlObject->property(currentSheet(),QString("Range(\"%1\").Width").arg(range),&v))
    {
        return v.toInt();
    }
    return 0;
}

int Excel::height(const QString &range)
{
    QVariant v;
    if( mp_exlObject->property( currentSheet(), QString("Range(\"%1\").Height").arg(range) , &v ) )
    {
        return v.toInt();
    }
    return 0;
}

void Excel::setUpdatesOn(bool on)
{
    if(!mp_exlObject) return;    
    if(on != m_updatesOn){
        //mp_exlObject->setProperty(0,"DisplayStatusBar",on);
        mp_exlObject->setProperty(0,"EnableEvents",on);
        setCalculation(on);
        m_updatesOn = on;
    }
}
void Excel::setCalculation(bool on)
{
    if(!mp_exlObject) return;
    if(!on) mp_exlObject->setProperty(0,"Calculation",xlCalculationManual);
    else {
        mp_exlObject->setProperty(0,"Calculation",xlCalculationAutomatic);
    }
}

void Excel::setScreenUpdate(bool on)
{
    mp_exlObject->setProperty(0,"ScreenUpdating", on);
}

void Excel::recalculate(){
    mp_exlObject->dynamicCall(0,"Calculate");
}

bool Excel::removeSheetsList(const QStringList &sheets)
{
    foreach(const QString &sheetname, sheets){
        if(!removeSheet(sheetname) ) return false;
    }
    return true;
}

int Excel::version()
{

    return 0;
}


AxObject::Class Excel::CreateChart(const Chart &chart)
{  
    mp_exlObject->clearBag();

    AxObject::Class pObj = 0 ;
    if(!m_opened) return 0;
    QVariant var;
    const int _XLChartType[] ={ xlXYScatterLinesNoMarkers,       xlLine  };

    mp_exlObject->dynamicCall(this->mp_currentSheet,QString("Range(\"%1\").Select").arg(chart.cellDataRange.toRange()));

    if(!chart.rect.isEmpty()){
        mp_exlObject->dynamicCall(this->mp_currentSheet, "Shapes.AddChart2",&var ,-1, _XLChartType[chart.type]
                                  ,chart.rect.x(),  chart.rect.y(),   chart.rect.width(),   chart.rect.height());
    }
    else {
        mp_exlObject->dynamicCall(this->mp_currentSheet, "Shapes.AddChart2",&var ,-1, _XLChartType[chart.type]);
    }

    pObj = var.toInt();
    if(pObj==0) return 0;

    mp_exlObject->method_run(pObj,"Select");
    mp_exlObject->setProperty(0,"ActiveChart.ChartType", _XLChartType[chart.type]);

    AxObject::Class  pRange = mp_exlObject->queryObject(currentSheet(), QString("Range(\"%1\")").arg(chart.cellDataRange.toRange()));
    mp_exlObject->dynamicCall(0,"ActiveChart.SetSourceData",0, AxObjectType(DISPATCH,(quint32)pRange),2);


    mp_exlObject->setProperty(0,"ActiveChart.ChartTitle.Text",chart.title);

    int series_count =0;
    if(mp_exlObject->property(0,"ActiveChart.SeriesCollection.Count",&var))
    {
        series_count = var.toInt();
        for(int i=1;i<series_count+1;i++){
            mp_exlObject->setProperty(0,QString("ActiveChart.SeriesCollection(%1).Border.Weight").arg(i),xlThin);
        }
    }
    const QColor std_colors[7] = {
        QColor(Qt::blue)
        ,QColor(Qt::magenta)
        ,QColor(24,135,9)
        ,QColor(244,131,0)
        ,QColor(133,153,7)
        ,QColor(137,82,37)
        ,QColor(Qt::black)
    };


    for(int i=0;i<qMin(series_count,7);i++)
    {
        mp_exlObject->setProperty(0, QString("ActiveChart.SeriesCollection(%1).Format.Line.ForeColor.RGB").arg(i+1), std_colors[i]);
    }


    mp_exlObject->setProperty(0,"ActiveChart.HasTitle", !chart.title.isEmpty());
    mp_exlObject->setProperty(0,"ActiveChart.ChartTitle.Caption", QString("Automated Measurement Output Chart %1").arg(series_count));
    mp_exlObject->setProperty(0,"ActiveChart.ChartTitle.Text", chart.title);
    mp_exlObject->setProperty(0,"ActiveChart.ChartTitle.AutoScaleFont", false);
    mp_exlObject->setProperty(0,"ActiveChart.ChartTitle.Font.Name", "Arial");
    mp_exlObject->setProperty(0,"ActiveChart.ChartTitle.Font.Size", 12);
    mp_exlObject->setProperty(0,"ActiveChart.ChartTitle.Font.Bold", true);

    mp_exlObject->setProperty(0,"ActiveChart.Axes(2).TickLabels.Font.Name","Arial");
    mp_exlObject->setProperty(0,"ActiveChart.Axes(2).TickLabels.Font.Size",8);
    mp_exlObject->setProperty(0,"ActiveChart.Axes(2).TickLabels.AutoScaleFont",false);
    mp_exlObject->setProperty(0,"ActiveChart.Axes(2).TickLabelPosition",xlLow );
    mp_exlObject->setProperty(0,"ActiveChart.Axes(2).HasMajorGridLines",true);

    mp_exlObject->setProperty(0,"ActiveChart.Axes(2).MajorGridLines.Border.Weight",xlHairLine );
    mp_exlObject->setProperty(0,"ActiveChart.Axes(2).MajorGridLines.Border.ColorIndex",16);
    mp_exlObject->setProperty(0,"ActiveChart.Axes(2).HasMinorGridlines",false);


    //xlValue
    if(!chart.yAxis.isEmpty())
    {
        mp_exlObject->setProperty(0,"ActiveChart.Axes(2).HasTitle", true);
        mp_exlObject->setProperty(0,"ActiveChart.Axes(2).AxisTitle.Caption", chart.yAxis);
        mp_exlObject->setProperty(0,"ActiveChart.Axes(2).AxisTitle.Font.Name","Arial");
        mp_exlObject->setProperty(0,"ActiveChart.Axes(2).AxisTitle.Font.Size",9);
        mp_exlObject->setProperty(0,"ActiveChart.Axes(2).AxisTitle.Font.Bold",true);
        mp_exlObject->setProperty(0,"ActiveChart.Axes(2).AxisTitle.AutoScaleFont",false);
    }

    //xlCategory
    if(!chart.xAxis.isEmpty()){
        mp_exlObject->setProperty(0,"ActiveChart.Axes(1).HasTitle" , true);
        mp_exlObject->setProperty(0,"ActiveChart.Axes(1).AxisTitle.Caption" , chart.xAxis);
        mp_exlObject->setProperty(0,"ActiveChart.Axes(1).AxisTitle.AutoScaleFont" , false);
        mp_exlObject->setProperty(0,"ActiveChart.Axes(1).AxisTitle.Font.Name" , "Arial");
        mp_exlObject->setProperty(0,"ActiveChart.Axes(1).AxisTitle.Font.Size" , 9);
        mp_exlObject->setProperty(0,"ActiveChart.Axes(1).AxisTitle.Font.Bold" , true);
    }

    const int _XLScaleType[] = {xlLinear,xlLogarithmic};
    mp_exlObject->setProperty(0,"ActiveChart.Axes(1).ScaleType" , _XLScaleType[chart.xScaleType]);
    mp_exlObject->setProperty(0,"ActiveChart.Axes(1).TickLabels.Font.Name" , "Arial");
    mp_exlObject->setProperty(0,"ActiveChart.Axes(1).TickLabels.Font.Size" , 8);
    mp_exlObject->setProperty(0,"ActiveChart.Axes(1).TickLabels.AutoScaleFont" , 8);
    mp_exlObject->setProperty(0,"ActiveChart.Axes(1).TickLabelPosition" , xlLow);
    mp_exlObject->setProperty(0,"ActiveChart.Axes(1).HasMajorGridlines" , chart.majorGridLines);
    mp_exlObject->setProperty(0,"ActiveChart.Axes(1).MinorGridlines.Border.Weight" , xlHairLine);
    mp_exlObject->setProperty(0,"ActiveChart.Axes(1).MinorGridlines.Border.ColorIndex" , 16);
    mp_exlObject->setProperty(0,"ActiveChart.Axes(1).HasMinorGridlines" , chart.minorGridLines);
    mp_exlObject->setProperty(0,"ActiveChart.Axes(1).MinorGridlines.Border.Weight" , xlHairLine);
    mp_exlObject->setProperty(0,"ActiveChart.Axes(1).MinorGridlines.Border.ColorIndex" , 15);



    QVariant width,left,height;
    mp_exlObject->property(0,"ActiveChart.ChartArea.Width",&width);
    mp_exlObject->property(0,"ActiveChart.ChartArea.Height",&height);
    mp_exlObject->setProperty(0,"ActiveChart.PlotArea.Height", height.toInt() -40);
    mp_exlObject->setProperty(0,"ActiveChart.PlotArea.Interior.ColorIndex",xlAutomatic);
    mp_exlObject->property(0,"ActiveChart.PlotArea.Left",&left);
    mp_exlObject->setProperty(0,"ActiveChart.PlotArea.Width", width.toInt() -2*left.toInt());

    // legend
    if(chart.legendVisible){
        mp_exlObject->setProperty(0,"ActiveChart.HasLegend",true);
        mp_exlObject->setProperty(0,"ActiveChart.Legend.Font.Name","Arial");
        mp_exlObject->setProperty(0,"ActiveChart.Legend.Font.Size",8);
        mp_exlObject->setProperty(0,"ActiveChart.Legend.IncludeInLayout",false);
        mp_exlObject->setProperty(0,"ActiveChart.Legend.Position",xlLegendPositionRight);
        mp_exlObject->setProperty(0,"ActiveChart.Legend.Format.Fill.Visible",true);
        mp_exlObject->setProperty(0,"ActiveChart.Legend.Format.Fill.ForeColor.ObjectThemeColor",msoThemeColorBackground1);
        mp_exlObject->setProperty(0,"ActiveChart.Legend.Format.Fill.ForeColor.TintAndShade",true);
        mp_exlObject->setProperty(0,"ActiveChart.Legend.Format.Fill.ForeColor.Brightness",0);
        mp_exlObject->setProperty(0,"ActiveChart.Legend.Format.Fill.Transparency",0);
        mp_exlObject->setProperty(0,"ActiveChart.Legend.Interior.PatternColorIndex",2);//white
        mp_exlObject->setProperty(0,"ActiveChart.Legend.Interior.Pattern",xlSolid);//white
        mp_exlObject->setProperty(0,"ActiveChart.Legend.Border.Weight",xlHairLine);
        mp_exlObject->setProperty(0,"ActiveChart.Legend.Border.ColorIndex",1); //Black
        mp_exlObject->setProperty(0,"ActiveChart.Legend.AutoScaleFont",false); //Black
    }

    return pObj;
}

bool Excel::SetChartData(AxObject::Class chart, const QString &range)
{
    return false;
}


/*


  LIST:ActiveSheet.Shapes("Chart 7").IncrementLeft -429
  expression .AddChart2(Style,XlChartType,Left,Top,Width,Height,NewLayout)
*/


//bool ExcelData::setBorders(qint32 row, qint32 col, qint32 width, qint32 height, ExcelData::XlBorders borders)
//{
//    return 1;
//}


/****************************************************************************
 * @function name: ExcelData::read - ---
 *
 * @param:
 *
 *      qint32 i
 *      qint32 j
 *      QVariant &data
 * @description: rads data from cell i,j
 * @return: ( bool ) success = true
 ****************************************************************************/
bool Excel::read(qint32 row, qint32 col, QVariant &data)
{
    bool result = false;
    if (m_opened && row > 0 && col > 0 )
    {
        result = mp_exlObject->property(this->mp_currentSheet, QString("Cells( %1, %2).Value").arg(row).arg(col),&data);
    }
    return result;
}

void Excel::setAutoSaveOn(bool on)
{
    m_autosave = on;
}

bool Excel::readRange(const QString &range, QVariantList *presult)
{
    QVariant data;
    bool r= mp_exlObject->property(this->mp_currentSheet,QString("Range(\"%1\").Value").arg(range),&data);
    if(presult) *presult = data.toList();
    return r;
}

bool Excel::writeRange(const QString &range, QVariantList l)
{
    // TODO
    return false;
}

bool Excel::setRangeAlignment(const QString &range, unsigned int align)
{
    if(m_opened && Range_Is_Valid(range)){
        return  mp_exlObject->setProperty(mp_currentSheet, QString("Range(\"%1\").HorizontalAlignment").arg(range), xlCenter);
    }
    return false;
}

bool Excel::drawFrame(const Rect &rect, const Frame &f){    
    return drawFrame(rect.toRange(),f);
}

bool Excel::drawFrame(const QString &range , const Frame &f)
{
    mp_exlObject->clearBag();
    if(m_opened){
        for(int i=0;i<5;i++)
        {
            if(f.drawLine(i))
            {
                const int xlBorders[] = {9,7,10,8,12,11};
                const int xlStyles[] = { -4142,1,-4119};
                const int xlWidth[] = {1, -4138 ,4, 2};

                mp_exlObject->setProperty(currentSheet(), QString("Range(\"%1\").Borders(%2).LineStyle")
                                          .arg(range).arg(xlBorders[i]), xlStyles[f.style(i)]);

                if( f.style(i) != Frame::LineDouble)
                    mp_exlObject->setProperty(currentSheet(), QString("Range(\"%1\").Borders(%2).Weight")
                                              .arg(range).arg(xlBorders[i]), xlWidth[f.width(i)]);
            }
        }
        return true;
    }
    return false;
}



/****************************************************************************
 * @function name: ExcelData::setColor()
 *
 * @param:
 *    qint32 i
 *    qint32 j
 *    const QColor background
 *    const QColor foreground
 * @description: Sets color of background ,foreground cell
 * @return: ( bool ) - success = true
 ****************************************************************************/
bool Excel::setColor(qint32 row, qint32 col, const QColor background, const QColor foreground)
{
    bool result = false;
    if (m_opened && row > 0 && col > 0 )
    {
        quint32 color;
        color = (background.red() << 0) | (background.green() << 8) | (background.blue() << 16);
        result = mp_exlObject->setProperty(currentSheet(), QString("Cells(%1,%2).Interior.Color").arg(row).arg(col), QVariant(color));

        color = (foreground.red() << 0) | (foreground.green() << 8) | (foreground.blue() << 16);
        result = mp_exlObject->setProperty(currentSheet(), QString("Cells(%1,%2).Font.Color").arg(row).arg(col), QVariant(color));
        result = true;
    }
    return result;
}

bool Excel::setColor(const Rect &rect, const QColor background, const QColor foreground)
{
    return    setColor(rect.toRange(),background,foreground);
}

bool Excel::setColor(const QString &range, const QColor background, const QColor foreground)
{
    bool result = false;
    if (m_opened && Range_Is_Valid(range) )
    {
        quint32 color;
        color = (background.red() << 0) | (background.green() << 8) | (background.blue() << 16);
        result = mp_exlObject->setProperty(currentSheet(), QString("Range(\"%1\").Interior.Color").arg(range), QVariant(color));

        color = (foreground.red() << 0) | (foreground.green() << 8) | (foreground.blue() << 16);
        result = mp_exlObject->setProperty(currentSheet(), QString("Range(\"%1\").Font.Color").arg(range), QVariant(color));

        result = true;
    }
    return result;
}

/****************************************************************************
 * @function name: ExcelData::color()
 *
 * @param:
 *      qint32 i
 *      qint32 j
 *      QColor &background
 *      QColor &foreground
 * @description: gets color of background and foregorund
 * @return: ( bool ) success = true
 ****************************************************************************/
bool Excel::color(qint32 row, qint32 col, QColor &background, QColor &foreground)
{
    bool result = false;
    if (m_opened && row > 0 && col > 0 )
    {
        do{
            QVariant v1, v2;
            if(!mp_exlObject->property(mp_currentSheet, QString("Cells(%1,%2).Interior.Color").arg(row).arg(col),&v1)) break;
            background = QColor::fromRgb(v1.toInt());

            if(!mp_exlObject->property(mp_currentSheet, QString("Cells(%1,%2).Font.Color").arg(row).arg(col),&v2)) break;
            foreground = QColor::fromRgba(v2.toInt());

            result = true;
        }while(0);
    }
    return result;
}


/****************************************************************************
 * @function name: ExcelData::setVisible()
 * @param:
 *      bool visible
 * @description: sets visibility of workbook
 * @return: ( void )
 ****************************************************************************/
bool Excel::setVisible(bool visible)
{
    if ( isOpen())
    {
        return  mp_exlObject->setProperty(0, "Visible", visible);
    }
    return false;
}

bool Excel::visible()
{
    if ( isOpen())
    {
        QVariant res;
        bool ok = mp_exlObject->property(0, "Visible",&res);
        return  ok && res.toBool();
    }
    return false;
}

bool Excel::valid()
{
    if ( isOpen())
    {
        QVariant tmp;
        return mp_exlObject->property(0, "Visible",&tmp);
    }
    return false;
}

bool Excel::resizeCells(double size)
{
    return mp_exlObject->dynamicCall(0,"Cells.Select") && mp_exlObject->setProperty(0, "Selection.ColumnWidth",size);
}


/****************************************************************************
 * @function name: ExcelData::close()
 *
 * @param:
 *             void
 * @description: closes workbook
 * @return: ( void )
 ****************************************************************************/
void Excel::close()
{       
    if(isOpen()){
        if(m_autosave) save();
        mp_exlObject->dynamicCall(currentWorkBook(), "Close");
        m_opened = false;
    }
}


int Excel::sheetsCount()
{
    QVariant var;
    mp_exlObject->property(currentWorkBook(),"Sheets.Count", &var);
    return var.toInt();
}

QStringList Excel::sheetsList()
{
    QStringList result;
    int sheets_count = sheetsCount();
    if(sheets_count>0)
    {
        for(int i=1;i<=sheets_count;i++)
        {
            QVariant var;
            mp_exlObject->property(currentWorkBook(), QString("Sheets(%1).Name").arg(i),&var);
            result += var.toString();
        }
    }
    return result;
}


QString Excel::Cell_To_Name(const Excel::Cell &cell,bool fixed)
{
    int index = 26*26*26*26;

    int col = cell.x();
    QString col_abc;
    while(index>1)
    {
        if(col>=index )
        {
            col_abc += QChar('A' +col/index-1);
            col = col % index;
        }
        index /=26;
    }
    col_abc += QChar('A' +col);

    QString cellName;
    if(fixed)
        cellName = QString("$%1$%2").arg(col_abc).arg(1+cell.y());
    else cellName = QString("%1%2").arg(col_abc).arg(1+cell.y());

    return cellName;
}

QString Excel::Rect_To_Range(const Excel::Rect &rect,bool fixed)
{
    Cell c = rect.p2();
    if(c.x()==0)        c.setX(1);
    else c.setX(c.x()-1);
    if(c.y() ==0)        c.setY(1);
    else c.setY(c.y()-1);
    return QString("%1:%2")
            .arg(Cell_To_Name(rect.p1(),fixed))
            .arg(Cell_To_Name(c,fixed));
}

Excel::Cell Excel::Name_To_Cell(const QString &xl_name)
{
    return Cell();
}

Excel::Rect Excel::Range_To_Rect(const QString &xl_range)
{
    return Rect();
}

bool Excel::Range_Is_Valid(const QString &range)
{
    // TODO
    return true;
}




bool Excel::test()
{
    mp_exlObject->blockSignals(1);
    QVariant data;
    bool result = read(1,1,data);
    mp_exlObject->blockSignals(0);
    if(!result)
        m_badFile =  true;
    return result;
}


Excel::Table::Table()
{
    mp_headerArea = 0;
    mp_dataArea=0;
    m_rows_count =0;
}

Excel::Table::Table(Excel *pexcel, const Excel::Rect &rect)
{
    mp_excel = pexcel;
    m_font = QFont();
    m_rows_count = 0;
    m_rect = rect;
    m_height =  rect.height();
    m_width = rect.width();
    mp_headerArea = 0;
    mp_dataArea=0;
}


void Excel::Table::setDataArea(Excel::DataArea *dataArea)
{
    if( dataArea != 0 )
    {
        mp_dataArea = dataArea;
        // new rect
        mp_dataArea->setX(rect().x());
        mp_dataArea->setY(mp_headerArea->height()+rect().y());
        mp_dataArea->setWidth(rect().width());
        mp_dataArea->setHeight(rect().height()-mp_headerArea->height());

        m_dataRange = mp_excel->object()->queryObject(mp_excel->currentSheet()
                                                      , QString("Range(\"%1\")").arg(mp_dataArea->rect().toRange()));
    }
}

void Excel::Table::setHeaderArea(Excel::DataArea *dataArea)
{
    mp_headerArea=dataArea;

    if(dataArea){
        mp_headerArea->setX(rect().x());
        mp_headerArea->setY(rect().y());
        mp_headerArea->setWidth(rect().width());
    }
}

void Excel::Table::setHeaderData(const QStringList &data)
{    
    if(mp_headerArea){
        mp_headerArea->placeData(mp_excel, rect().x(), rect().y(), data );
    }
}

void Excel::Table::setTableData(const QStringList &data)
{    
    if(mp_dataArea)
        mp_dataArea->placeData(mp_excel,  rect().x(), rect().y(), data);
}

bool Excel::Table::appendDataRow(const QStringList &data)
{
    bool result=false;

    if(!data.isEmpty())
    {
        int i=0;
        int row=0,column=0;
        foreach(const QString &txt, data)
        {
            row = i/m_width + 1 + m_rows_count;
            column = i%m_width+1;
            if(!mp_excel->write(m_dataRange, row, column, txt ,m_font))
                return false;
            i++;
        }
        m_rows_count = row;
        result= true;
    }
    return result;
}




Excel::TableHeader1Line::TableHeader1Line()
    :DataArea()
{
    QFont font;
    font.setBold(1);
    font.setPointSize(10);
    font.setFamily("Arial");

    Frame frame(Frame::DrawBottom
                |Frame::DrawLeft
                |Frame::DrawRight
                |Frame::DrawTop
                );
    frame.setStyle(Frame::Bottom, Frame::LineDouble);
    frame.setStyle(Frame::Top, Frame::LineContinuous);
    frame.setStyle(Frame::Left, Frame::LineContinuous);
    frame.setStyle(Frame::Right, Frame::LineContinuous);

    m_font = font;
    m_frame = frame;
    m_foreground =   Qt::darkGray;
    m_background = Qt::white;
    this->m_height = 1;
}

void Excel::TableHeader1Line::placeData(Excel *pexcel, int x0, int y0, const QStringList &data)
{    
    QStringList texts;
    QMap<int,QString> tooltips;

    QRegExp rx("(.*)<tooltip>(.*)<\\/>");
    for(int i=0;i<data.count();i++)
    {
        if(rx.exactMatch(data[i])){
            tooltips[i]=rx.cap(2);
            texts.append(rx.cap(1));
        }
        else
            texts.append(data[i]);
    }

    if(texts.count())
    {
        pexcel->write(rect().toRange(), texts);
        pexcel->setColor(rect().toRange(), Qt::darkGray, Qt::white);
        pexcel->drawFrame(rect().toRange(), m_frame);
        // set tooltips
        foreach(const int &i, tooltips.keys())
        {
            pexcel->setCellHint(rect().p1().xlCol()+i, rect().p1().xlRow(), tooltips[i]);
        }
    }
}


Excel::TableStandardBody::TableStandardBody()
    :DataArea()
{
    QFont font;
    font.setPointSize(10);
    font.setFamily("Arial");
    font.setBold(0);
    Frame frame(Frame::DrawBottom
                |Frame::DrawLeft
                |Frame::DrawRight
                );
    frame.setStyle(Frame::Bottom, Frame::LineContinuous);
    frame.setStyle(Frame::Left, Frame::LineContinuous);
    frame.setStyle(Frame::Right, Frame::LineContinuous);
    frame.setStyle(Frame::Top, Frame::LineContinuous);
    m_font = font;
    m_frame = frame;
    m_foreground =   Qt::darkGray;
    m_background = Qt::white;
}

void Excel::TableStandardBody::placeData(Excel *pexcel, int x0, int y0, const QStringList &texts)
{    
    if(texts.count()>0){
        int rows = texts.count()/rect().width();
        if( rows==0) rows=1;
        for(int i=0;i<qMin(rows,rect().height());i++)
        {
            // put row
            pexcel->write(rect().row(i).toRange(), texts.mid(i*rect().width(),rect().width()) ,m_font);

        }
    }
    pexcel->drawFrame(rect().toRange(),m_frame);
}
