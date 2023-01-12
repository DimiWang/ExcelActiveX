/**
 * @file:exceldata.h   -
 * @description: This is Excel communication class.
 * @project: BENCH OnSemiconductor
 * @date: 2014\11\27 13-51-10
 *
 */


#ifndef EXCEL_H
#define EXCEL_H

#include <QObject>
#include "axobject.h"
#include <QStringList>
#include <QString>
#include "excelenums.h"

#include <QRect>

class Excel : public QObject
{
    Q_OBJECT
public:

    class Frame
    {
      public:
        Frame(){}
        enum {
            DrawBottom = 1,
            DrawLeft = 2,
            DrawRight = 4,
            DrawTop = 8,
            DrawHorizontal = 16,
            DrawVertical =32
        };


        enum {
            Bottom,
            Left,
            Right,
            Top,
            Horizontal,
            Vertical
        };// line

        enum {
            LineNone ,
            LineContinuous,
            LineDouble
        };//style

        enum {
               WidthHairline,
               WidthMedium ,
               WidthThick ,
               WidthThin
            };// width

        void setDrawLines(quint32 lines){
            m_draw_lines = lines;
        }
        void setStyle(int line ,int style){
            m_styles[line] = style;
        }

        void setWidth(int line, int width){
            m_widths[line] = width;
        }
        int style(int line) const{
            return m_styles[line];
        }
        int width(int line)const{
            return m_widths[line];
        }

        bool drawLine(int line) const{
            return m_draw_lines&(1<<line);
        }


        Frame(quint32 draw_lines){
            setDrawLines(draw_lines);
            for(int i=0;i<6;i++){
                m_styles[i] = LineContinuous;
                m_widths[i] = WidthThin;
            }
        }

        static const Frame doubleFrame(){
            Frame frame(DrawBottom|DrawLeft|DrawRight|DrawTop);

            frame.setStyle(Bottom, LineDouble);
            frame.setStyle(Top, LineDouble);
            frame.setStyle(Left, LineDouble);
            frame.setStyle(Right, LineDouble);
            return frame;
        }

        static const Frame singleFrame(){
            Frame frame(DrawBottom|DrawLeft|DrawRight|DrawTop);

            frame.setStyle(Bottom, LineContinuous);
            frame.setStyle(Top, LineContinuous);
            frame.setStyle(Left, LineContinuous);
            frame.setStyle(Right, LineContinuous);
            return frame;
        }

        static const Frame thickFrame(){
            Frame frame(DrawBottom|DrawLeft|DrawRight|DrawTop);

            frame.setStyle(Bottom, LineContinuous);
            frame.setStyle(Top, LineContinuous);
            frame.setStyle(Left, LineContinuous);
            frame.setStyle(Right, LineContinuous);

            frame.setWidth(Bottom, WidthThick);
            frame.setWidth(Top, WidthThick);
            frame.setWidth(Left, WidthThick);
            frame.setWidth(Right, WidthThick);
            return frame;
        }

    private:
        quint32 m_draw_lines;
        int m_styles[6];
        int m_widths[16];

    };

    class Cell
    {
    public:
        Cell(){m_x=0;m_y=0;}
        Cell(const QString &range){
            QRegExp rx("([a-zA-Z]+)([0-9]+)");
            if(rx.exactMatch(range.toUpper())){
                m_x = 0;
                for(int i=rx.cap(1).count()-1;i>=0;i--)
                {
                    if(i==rx.cap(1).count()-1) m_x = rx.cap(1)[i].toAscii()-'A';
                    else m_x += ((rx.cap(1)[i].toAscii()-'A')+1)*26;
                }
                m_y = rx.cap(2).toInt()-1;
            }
        }

        Cell(int x, int y)
        {
            if(m_x>=0) this->m_x = x;
            else  this->m_x = 0;
            if(m_y>=0) this->m_y = y;
            else  this->m_y = 0;
        }
        int x() const {return m_x;}
        int y() const {return m_y;}
        void setX(int x) {m_x=x;}
        void setY(int y) {m_y=y;}
        QString toRange(bool fixed = false) const
        {
            return Cell_To_Name(*this,fixed);
        }

        bool isEmpty() const
        {
            return x()==-1 && y()==-1;
        }
        int xlRow() const
        {
            return y()+1;
        }
        int xlCol() const
        {
            return x()+1;
        }

    private:
        int m_x;
        int m_y;

    };

    class Rect
    {
    public:
        Rect(){}
        Rect(int x,int y,int width,int height)
        {
            if(width>0 && height >0 && x>=0 && y>=0)
            {
                this->m_p1.setX(x);
                this->m_p1.setY(y);
                this->m_p2.setX(x+width);
                this->m_p2.setY(y+height);
            }
        }
        int x() const  { return m_p1.x();}
        int y()const  { return m_p1.y();}
        int width() const {
            return m_p2.x() - m_p1.x();
        }
        int height() const {
            return m_p2.y() - m_p1.y();
        }
        void setX(int x) {
            m_p1.setX(x);
        }
        void setY(int y) {
            m_p1.setY(y);
        }
        void setWidth(int w) {
            m_p2.setX(w+m_p1.x());
        }
        void setHeight(int h) {
            m_p2.setY(h+m_p1.y());
        }
        QString toRange(bool fixed=false) const {
            return Rect_To_Range(*this,fixed);
        }
        bool isEmpty() const {
            return m_p1.isEmpty() && m_p2.isEmpty();
        }
        Cell p1() const {return m_p1;}
        Cell p2() const {return m_p2;}

        Rect row( int i){
            Rect r = *this;
            if(i<this->height())    {
                r.setY(i+y());
                r.setHeight(1);
            }
            return r;
        }

        Rect column(int i){
            Rect r = *this;
            if(i<this->width())    {
                r.setX(i+x());
                r.setWidth(1);
            }
            return r;
        }

        Cell cell(int x, int y){
            Cell c;
            if(x<width() && y<height()){
                c.setX(x+this->x());
                c.setY(y+this->y());
            }
            return c;
        }

    private:
        Cell m_p1;
        Cell m_p2;
    };


    class DataArea
    {
        public:
            DataArea(){
                m_x = -1;  m_y = -1;
                m_width = -1; m_height = -1;                
            }
            ~DataArea(){}
            virtual void placeData(Excel *pexcel,  int x, int y, const QStringList &data)=0;
            int height() const {return m_height;}
            int width() const {return m_width;}
            int x() const {return m_x;}
            int y() const {return m_y;}
            void setX(int x) {m_x=x;}
            void setY(int y) {m_y=y;}
            void setHeight(int height) {m_height = height;}
            void setWidth(int width) {m_width = width;}
            Rect rect() { return Rect(x(),y(),width(),height());}

        protected:
            int m_width;
            int m_height;
            int m_x;
            int m_y;
    };


    class Table
    {
    public:
        Table();
        Table(Excel*pexcel, const Rect &rect);
        ~Table(){
            if(mp_dataArea != 0)
                delete mp_dataArea;
            if(mp_headerArea != 0)
                delete mp_headerArea;
            mp_dataArea =0;
            mp_headerArea=0;
        }

        void setDataArea(DataArea *dataArea);
        void setHeaderArea(DataArea *dataArea);

        void setHeaderData(const QStringList &data);
        void setTableData(const QStringList &data);

        bool appendDataRow(const QStringList &data);
        int width() { return m_rect.width();}
        int height() {return m_rect.height();}

        Rect rect() {return m_rect;}
        Rect dataRect() {return mp_dataArea->rect();}
        Rect headerRect() {return mp_headerArea->rect();}
        AxObject::Class dataRange() {return m_dataRange;}
        int rowsCount() {return m_rows_count;}
        void setRowsCount(int rc){ m_rows_count=rc;}


    private:
        Excel *mp_excel;
        AxObject::Class m_dataRange;// excel id
        Rect m_rect;
        DataArea *mp_headerArea;
        DataArea *mp_dataArea;
        int m_rows_count;  //current text data added to rows
        int m_height;// rows
        int m_width; // columns
        QFont m_font;
    };



    class TableHeader1Line:public DataArea
    {
        public:
        TableHeader1Line();
        virtual void placeData(Excel *pexcel, int x, int y, const QStringList &data);
        protected:
            QColor m_foreground;
            QColor m_background;
            QFont m_font;
            Frame m_frame;
    };

    class TableStandardBody:public DataArea
    {
        public:
        TableStandardBody();
        virtual void placeData(Excel *pexcel, int x, int y, const QStringList &data);
        protected:
            QColor m_foreground;
            QColor m_background;
            QFont m_font;
            Frame m_frame;
    };



    enum ChartType{
        Chart_ScatterLine, Chart_Line
    };


    enum ScaleType{
        Scale_Linear, Scale_Logarithmic
    };


    struct Chart{
        Rect cellDataRange;
        QString title;
        QRect rect;
        ChartType type;
        ScaleType xScaleType;
        ScaleType yScaleType;
        QString xAxis;
        QString yAxis;
        bool legendVisible;
        bool minorGridLines;
        bool majorGridLines;
    };





    explicit Excel(const QString filename, bool use_thread=false,bool autosave=true);
    ~Excel();
    static bool validName(const QString &name);
    void setErrorSlot(QObject *pobj, const char *slot);
    QString fileName();
    /* opens excel document*/
    bool open();
    bool activate();
    void release();
    void clearAbort();
    bool isAborted();
    bool isReadOnly();
    /* returns if document is opened*/
    bool isOpen() const;
    bool setZoom(int val);
    int zoom();
    /* add sheet with name to opened workbook*/
    bool addSheet(const QString &sheetname);
    /* sets current sheet name*/
    bool setCurrentSheet(const QString &sheetname);
    bool setCellHint(qint32 row, qint32 col, const QString &text);
    AxObject::Class currentSheet();
    AxObject::Class currentWorkBook();

    int sheetsCount();
    QStringList sheetsList();
    /* removes sheet sheetname*/
    bool removeSheet(const QString &sheetname);
    bool removeSheet( int sheetnumber);
    bool removeSheetsList(const QStringList &sheets);

    /* write data to cell*/
    bool write(Excel::Cell cell, const QVariant &data, const QFont &font= QFont() );
    bool write(qint32 row, qint32 col, const QVariant &data, const QFont &font= QFont()  );
    bool write(AxObject::Class range, qint32 row, qint32 col, const QVariant &data, const QFont &font= QFont()  );
    bool write(const QString &range, const QStringList &l, const QFont &font=QFont());
    bool mergeRange(const QString &range, bool on=true);
    //reads range of data
    bool readRange(const QString &range, QVariantList *presult);
    bool writeRange(const QString &range, QVariantList l);
    enum {AlignHorizontalCenter, AlignHorizontalRight,AlignHorizontalLeft
          , AlignVerticalTop,AlignVerticalCenter, AlignVerticalBottom };

    bool setRangeAlignment(const QString &range, unsigned int align);
    // shows current cell
    bool cellVisible(int row, int col);

    /* reads data from cell*/
    bool read(qint32 row, qint32 col, QVariant &data);
    void setAutoSaveOn(bool on);
    bool autoSaveOn() const {return m_autosave;}
    bool drawFrame(const QString &range, const Frame &f);
    bool drawFrame(const Rect &rect, const Frame &f);
    /* sets color of cell*/
    bool setColor(qint32 row, qint32 col, const QColor background, const QColor foreground);
    bool setColor(const QString &range, const QColor background, const QColor foreground);
    bool setColor(const Rect &rect, const QColor background, const QColor foreground);
    /* gets color of cell */
    bool color(qint32 row, qint32 col, QColor &background, QColor &foreground);
    /* sets visible workbook*/
    bool setVisible(bool visible);
    bool visible();
    bool valid();
    bool resizeCells(double size);
    /* closes workbook*/
    void close(void);
    /* saves workbook to file*/
    bool save();
    bool saveAs(const QString &);

    QStringList namedRanges();

    bool test();
    static QString Cell_To_Name(const Cell &cell, bool fixed=false);
    static QString Rect_To_Range(const Rect &rect, bool fixed=false);
    static Cell Name_To_Cell(const QString &xl_name);
    static Rect Range_To_Rect(const QString &xl_range);
    static bool Range_Is_Valid(const QString &range);

    Excel::Table *CreateTable(const Rect &range
                                , const QStringList &headers
                                , const QStringList &data = QStringList());

    Excel::Table *CreateTable(const Rect &range
                                , DataArea *tableHeader
                                , DataArea *tableData =0);

    bool AppendRow(Table *ptable, const QStringList &data);
    bool SetDataToColumn(Table *ptable, const QVariantList &data, int column);
    bool SetDataToRange(const Excel::Rect &rect, QVariantList data);
    bool SetDataToRow(Table *ptable, const QVariantList &data, int row);

    int width(const QString &range);
    int height(const QString &range);

    void setUpdatesOn(bool on);
    void setCalculation(bool on);
    void setScreenUpdate(bool on);
    void recalculate();
    static int version();
    // *****************charts************************

    // charts
    AxObject::Class CreateChart(const Chart &chart);

    AxObject *object() {return mp_exlObject;}

    bool SetChartData(AxObject::Class chart, const QString &range);

private:    
    AxObject *mp_exlObject;

    AxObject::Class mp_currentSheet;
    AxObject::Class mp_currentWorkBook;

    QString m_filename;
    QString m_sheetname;
    bool m_opened;
    bool m_saved;
    bool m_autosave;
    bool m_badFile;
    bool m_updatesOn;
signals:

public slots:

};

//class AxBag:public QMap<QString , AxObject::Class>
//{


//};


#endif // EXCELDATA_H

