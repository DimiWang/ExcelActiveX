#ifndef AXOBJECT_H
#define AXOBJECT_H

//#ifdef RICC
    #include "WinSock2.h"
    #include <Windows.h>
//#else
//    #include <Windows.h>
//#endif



#include <QString>
#include <QVariant>
#include "qaxtypes.h"
#include <QThread>
#include <QDebug>


#define AxObjectType(TYPE,DATA) QString("@" #TYPE "(%1)").arg(DATA)




//*********************************************************************
//                              CLASS
//
//                      ActiveX client class
//*********************************************************************
class AxObject :public QThread
{
    Q_OBJECT

public:
    enum {Normal, Ignore, Abort, Retry};
    typedef int Class;



    AxObject(const QString &app_name,bool use_thread, bool log_errors = true);
    virtual ~AxObject();

    void Start();

    Class id() {return mp_object;} // activex ID

    bool isValid() const { return mp_object !=0;}

    // query object
    Class queryObject(Class parent, const QString &obj_pathname);

    //
    Class property_get_class(Class parent, const QString &obj_name);

    bool property_put(Class parent, const QString &prop,const QVariant &v);

    bool setProperty(Class parent, const QString &prop_path, const QVariant &v);

    bool property_put_variant(Class parent, const QString &prop,const VARIANT &v);

    bool setPropertyVariant(Class parent, const QString &prop_path, const VARIANT &v);

    bool property_get(Class parent, const QString &prop, QVariant *pvalue);

    bool property(Class parent, const QString &prop_path, QVariant *pvalue);
    void clearAbort();

    bool method_run(Class parent, const QString &method, QVariant *pres=&QVariant()
                                                , const QVariant &v1=QVariant()
                                                , const QVariant &v2=QVariant()
                                                , const QVariant &v3=QVariant()
                                                , const QVariant &v4=QVariant()
                                                , const QVariant &v5=QVariant()
                                                , const QVariant &v6=QVariant()
                                                , const QVariant &v7=QVariant()
                                                , const QVariant &v8=QVariant()
                                                );

    bool dynamicCall(Class parent,const QString &method_path, QVariant *pres=&QVariant()
                                                , const QVariant &v1=QVariant()
                                                , const QVariant &v2=QVariant()
                                                , const QVariant &v3=QVariant()
                                                , const QVariant &v4=QVariant()
                                                , const QVariant &v5=QVariant()
                                                , const QVariant &v6=QVariant()
                                                , const QVariant &v7=QVariant()
                                                , const QVariant &v8=QVariant()
                                                );


    Class object(const QString &name, Class parent_id=0);

    //void removeObject(const QString &name, int parent_id);
    bool objectExists(const QString &name, Class parent_id=0);
    void setErrorSlot(QObject *pobj, const char *);
    void release();

    void assignObject(const QString &obj_name, Class obj, Class parent_id=0, bool constant=false);
    void releaseObject(const QString &obj_name, Class parent_id=0 );
    void clearBag();

    void error(HRESULT hr, const QString &text, const QString &object_name =QString());

    Class findCachedObject(const QString &obj_name,Class parent_id=0) ;

    QString name() const { return m_object_name; }
    void run();

    static void QVariant_to_VARIANT(const QVariant &var,VARIANT &arg);    
    static void VARIANT_to_QVariant(const VARIANT &arg,QVariant &var);
    static void AxObject::clearVARIANT(VARIANT *var);

    static void QVariantList_to_2D_VARIANT(const QVariantList &list, int dimx,int dimy, VARIANT &arg);
    static void QStringList_to_2D_VARIANT(const QStringList &list, int dimx, int dimy, VARIANT &arg);

    int state() const { return m_state;}
    void finish() {m_finish=1;}


protected:
    HRESULT m_last_hr;

    bool m_use_thread;
    QString m_errorInfo;

    volatile bool m_finish;
    volatile bool m_trig;

    struct{
        int autoType;
        VARIANT *pvResult;
        IDispatch * pDisp;
        QString name;
        int cArgs;
        HRESULT result;
    }Request;

    struct{
        QString app_name;
        IDispatch *result;
    }Create;

signals:
    void signal_error(QString err_text, int *operation);

private:


    HRESULT AxRequest_Wk(int autoType
                             , VARIANT *pvResult
                             , IDispatch * pDisp
                             , const QString &name
                             , int cArgs =0);

    IDispatch *CreateObject(const QString &app_name);



    HRESULT AxRequest(int autoType
                             , VARIANT *pvResult
                             , IDispatch * pDisp
                             , const QString &name
                             , int cArgs =0);


    QString genSpecialKey(const QString &obj_name, int parent_id) const;


    volatile bool m_ignore;
    volatile int m_state;


    Class mp_object;
    QString m_object_name;
    QString m_app_name;


    QMap<QString , Class> obj_bag;
    QMap<QString , Class> obj_bagConst;

    QList<Class> obj_garbage;

    void addToObjectList(Class obj);
    void CleanUpArgs();



    VARIANT res;
    VARIANT Args[10];
};





#endif // AXOBJECT_H

