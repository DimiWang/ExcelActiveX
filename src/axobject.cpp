
//#define USE_WAITCONDITION
#define USE_THREAD




#include "axobject.h"
#include <QRegExp>
#include <QDebug>
#include <QWaitCondition>

#include "OleAuto.h"

static QList<SAFEARRAY *> safearray_garbage;



void ClearSafeArrayGarbage(){

//    foreach(SAFEARRAY *p,safearray_garbage){
//        SafeArrayDestroy(p);
//    }
    safearray_garbage.clear();
}



#ifdef USE_WAITCONDITION
QWaitCondition request;
#endif

// property in frmat Voltage, Signal,...
static const QRegExp prop_rx("([\\w\\d\\s]+)");
// property n format Voltage(0), Voltage("INPUT"), Signal("OUT"),...
static const QRegExp prop_rx_index("([\\w\\d]+)\\(([\"\\w\\d\\s,&:\\-\\/.#\\[\\]]+)\\)");



void AxObject::QVariant_to_VARIANT(const QVariant &var, VARIANT &arg)
{    
    static const QRegExp ax_convert("\\@([\\w\\d]+)\\(([\\w\\d\\.\\-\\,e]+)\\)");

    // ------------------------- Integer
    if(var.type() == QVariant::Int || var.type() == QVariant::UInt
            || var.type() == QVariant::LongLong ||var.type() == QVariant::Bool)
    {
        arg.vt = VT_I4;
        arg.lVal = var.toInt();
    }
    // ------------------------- Double
    else if(var.type() == QVariant::Double)
    {
        arg.vt = VT_R8;
        arg. dblVal = var.toDouble();
    }


    // -------------------------String
    else if(var.type() == QVariant::String )
    {
        const QRegExp ax_vbformat("(\\w+)\\((.*)\\)");
        // convert to any format
        if(var.toString().startsWith('@') && ax_convert.exactMatch(var.toString()))
        {
            // ------ format :
            if(ax_convert.capturedTexts().count() ==3 )
            {
                if(ax_convert.cap(1) == "I2"){
                }
                else if(ax_convert.cap(1) == "NULL"){
                    arg.vt =   VT_NULL;
                    arg.lVal = 0;
                }
                else if(ax_convert.cap(1) == "EMPTY"){
                    arg.vt = VT_EMPTY;
                    arg.lVal = 0;
                }
                else if(ax_convert.cap(1) == "I4"){
                    arg.vt = VT_I4;
                    arg.lVal = ax_convert.cap(2).toInt();
                }
                else if(ax_convert.cap(1) == "R4"){
                    arg.vt =VT_R4;
                    arg.dblVal = ax_convert.cap(2).toDouble();
                }
                else if(ax_convert.cap(1) == "R8"){
                    arg.vt = VT_R8;
                    arg.dblVal = ax_convert.cap(2).toDouble();
                }
                else if(ax_convert.cap(1) == "BSTR"){
                    arg.vt = VT_BSTR;//|VT_BYREF;
                    arg.bstrVal = SysAllocString((const OLECHAR*)var.toString().utf16());
                }
                else if(ax_convert.cap(1) == "DISPATCH"){
                    arg.vt = VT_DISPATCH;
                    arg.lVal = ax_convert.cap(2).toInt();
                }
                else if(ax_convert.cap(1) == "ERROR"){
                    arg.vt = VT_ERROR;
                    arg.lVal = ax_convert.cap(2).toInt();
                }
                else if(ax_convert.cap(1) == "BOOL"){
                    arg.vt =VT_BOOL;
                    arg.lVal = ax_convert.cap(2).toInt();
                }
                else if(ax_convert.cap(1) == "VARIANT"){
                    arg.vt = VT_VARIANT;
                    arg.lVal = ax_convert.cap(2).toInt();
                }
                else if(ax_convert.cap(1) == "UNKNOWN"){
                    arg.vt = VT_UNKNOWN;
                    arg.lVal = ax_convert.cap(2).toInt();
                }
                else if(ax_convert.cap(1) == "UI1" || ax_convert.cap(1) == "I1"){
                    arg.vt = VT_UI1;
                    arg.lVal = ax_convert.cap(2).toInt();
                }
                else {
                    return ;
                }
            }
        }
        // ---- visual basic format
        else if(ax_vbformat.exactMatch(var.toString()))
        {
            if(ax_vbformat.cap(1) == "CByte")
            {
                arg.vt =     VT_I1;
                arg.bVal = ax_vbformat.cap(2).toInt();
            }
            else if(ax_vbformat.cap(1) == "CInt")
            {
                arg.vt =     VT_I2;
                arg.bVal = ax_vbformat.cap(2).toInt();
            }
            else if(ax_vbformat.cap(1) == "CDbl")
            {
                arg.vt = VT_R4;
                arg.fltVal = ax_vbformat.cap(2).toDouble();
            }
            else if(ax_vbformat.cap(1) == "CLng")
            {
                arg.vt =     VT_I4;
                arg.bVal = ax_vbformat.cap(2).toInt();
            }
            else{
                arg.vt = VT_BSTR;//|VT_BYREF;
                arg.bstrVal = SysAllocString((const OLECHAR*)var.toString().utf16());
            }
        }
        // convert as normal string
        else{
            arg.vt = VT_BSTR;//|VT_BYREF;
            arg.bstrVal = SysAllocString((const OLECHAR*)var.toString().utf16());
        }
        //arg.pbstrVal = new BSTR(arg.bstrVal);
    }


    // --------------------- StringList
    else if(var.type() == QVariant::StringList)
    {


        const QStringList list = var.toStringList();
        const int count = list.count();
        SAFEARRAY *array = SafeArrayCreateVector(VT_BSTR, 0, count);
        for (LONG index = 0; index < count; ++index) {
            QString elem = list.at(index);
            BSTR bstr = QStringToBSTR(elem);
            SafeArrayPutElement(array, &index, bstr);
            SysFreeString(bstr);
        }

        arg.vt = VT_ARRAY|VT_BSTR;
        arg.parray = array;
        arg.pparray = new SAFEARRAY*(arg.parray);
        arg.vt |= VT_BYREF;
        safearray_garbage.append(array);
    }

    else if(var.type()==QVariant::Color)
    {
        QColor col = qvariant_cast<QColor>(var);
        arg.vt = VT_I4;
        arg.lVal = qRgba(col.blue(), col.green(), col.red(), 0x00);
        //                arg.plVal = new long(arg.lVal);
        //                arg.vt |= VT_BYREF;
    }

}

void AxObject::QVariantList_to_2D_VARIANT(const QVariantList &list, int dimx, int dimy,  VARIANT &arg)
{
    SAFEARRAYBOUND Bound[2];
    Bound[0].lLbound = 1;
    Bound[0].cElements = dimy;
    Bound[1].lLbound = 1;
    Bound[1].cElements = dimx;

    SAFEARRAY * psaData = SafeArrayCreate(VT_VARIANT, 2, Bound);    
    VARIANT HUGEP * pData = NULL;
    HRESULT hr = SafeArrayAccessData(psaData, (void HUGEP * FAR *)&pData);    
    if (SUCCEEDED(hr))
    {        
        for (int i = 0; i < list.count(); ++i,++pData)
        {            
            if(i >= (dimx*dimy) ) break;
            ::VariantInit(pData);            
            QVariant_to_VARIANT(list[i], *pData);            
            //pData->vt = VT_BSTR;
            //pData->bstrVal = SysAllocString((const OLECHAR*)l[i].utf16());
        }        
        SafeArrayUnaccessData(psaData);
    }

    arg.vt = VT_ARRAY | VT_VARIANT;
    arg.parray = psaData;    
    safearray_garbage.append(psaData);
}

void AxObject::QStringList_to_2D_VARIANT(const QStringList &list, int dimx, int dimy, VARIANT &arg)
{
    SAFEARRAYBOUND Bound[2];
    Bound[0].lLbound = 1;
    Bound[0].cElements = dimy;
    Bound[1].lLbound = 1;
    Bound[1].cElements = dimx;

    SAFEARRAY * psaData = SafeArrayCreate(VT_VARIANT, 2, Bound);

    VARIANT HUGEP * pData = NULL;
    HRESULT hr = SafeArrayAccessData(psaData, (void HUGEP * FAR *)&pData);
    if (SUCCEEDED(hr))
    {
        for (short i = 0; i < list.count(); ++i,++pData)
        {
            ::VariantInit(pData);
            pData->vt = VT_BSTR;
            pData->bstrVal = SysAllocString((const OLECHAR*)list[i].utf16());
        }

        SafeArrayUnaccessData(psaData);
    }

    arg.vt = VT_ARRAY | VT_VARIANT;
    arg.parray = psaData;
    safearray_garbage.append(psaData);
}




void AxObject::clearVARIANT(VARIANT *var)
{
    if (var->vt & VT_BYREF)
    {
        switch(var->vt)
        {
        case VT_BSTR|VT_BYREF:
            SysFreeString(*var->pbstrVal);
            delete var->pbstrVal;
            break;
        case VT_BOOL|VT_BYREF:
            delete var->pboolVal;
            break;
        case VT_I1|VT_BYREF:
            delete var->pcVal;
            break;
        case VT_I2|VT_BYREF:
            delete var->piVal;
            break;
        case VT_I4|VT_BYREF:
            delete var->plVal;
            break;
        case VT_INT|VT_BYREF:
            delete var->pintVal;
            break;
        case VT_UI1|VT_BYREF:
            delete var->pbVal;
            break;
        case VT_UI2|VT_BYREF:
            delete var->puiVal;
            break;
        case VT_UI4|VT_BYREF:
            delete var->pulVal;
            break;
        case VT_UINT|VT_BYREF:
            delete var->puintVal;
            break;
#if !defined(Q_OS_WINCE) && defined(_MSC_VER) && _MSC_VER >= 1400
        case VT_I8|VT_BYREF:
            delete var->pllVal;
            break;
        case VT_UI8|VT_BYREF:
            delete var->pullVal;
            break;
#endif
        case VT_CY|VT_BYREF:
            delete var->pcyVal;
            break;
        case VT_R4|VT_BYREF:
            delete var->pfltVal;
            break;
        case VT_R8|VT_BYREF:
            delete var->pdblVal;
            break;
        case VT_DATE|VT_BYREF:
            delete var->pdate;
            break;
        case VT_DISPATCH|VT_BYREF:
            (*var->ppdispVal)->Release();
            delete var->ppdispVal;
            break;
        case VT_ARRAY|VT_VARIANT|VT_BYREF:
        case VT_ARRAY|VT_UI1|VT_BYREF:
        case VT_ARRAY|VT_BSTR|VT_BYREF:
            SafeArrayDestroy(*var->pparray);
            delete var->pparray;
            break;
        case VT_VARIANT|VT_BYREF:
            delete var->pvarVal;
            break;
        }
        VariantInit(var);
    } else {
        VariantClear(var);
    }
}


void AxObject::VARIANT_to_QVariant(const VARIANT &arg, QVariant &var)
{        

    switch(arg.vt)
    {
    case VT_BOOL:
        var = (bool)(arg.lVal == 65535);
        break;
    case VT_EMPTY:
        break;
    case VT_NULL:
        var = 0;
        break;
    case VT_I2:
    case VT_I4:
    case VT_UI2:
    case VT_UI4:
        var = (unsigned int)arg.iVal;
        break;

    case VT_I1:
    case VT_UI1:
        var = (unsigned int)arg.bVal;
        break;


    case VT_I8:
    case VT_UI8:
    case VT_INT:
    case VT_UINT:
    case VT_DECIMAL:
        var = (int)arg.lVal;
        break;
    case VT_R4:
        var = (double)arg.fltVal;
        break;
    case VT_R8:
        var = (double)arg.dblVal;
        break;

    case VT_DISPATCH:
        var = (unsigned int)arg.lVal;
        break;

    case VT_BSTR:
        var = QString::fromUtf16(arg.bstrVal);
        break;

    case VT_ARRAY|VT_VARIANT:
    case VT_ARRAY|VT_VARIANT|VT_BYREF:
    {
        SAFEARRAY *array = 0;
        if ( arg.vt & VT_BYREF )
            array = *arg.pparray;
        else
            array = arg.parray;

        UINT cDims = array ? SafeArrayGetDim(array) : 0;
        switch(cDims) {
        case 1:
        {
            QVariantList list;

            long lBound, uBound;
            SafeArrayGetLBound( array, 1, &lBound );
            SafeArrayGetUBound( array, 1, &uBound );

            for ( long i = lBound; i <= uBound; ++i ) {
                VARIANT tmp;
                VariantInit( &tmp );
                SafeArrayGetElement( array, &i, &tmp );
                QVariant qvar;
                VARIANT_to_QVariant( tmp, qvar );
                clearVARIANT( &tmp );
                list << qvar;
            }

            var = list;
        }
            break;

        case 2:
        {
            QVariantList listList; //  a list of lists
            long dimIndices[2];

            long xlBound, xuBound, ylBound, yuBound;
            SafeArrayGetLBound(array, 1, &xlBound);
            SafeArrayGetUBound(array, 1, &xuBound);
            SafeArrayGetLBound(array, 2, &ylBound);
            SafeArrayGetUBound(array, 2, &yuBound);

            for (long x = xlBound; x <= xuBound; ++x) {
                QVariantList list;

                dimIndices[0] = x;
                for (long y = ylBound; y <= yuBound; ++y) {
                    VARIANT tmp;
                    VariantInit(&tmp);
                    dimIndices[1] = y;
                    SafeArrayGetElement(array, dimIndices, &tmp);
                    QVariant qvar;
                    VARIANT_to_QVariant(tmp, qvar);
                    clearVARIANT(&tmp);
                    list << qvar;
                }

                listList << QVariant(list);
            }
            var = listList;
        }
            break;
        default:
            var = QVariantList();
            break;
        }
    }
        break;

        /* types not supported*/

        /*
                VT_CY:
                VT_DATE:
        VT_ERROR:

        VT_VARIANT:
        VT_UNKNOWN:

               case  VT_VOID	:
               case  VT_HRESULT	:
               case  VT_PTR	:
               case  VT_SAFEARRAY	:
               case  VT_CARRAY	:
               case  VT_USERDEFINED	:
               case  VT_LPSTR	:
               case  VT_LPWSTR	:
               case  VT_RECORD	:
               case  VT_INT_PTR	:
               case  VT_UINT_PTR	:
               case  VT_FILETIME	:
               case  VT_BLOB	:
               case  VT_STREAM	:
               case  VT_STORAGE	:
               case  VT_STREAMED_OBJECT	:
               case  VT_STORED_OBJECT	:
               case  VT_BLOB_OBJECT	:
               case  VT_CF	:
               case  VT_CLSID	:
               case  VT_VERSIONED_STREAM	:
               case  VT_BSTR_BLOB	:
               case  VT_VECTOR	:
               case  VT_BYREF	:
               case  VT_RESERVED	:
               case  VT_ILLEGAL	:
               case  VT_ILLEGALMASKED	:
               case VT_TYPEMASK	:
                break;
*/
    }

}

\










/****************************************************************************
    * @function name:  AxRequest()
    * @param:
    *       int autoType -
    *       VARIANT *pvResult - invoke function result
    *       IDispatch * pDisp - object
    *       const QString &name - item name name
    *       int cArgs =0 - arguments count for function
    *       VARIANT *pArgs=0  - arguments (first at the end)
    * @description: Request function or item from windows
    * @return: ( void)
    ****************************************************************************/

HRESULT AxObject::AxRequest(int autoType
                            , VARIANT *pvResult
                            , IDispatch * pDisp
                            , const QString &name
                            , int cArgs )
{

    if(m_use_thread)
    {
        Request.autoType = autoType;
        Request.pvResult = pvResult;
        Request.pDisp  = pDisp;
        Request.name = name;
        Request.cArgs = cArgs;
        m_trig = true;
#ifdef USE_WAITCONDITION
        request.wakeAll();
#endif
        while(m_trig);
        return Request.result;
    }
    else
    {

        HRESULT hr=-1;
        if( m_state == AxObject::Normal)
        {
repeat_request:                        
            hr = AxRequest_Wk(autoType,pvResult,pDisp,name,cArgs);                        

            if( m_state == AxObject::Ignore)
            {
                m_state =  AxObject::Normal;
                hr = 0;
            }
            else if(m_state == AxObject::Retry)
            {
                m_state = AxObject::Normal;
                goto repeat_request;
            }            
        }          
        else{

        }

         ClearSafeArrayGarbage();
        return hr;
    }
}



HRESULT AxObject::AxRequest_Wk(int autoType
                               , VARIANT *pvResult
                               , IDispatch * pDisp
                               , const QString &name
                               , int cArgs )
{
    HRESULT hr = -1;
    EXCEPINFO exception;
    if(!pDisp) {   return 0;  }
    DISPPARAMS dp = { NULL, NULL, 0, 0};
    // [1] Call Invoke for object
    // Begin variable-argument list...

    do{
        // Variables used...

        DISPID dispidNamed = DISPID_PROPERTYPUT;
        DISPID dispID;
        LPOLESTR item= (LPOLESTR)name.utf16();

        // Get DISPID for name passed...
        hr = pDisp->GetIDsOfNames(IID_NULL, &item, 1, LOCALE_USER_DEFAULT, &dispID);

        if(FAILED(hr)) {
            error(hr, QString("Bad Item %1").arg(name));
            return -1;
        }

        // Build DISPPARAMS
        dp.cArgs = cArgs;
        dp.rgvarg = Args;
        // Handle special-case for property-puts!
        if(autoType & DISPATCH_PROPERTYPUT) {
            dp.cNamedArgs = 1;
            dp.rgdispidNamedArgs = &dispidNamed;
        }

        // Make the call!
        try{

            hr = pDisp->Invoke(dispID, IID_NULL, LOCALE_SYSTEM_DEFAULT, autoType, &dp, pvResult, &exception, NULL);
        }
        catch(...) {
            break;
        }
        if(FAILED(hr)) {
            break;
        }
        else{
            if(pvResult->vt == VT_DISPATCH){
                addToObjectList(pvResult->lVal);
            }
        }
        // End variable-argument section...
    }while(0);


    //[2] print  errors
    if(FAILED(hr))
    {
        switch(hr)
        {
        case DISP_E_BADPARAMCOUNT:
            error(hr, QString("%1 Bad parameter count").arg(name));
            break;

        case DISP_E_BADVARTYPE:
            error(hr, QString("%1 Bad variable type").arg(name));
            break;

        case DISP_E_MEMBERNOTFOUND:
            error(hr, QString("Member %1 not found ").arg(name));
            break;

        case DISP_E_NONAMEDARGS:
            error(hr, QString("%1 No named args").arg(name));
            break;

        case DISP_E_OVERFLOW:
            error(hr, QString("%1 Overflow").arg(name));
            break;

        case DISP_E_UNKNOWNNAME:
            error(hr, QString("Unknown name %1").arg(name));
            break;

        case DISP_E_PARAMNOTFOUND:
            error(hr, "Parameter no found",name);
            break;

        case DISP_E_TYPEMISMATCH:
            error(hr, "Type mismatch",name);
            break;

        case DISP_E_UNKNOWNINTERFACE :
            error(hr, "Unknown interface",name);
            break;

        case DISP_E_UNKNOWNLCID :
            error(hr, "Unknown LCID",name);
            break;

        case DISP_E_PARAMNOTOPTIONAL:
            error(hr, "Parameter not optional",name);
            break;

        case DISP_E_EXCEPTION:
        default:
            error(hr, QString("%1 Exception %4 (%2.%3)")
                  .arg(name)
                  .arg(QString::fromUtf16((const ushort*)exception.bstrSource))
                  .arg(QString::fromUtf16((const ushort*)exception.bstrDescription))
                  .arg(QString::number(hr,16)));
            break;
        }
    }

    return hr;
}

/****************************************************************************
    * @function name:  AxCreateObject()
    * @param:
    *      const QString &app_name - application name
    * @description: CreateObject(...) function
    * @return: (IDispatch *) returns pointer to application
    ****************************************************************************/

IDispatch * AxObject::CreateObject(const QString &app_name)
{    
    IDispatch *pInstance =0;
    // Get CLSID for our server...
    do
    {
        CLSID clsid;
        HRESULT hr = CLSIDFromProgID((LPCOLESTR)app_name.utf16(), &clsid);
        if(FAILED(hr)) {
            error(hr, QString("Error get CLSID(GUID) %1").arg(app_name));
            break;
        }
        // Start server and get IDispatch...

        hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void **)&pInstance);
        if(FAILED(hr)) {
            error(hr, QString("Error create instance %1").arg(app_name));
            break;
        }
    }
    while(0);
    m_object_name =  app_name;
    return pInstance;
}





void AxObject::release()
{ 
    foreach(Class obj_release, obj_garbage){
        while(((IDispatch*)obj_release)->Release()>0);
    }
    obj_garbage.clear();
}


void AxObject::clearBag()
{
    obj_bag.clear();
    foreach(const QString &name, obj_bagConst.keys())
    {
        obj_bag[name] = obj_bagConst[name];
    }    
}

void AxObject::clearAbort()
{
    this->m_trig = false;
    m_state = AxObject::Normal;
}

void AxObject::error(HRESULT hr, const QString &text, const QString &object_name)
{    
    if(hr != 0 )
    {
        int op = Normal;        
        emit signal_error(m_errorInfo+"\n"+ "Object :" + object_name + "\n" + text, &op);        
        m_state = op;        
    }
    m_last_hr = hr;
}

void AxObject::run()
{
    m_trig = true;
    Start();
    m_finish = false;
    m_trig = false;
    while(!m_finish)
    {
#ifdef USE_WAITCONDITION
        mutex.lock();
        request.wait(&mutex);
#else
        while(!m_trig) msleep(1);
#endif


        if(m_finish) break;
        Request.result = AxRequest_Wk(Request.autoType,Request.pvResult, Request.pDisp, Request.name, Request.cArgs);

        if(Request.result != 0 ) {
            // TODO
        }

#ifdef USE_WAITCONDITION
        mutex.unlock();
#endif
        m_trig = false;
        ClearSafeArrayGarbage();
    }
    exec();
}



void AxObject::CleanUpArgs()
{
    // free memory
    for(int i=0;i<10;i++)
    {
        if(Args[i].vt == VT_BSTR && Args[i].bstrVal) SysFreeString(Args[i].bstrVal);
        else if(Args[i].vt == (VT_ARRAY|VT_UI1) ) SafeArrayDestroy(Args[i].parray);
        Args[i].lVal=-1;
        VariantInit(&Args[i]);
    }

    if(res.vt == VT_BSTR && res.bstrVal) SysFreeString(res.bstrVal);
    else if(res.vt == (VT_ARRAY|VT_UI1) ) SafeArrayDestroy(res.parray);
    VariantInit(&res);
}




AxObject::Class AxObject::findCachedObject(const QString &obj_name, Class parent_id)
{    
    QString special_name = genSpecialKey(obj_name,parent_id);
    AxObject::Class result=0;
    if(obj_name== name() || obj_name == "Application") result = mp_object;
    else {
        if(obj_bag.contains(special_name)){
            return obj_bag[special_name];
        }
    }
    return result;
}


AxObject::AxObject(const QString &app_name, bool use_thread, bool log_errors)
{        
    m_state = AxObject::Normal;
    m_use_thread = use_thread;
    m_app_name = app_name;

    if(use_thread){
        m_trig = true;
        m_finish = false;
        start(QThread::NormalPriority);
        while(m_trig);
    }
    else{
        Start();
    }
}

AxObject::~AxObject()
{        
    ClearSafeArrayGarbage();
    release();
    m_finish = true;
    m_trig = true;

#ifdef USE_WAITCONDITION
    request.wakeAll();
#endif
    terminate();
}

void AxObject::Start()
{
    CoInitialize(0);
    mp_object = (Class)CreateObject(m_app_name);    
    // check if valid means !=0and add to constatnt table
    obj_bagConst[genSpecialKey("Application",id())] = 0;
}


/****************************************************************************
    * @function name:  querySubObject()
    * @param:
    *       QString name - subobject name
    * @return: (IDispatch *) pointer to subibject
    ****************************************************************************/


AxObject::Class AxObject::queryObject(AxObject::Class pobj, const QString &obj_pathname)
{
    do{
        QStringList l = QString(obj_pathname).split('.');
        for(int i=0;i<l.size();i++)
        {
            AxObject::Class obj_was_used = findCachedObject(l.at(i),pobj);
            if(obj_was_used ==0) {
                AxObject::Class new_obj = property_get_class(pobj,l.at(i));
                if(!new_obj) break;
                assignObject(l.at(i),new_obj,pobj);
                pobj = new_obj;
            }
            else{
                pobj = obj_was_used;
            }
        }
    }while(0);
    return pobj;
}


/****************************************************************************
    * @function name:  queryObject()
    * @param:
    *       IDispatch * pobj - object pointer
    *       QString name - subobject name
    * @return: (IDispatch *) pointer to subibject
    ****************************************************************************/
AxObject::Class AxObject::property_get_class(Class pobj, const QString &obj_name)
{
    VARIANT &index = Args[0];
    CleanUpArgs();
    HRESULT hr;

    Class object_handler =0;
    do
    {
        if(prop_rx.exactMatch(obj_name))
        {
            hr = AxRequest(DISPATCH_PROPERTYGET, &res, (IDispatch *)pobj, obj_name);
            if(res.vt == VT_DISPATCH)
                object_handler = (int)res.pdispVal;
        }
        else if(prop_rx_index.exactMatch(obj_name))
        {
            QString name = prop_rx_index.cap(1);
            QString tmp = prop_rx_index.cap(2);
            if(tmp.contains('\"'))
            {
                //is text
                tmp.replace('\"',"");
                QVariant_to_VARIANT(tmp,index);
                hr = AxRequest(DISPATCH_PROPERTYGET, &res, (IDispatch *)pobj, name, 1);
                if(res.vt == VT_DISPATCH)
                    object_handler = (int)res.pdispVal;
            }
            //is list of integer
            else if(tmp.contains(","))
            {
                QStringList l = tmp.split(",");
                int argn=0;
                foreach(const QString &par,l){
                    QVariant_to_VARIANT(par.toInt(),Args[l.size()-argn-1]);
                    argn++;
                }
                hr = AxRequest(DISPATCH_PROPERTYGET, &res, (IDispatch *)pobj, name, argn);
                if(res.vt == VT_DISPATCH)
                    object_handler = (int)res.pdispVal;
            }
            else //is integer
            {
                QVariant_to_VARIANT(tmp.toInt(),index);
                hr = AxRequest(DISPATCH_PROPERTYGET, &res, (IDispatch *)pobj, name, 1);
                if(res.vt == VT_DISPATCH)
                    object_handler = (int)res.pdispVal;
            }
        }
        else{
            error(hr, QString("Error parsing object in path (%1)").arg(obj_name));
        }

        if(!FAILED(hr))
        {
            assignObject(obj_name,object_handler,pobj);
        }
    }
    while(0);
    return object_handler;
}

bool AxObject::property_put(Class pobj, const QString &prop, const QVariant &v)
{
    m_errorInfo =  QString("property_put(pobj=%1,prop=%2)").arg(pobj).arg(prop);

    bool result = false;
    do{
        HRESULT hr=0;
        CleanUpArgs();

        QVariant_to_VARIANT(v,Args[0]);
        // just name
        if(prop_rx.exactMatch(prop))
        {
            hr = AxRequest(DISPATCH_PROPERTYPUT, &res, (IDispatch *)pobj, prop, 1);
        }
        //name with index
        else if(prop_rx_index.exactMatch(prop))
        {
            QString &name = prop_rx_index.cap(1);
            // index text
            QString tmp = prop_rx_index.cap(2);
            // index is a string
            if(tmp.contains('\"'))
            {
                tmp.replace('\"',"");
                QVariant_to_VARIANT(tmp,Args[1]);
                hr = AxRequest(DISPATCH_PROPERTYPUT, &res, (IDispatch *)pobj, name, 2);
            }
            //index is integer value
            else
            {
                QVariant_to_VARIANT(tmp.toInt(),Args[1]);
                hr = AxRequest(DISPATCH_PROPERTYPUT, &res, (IDispatch *)pobj, name, 2);
            }
        }
        result = !FAILED(hr);
    }while(0);
    return result;
}

bool AxObject::setProperty(Class parent, const QString &prop_path, const QVariant &v)
{
    bool result = false;
    Class pobj;

    if(parent==0) pobj= mp_object;
    else pobj = parent;

    do{

        QStringList l = QString(prop_path).split('.');
        for(int i=0;i<l.size();i++)
        {
            if(i == (l.size()-1))
            {
                const QString &item = l[i];
                result = property_put(pobj,item, v);
            }
            else{
                Class obj_was_used = findCachedObject(l.at(i),pobj);
                if(obj_was_used ==0) {
                    Class new_obj = property_get_class(pobj,l.at(i));
                    if(!new_obj) break;
                    assignObject(l.at(i),new_obj,pobj);
                    pobj = new_obj;
                }
                else pobj = obj_was_used;
            }
        }
    }while(0);
    return result;
}

bool AxObject::property_put_variant(AxObject::Class pobj, const QString &prop, const VARIANT &v)
{
    bool result = false;
    do{
        HRESULT hr=0;
        CleanUpArgs();

        Args[0]=v;
        // just name
        if(prop_rx.exactMatch(prop))
        {
            hr = AxRequest(DISPATCH_PROPERTYPUT, &res, (IDispatch *)pobj, prop, 1);
        }
        //name with index
        else if(prop_rx_index.exactMatch(prop))
        {
            QString &name = prop_rx_index.cap(1);
            // index text
            QString tmp = prop_rx_index.cap(2);
            // index is a string
            if(tmp.contains('\"'))
            {
                tmp.replace('\"',"");
                QVariant_to_VARIANT(tmp,Args[1]);
                hr = AxRequest(DISPATCH_PROPERTYPUT, &res, (IDispatch *)pobj, name, 2);
            }
            //index is integer value
            else
            {
                QVariant_to_VARIANT(tmp.toInt(),Args[1]);
                hr = AxRequest(DISPATCH_PROPERTYPUT, &res, (IDispatch *)pobj, name, 2);
            }
        }
        result = !FAILED(hr);
    }while(0);
    return result;
}


bool AxObject::setPropertyVariant(AxObject::Class pobj, const QString &prop_path, const VARIANT &v)
{    
    m_errorInfo =  QString("setPropertyVariant(pobj=%1,prop_path=%2)").arg(pobj).arg(prop_path);
    bool result = false;
    do{

        QStringList l = QString(prop_path).split('.');
        for(int i=0;i<l.size();i++)
        {
            if(i == (l.size()-1))
            {
                const QString &item = l[i];
                result = property_put_variant(pobj,item, v);
            }
            else{
                Class obj_was_used = findCachedObject(l.at(i),pobj);
                if(obj_was_used ==0) {
                    Class new_obj = property_get_class(pobj,l.at(i));
                    if(!new_obj) break;
                    assignObject(l.at(i),new_obj,pobj);
                    pobj = new_obj;
                }
                else pobj = obj_was_used;
            }
        }
    }while(0);
    return result;
}

/****************************************************************************
    * @function name:  AxPropertyPut()
    * @param:
    *       IDispatch * pobj -object pointer
    *       QString name - name of property
    *       const QVariant &v - value in varant type
    * @description: AxPropertyPut calls AxRequest with parameter of variant type
    * @return: (bool) if success returns true
    ****************************************************************************/


bool AxObject::property_get(Class pobj, const QString &prop, QVariant *pvalue)
{
    m_errorInfo =  QString("property_get(pobj=%1,prop=%2)").arg(pobj).arg(prop);

    bool result = false;
    do{
        HRESULT hr=0;
        CleanUpArgs();

        // just name
        if(prop_rx.exactMatch(prop))
        {
            hr = AxRequest(DISPATCH_PROPERTYGET, &res, (IDispatch *)pobj, prop);
        }
        //name with index
        else if(prop_rx_index.exactMatch(prop))
        {
            QString &name = prop_rx_index.cap(1);
            // index text
            QString tmp = prop_rx_index.cap(2);
            // index is a string
            if(tmp.contains('\"'))
            {
                tmp.replace('\"',"");
                QVariant_to_VARIANT(tmp, Args[0]);
                hr = AxRequest(DISPATCH_PROPERTYGET, &res, (IDispatch *)pobj, name, 1);
            }
            //index is integer value
            else
            {
                QVariant_to_VARIANT(tmp.toInt(), Args[0]);
                hr = AxRequest(DISPATCH_PROPERTYGET, &res, (IDispatch *)pobj, name, 1);
            }
        }
        if(!FAILED(hr))
        {
            result = true;
            if(pvalue) VARIANT_to_QVariant(res,*pvalue);
        }
    }while(0);
    return result;
}


bool AxObject::property(Class parent, const QString &prop_path, QVariant *pvalue)
{
    Class pobj;

    if(parent==0) pobj= mp_object;
    else pobj = parent;

    m_errorInfo =  QString("property(pobj=%1,prop_path=%2)").arg(pobj).arg(prop_path);

    bool result = false;
    do{
        QStringList l = QString(prop_path).split('.');
        for(int i=0;i<l.size();i++)
        {
            const QString &item = l[i];
            if(i==l.size()-1)
            {
                result = property_get(pobj,item,pvalue);
            }
            else{
                Class obj_was_used = findCachedObject(item,pobj);
                if(obj_was_used ==0) {
                    Class new_obj = property_get_class(pobj,item);
                    if(!new_obj) break;
                    assignObject(item,new_obj,pobj);
                    pobj = new_obj;
                }
                else pobj = obj_was_used;
            }
        }
    }while(0);
    return result;
}

bool AxObject::method_run(Class pobj, const QString &method, QVariant *pres
                          , const QVariant &v1
                          , const QVariant &v2
                          , const QVariant &v3
                          , const QVariant &v4
                          , const QVariant &v5
                          , const QVariant &v6
                          , const QVariant &v7
                          , const QVariant &v8)
{    
    bool result = false;
    m_errorInfo =  QString("method_run(pobj=%1,method=%2,v1=%3)").arg(pobj).arg(method).arg(v1.toString());
    do{
        HRESULT hr=0;
        CleanUpArgs();

        int j=0;
        int max_par_count =8;
        while(--max_par_count>=0)
        {
            switch(max_par_count)
            {
            case 0:
                if(!v1.isNull())  QVariant_to_VARIANT(v1,Args[j++]);
                break;
            case 1:
                if(!v2.isNull())  QVariant_to_VARIANT(v2,Args[j++]);
                break;
            case 2:
                if(!v3.isNull())  QVariant_to_VARIANT(v3,Args[j++]);
                break;
            case 3:
                if(!v4.isNull())  QVariant_to_VARIANT(v4,Args[j++]);
                break;
            case 4:
                if(!v5.isNull())  QVariant_to_VARIANT(v5,Args[j++]);
                break;
            case 5:
                if(!v6.isNull())  QVariant_to_VARIANT(v6,Args[j++]);
                break;
            case 6:
                if(!v7.isNull())  QVariant_to_VARIANT(v7,Args[j++]);
                break;
            case 7:
                if(!v8.isNull())  QVariant_to_VARIANT(v8,Args[j++]);
                break;
            }
        }
        // just name
        hr = AxRequest(DISPATCH_METHOD, &res, (IDispatch *)pobj, method, j)>=0;
        result = !FAILED(hr);
        if(result){
            if(pres) VARIANT_to_QVariant(res,*pres);
        }
    }while(0);

    return result;
}




bool AxObject::dynamicCall(AxObject::Class parent, const QString &method_path, QVariant *pres
                            ,const QVariant &v1
                            , const QVariant &v2
                            , const QVariant &v3
                            , const QVariant &v4
                            , const QVariant &v5
                            , const QVariant &v6
                            , const QVariant &v7
                            , const QVariant &v8)
{    
    bool result = false;
    m_errorInfo =  method_path;
    do{
        QStringList l = QString(method_path).split('.');
        Class pobj;

        if(parent==0) pobj= mp_object;
        else pobj = parent;

        for(int i=0;i<l.size();i++)
        {
            if(i==l.size()-1)
            {
                const QString &item = l[i];
                result = method_run(pobj, item, pres, v1, v2, v3, v4, v5, v6, v7, v8);
            }
            else{
                Class obj_was_used = findCachedObject(l.at(i),pobj);
                if(obj_was_used ==0) {
                    Class new_obj = property_get_class(pobj, l.at(i));
                    if(!new_obj) break;
                    assignObject(l.at(i),new_obj, pobj);
                    pobj = new_obj;
                }
                else pobj = obj_was_used;
            }
        }
    }while(0);
    return result;
}



void AxObject::assignObject(const QString &obj_name, AxObject::Class pobj, AxObject::Class parent_id, bool constant)
{        
    if(constant){
        obj_bagConst[genSpecialKey(obj_name,parent_id)]  = pobj;
    }
    addToObjectList(pobj);
}

AxObject::Class AxObject::object(const QString &obj_name, AxObject::Class parent_id)
{
    if(parent_id==0) parent_id = id();
    QString special_name = genSpecialKey(obj_name,parent_id);
    if(obj_bag.contains(special_name))
        return obj_bag[special_name];
    return 0;
}


bool AxObject::objectExists(const QString &obj_name, AxObject::Class parent_id)
{    
    return obj_bag.contains(genSpecialKey(obj_name,parent_id))
            || obj_bagConst.contains(genSpecialKey(obj_name,parent_id));
}

void AxObject::setErrorSlot(QObject *pobj, const char *slot)
{
    if(!m_use_thread){
        connect(this, SIGNAL(signal_error(QString,int*)),pobj,slot);
    }
    else{
        connect(this, SIGNAL(signal_error(QString,int*)),pobj,slot,Qt::BlockingQueuedConnection);
    }
}


void AxObject::addToObjectList(AxObject::Class obj)
{
    if(obj != 0 && !obj_garbage.contains((int)obj))
        obj_garbage.insert(0,(int)obj);
}


void AxObject::releaseObject(const QString &obj_name, AxObject::Class parent_id)
{
    Class obj = findCachedObject(obj_name,parent_id);
    if(obj){
        ((IDispatch*)obj)->Release();
        obj_bag.remove(genSpecialKey(obj_name,parent_id));
        obj_garbage.removeAll(obj);
    }
}


QString AxObject::genSpecialKey(const QString &obj_name, int parent_id) const
{
    return QString("%1::%2").arg(parent_id).arg(obj_name);
}


