// ControlWPS.h: interface for the CControlWPS class.
//
//////////////////////////////////////////////////////////////////////

#if !defined(_CONTROLWPS_H_)
#define _CONTROLWPS_H_

#include <atlbase.h>
#include "app_event.h"

class CControlWPS  
{
public:
	CControlWPS();
	~CControlWPS();
public:
	void OpenWPS();
	void CloseWPS();
private:
	IDispatch* m_pIDisp;
	WPS::_ApplicationPtr m_spApp;///wps��Ӧ�ó���
	KAppEvent m_AppEvent;///Ӧ�ó������Ӧ�¼�
private:
	HRESULT GetProperty(IDispatch *pDisp, LPCOLESTR lpsz, VARIANT *pVar);
	HRESULT PutProperty(IDispatch *pDisp, LPCOLESTR lpsz, VARIANT *pVar);
	HRESULT Invoke0(IDispatch *pDisp, LPCOLESTR lpszName, VARIANT *pvarRet);
	HRESULT InvokeN(IDispatch *pDisp, LPCOLESTR lpszName, VARIANT *varParams, int nParams, VARIANT *pvarRet);
};

#endif // !defined(_CONTROLWPS_H_)
