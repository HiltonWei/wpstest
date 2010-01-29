// ControlWPS.cpp: implementation of the CControlWPS class.
//
//////////////////////////////////////////////////////////////////////

#include "stdafx.h"
#include "ControlWPS.h"
//////////////////////////////////////////////////////////////////////
// Construction/Destruction
//////////////////////////////////////////////////////////////////////

CControlWPS::CControlWPS() :
m_pIDisp(NULL)
{
	CoInitialize(NULL);
}

CControlWPS::~CControlWPS()
{
	CloseWPS();
	CoUninitialize();
}

// ����WPS���֣�����Ӿ������֣��Լ�����һ��ͼƬ
void CControlWPS::OpenWPS()
{
	USES_CONVERSION;

	CLSID clsid = {0}; 

	// ��ȡWPS���ֵ� CLSID
	HRESULT hr = CLSIDFromProgID(L"wps.application", &clsid);
	if(FAILED(hr))
		return;
	
	// ͨ��CLSID����WPS
	hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&m_pIDisp);
	
	if(FAILED(hr))
		return;

	// ʹWPS�ɼ�
	VARIANT var = {0};
	var.vt = VT_BOOL;
	var.boolVal = VARIANT_TRUE;
	PutProperty(m_pIDisp, L"Visible", &var);
/**/
	//װ���¼�
	m_spApp = m_pIDisp;///������
	m_AppEvent.init(m_spApp);
	m_AppEvent.EventAdvise(m_spApp);
	
	

	// ��ȡDocuments����
	VARIANT varDocuments = {0};
	GetProperty(m_pIDisp, L"Documents", &varDocuments);

	// ��Documents���½�һƪ�ĵ�
	VARIANT varDocument = {0};
	Invoke0(varDocuments.pdispVal, L"Add", &varDocument);

	
	// ��ȡSelection����
	VARIANT varSelection = {0};
	GetProperty(m_pIDisp, L"Selection", &varSelection);

	// ��ȡSelection��Range����
	VARIANT varRange = {0};
	GetProperty(varSelection.pdispVal, L"Range", &varRange);

	// ��ȡParagraphFormat�����Ա����ö��뷽ʽ
	VARIANT varParagraphFormat = {0};
	GetProperty(varSelection.pdispVal, L"ParagraphFormat", &varParagraphFormat);

	// �������־��ж���
	var.vt = VT_I4;
	var.lVal = 1;
	PutProperty(varParagraphFormat.pdispVal, L"Alignment", &var);

	// �������֣��������Ǿ�����ʾ��
	var.vt = VT_BSTR;
	var.bstrVal = ::SysAllocString(A2OLE("hello,world\n"));
	PutProperty(varRange.pdispVal, L"Text", &var);

	// ��ȡShapes���϶���
	VARIANT varShapes = {0};
	GetProperty(varDocument.pdispVal, L"Shapes", &varShapes);

	// ���ò������ò���������Shapes��AddPicture������
	CComVariant varParams[8];
	varParams[0].vt = VT_BSTR;			//	Anchor
	varParams[0].bstrVal = ::SysAllocString(A2OLE(""));
	varParams[1].vt = VT_I4;			//	Height
	varParams[1].lVal = 60;				
	varParams[2].vt = VT_I4;			//	Width
	varParams[2].lVal = 148;
	varParams[3].vt = VT_I4;			//	Top
	varParams[3].lVal = 50;
	varParams[4].vt = VT_I4;			//	Left
	varParams[4].lVal = 100;
	varParams[5].vt = VT_BOOL;			//	SaveWithDocument
	varParams[5].boolVal = VARIANT_TRUE;
	varParams[6].vt = VT_BOOL;			//	LinkToFile
	varParams[6].boolVal = VARIANT_FALSE;
	varParams[7].vt = VT_BSTR;			//	FileName
	varParams[7].bstrVal = ::SysAllocString(A2OLE("http://img.kingsoft.com/publish/kingsoft/images/gb/sy/logo.gif"));
	
	// ����ͼƬ
	InvokeN(varShapes.pdispVal, L"AddPicture", varParams, 8, &var);
}

// �ر�WPS
void CControlWPS::CloseWPS()
{
	CComVariant varParams[3];
	varParams[0].vt = VT_BOOL;
	varParams[0].boolVal = VARIANT_FALSE;
	varParams[1].vt = VT_BOOL;
	varParams[1].boolVal = VARIANT_FALSE;
	varParams[2].vt = VT_BOOL;
	varParams[2].boolVal = VARIANT_FALSE;

	// ���ȵ���WPS.Application��Quit���������Ҳ������޸�
	m_AppEvent.EventUnadvise(m_spApp);
	InvokeN(m_pIDisp, L"Quit", varParams, 3, NULL);
	if(m_pIDisp != NULL)
	{
		// �ͷŶ���
		m_pIDisp->Release();
		m_pIDisp = NULL;
	}
}

// ִ��û�в����ķ���
HRESULT CControlWPS::Invoke0(IDispatch *pDisp, LPCOLESTR lpszName, VARIANT *pvarRet)
{
	if(pDisp == NULL)
		return E_FAIL;

	HRESULT hr;
	DISPID dispid;

	hr = pDisp->GetIDsOfNames(IID_NULL, (LPOLESTR*)&lpszName, 1, LOCALE_USER_DEFAULT, &dispid);
	if (SUCCEEDED(hr))
	{
		DISPPARAMS dispparams = { NULL, NULL, 0, 0};
		return pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, pvarRet, NULL, NULL);
		
	}
	return hr;
}

// ִ�к���N�������ķ���
HRESULT CControlWPS::InvokeN(IDispatch *pDisp, LPCOLESTR lpszName, VARIANT *varParams, int nParams, VARIANT *pvarRet)
{
	if(pDisp == NULL)
		return E_FAIL;

	HRESULT hr;
	DISPID dispid;
	
	hr = pDisp->GetIDsOfNames(IID_NULL, (LPOLESTR*)&lpszName, 1, LOCALE_USER_DEFAULT, &dispid);
	if (SUCCEEDED(hr))
	{
		DISPPARAMS dispparams = { varParams, NULL, nParams, 0};
		return pDisp->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, pvarRet, NULL, NULL);
		
	}
	return hr;

}

// ��ȡ����ֵ
HRESULT CControlWPS::GetProperty(IDispatch *pDisp, LPCOLESTR lpsz, VARIANT *pVar)
{
	if(pDisp == NULL)
		return E_FAIL;

	DISPID dwDispID;

	pDisp->GetIDsOfNames(IID_NULL, (LPOLESTR*)&lpsz, 1, LOCALE_USER_DEFAULT, &dwDispID);
	DISPPARAMS dispparams = {NULL, NULL, 0, 0};
	
	return pDisp->Invoke(dwDispID, IID_NULL,
		LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET,
		&dispparams, pVar, NULL, NULL);
}

// ��������ֵ
HRESULT CControlWPS::PutProperty(IDispatch *pDisp, LPCOLESTR lpsz, VARIANT *pVar)
{
	if(pDisp == NULL)
		return E_FAIL;

	DISPID dwDispID;

	pDisp->GetIDsOfNames(IID_NULL, (LPOLESTR*)&lpsz, 1, LOCALE_USER_DEFAULT, &dwDispID);
	DISPPARAMS dispparams = {NULL, NULL, 1, 1};
	dispparams.rgvarg = pVar;
	DISPID dispidPut = DISPID_PROPERTYPUT;
	dispparams.rgdispidNamedArgs = &dispidPut;
	
	if (pVar->vt == VT_UNKNOWN || pVar->vt == VT_DISPATCH || 
		(pVar->vt & VT_ARRAY) || (pVar->vt & VT_BYREF))
	{
		HRESULT hr = pDisp->Invoke(dwDispID, IID_NULL,
			LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUTREF,
			&dispparams, NULL, NULL, NULL);
		if (SUCCEEDED(hr))
			return hr;
	}
	
	return pDisp->Invoke(dwDispID, IID_NULL,
		LOCALE_USER_DEFAULT, DISPATCH_PROPERTYPUT,
		&dispparams, NULL, NULL, NULL);
}