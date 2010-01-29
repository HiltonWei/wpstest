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

// 启动WPS文字，并添加居中文字，以及插入一幅图片
void CControlWPS::OpenWPS()
{
	USES_CONVERSION;

	CLSID clsid = {0}; 

	// 获取WPS文字的 CLSID
	HRESULT hr = CLSIDFromProgID(L"wps.application", &clsid);
	if(FAILED(hr))
		return;
	
	// 通过CLSID启动WPS
	hr = CoCreateInstance(clsid, NULL, CLSCTX_LOCAL_SERVER, IID_IDispatch, (void**)&m_pIDisp);
	
	if(FAILED(hr))
		return;

	// 使WPS可见
	VARIANT var = {0};
	var.vt = VT_BOOL;
	var.boolVal = VARIANT_TRUE;
	PutProperty(m_pIDisp, L"Visible", &var);
/**/
	//装载事件
	m_spApp = m_pIDisp;///主程序
	m_AppEvent.init(m_spApp);
	m_AppEvent.EventAdvise(m_spApp);
	
	

	// 获取Documents集合
	VARIANT varDocuments = {0};
	GetProperty(m_pIDisp, L"Documents", &varDocuments);

	// 在Documents中新建一篇文档
	VARIANT varDocument = {0};
	Invoke0(varDocuments.pdispVal, L"Add", &varDocument);

	
	// 获取Selection对象
	VARIANT varSelection = {0};
	GetProperty(m_pIDisp, L"Selection", &varSelection);

	// 获取Selection的Range对象
	VARIANT varRange = {0};
	GetProperty(varSelection.pdispVal, L"Range", &varRange);

	// 获取ParagraphFormat对象，以便设置对齐方式
	VARIANT varParagraphFormat = {0};
	GetProperty(varSelection.pdispVal, L"ParagraphFormat", &varParagraphFormat);

	// 设置文字居中对齐
	var.vt = VT_I4;
	var.lVal = 1;
	PutProperty(varParagraphFormat.pdispVal, L"Alignment", &var);

	// 插入文字，该文字是居中显示的
	var.vt = VT_BSTR;
	var.bstrVal = ::SysAllocString(A2OLE("hello,world\n"));
	PutProperty(varRange.pdispVal, L"Text", &var);

	// 获取Shapes集合对象
	VARIANT varShapes = {0};
	GetProperty(varDocument.pdispVal, L"Shapes", &varShapes);

	// 设置参数，该参数将传入Shapes的AddPicture方法中
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
	
	// 插入图片
	InvokeN(varShapes.pdispVal, L"AddPicture", varParams, 8, &var);
}

// 关闭WPS
void CControlWPS::CloseWPS()
{
	CComVariant varParams[3];
	varParams[0].vt = VT_BOOL;
	varParams[0].boolVal = VARIANT_FALSE;
	varParams[1].vt = VT_BOOL;
	varParams[1].boolVal = VARIANT_FALSE;
	varParams[2].vt = VT_BOOL;
	varParams[2].boolVal = VARIANT_FALSE;

	// 首先调用WPS.Application的Quit方法，并且不保存修改
	m_AppEvent.EventUnadvise(m_spApp);
	InvokeN(m_pIDisp, L"Quit", varParams, 3, NULL);
	if(m_pIDisp != NULL)
	{
		// 释放对象
		m_pIDisp->Release();
		m_pIDisp = NULL;
	}
}

// 执行没有参数的方法
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

// 执行含有N个参数的方法
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

// 获取属性值
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

// 设置属性值
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