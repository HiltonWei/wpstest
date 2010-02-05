#pragma once
class __declspec(uuid("{7209C677-015F-4D67-9272-E5ED03ED8F81}")) Cwpseventplugins;

using namespace AddInDesignerObjects;

class Cwpseventplugins : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<Cwpseventplugins, &__uuidof(Cwpseventplugins)>,
	public IDispatchImpl<_IDTExtensibility2, &IID__IDTExtensibility2, &LIBID_AddInDesignerObjects>
{
protected:
	WPS::_ApplicationPtr m_spApp;///wps的应用程序
	KAppEvent m_AppEvent;///应用程序的响应事件
public:
	DECLARE_REGISTRY_RESOURCEID(IDR_WPSCOMADDONS)
	DECLARE_PROTECT_FINAL_CONSTRUCT()

	BEGIN_COM_MAP(Cwpseventplugins)
		COM_INTERFACE_ENTRY(IDispatch)
		COM_INTERFACE_ENTRY(_IDTExtensibility2)
	END_COM_MAP()

	Cwpseventplugins()
	{
	}

	~Cwpseventplugins()
	{
	}

public:
	STDMETHOD(OnConnection)(IDispatch * Application, 
		ext_ConnectMode ConnectMode, IDispatch * AddInInst, SAFEARRAY * * custom)
	{
		m_spApp = Application;///主程序
		// 设置程序事件
		m_AppEvent.init(m_spApp);
		m_AppEvent.EventAdvise(m_spApp);
		return S_OK;
	}

	STDMETHOD(OnDisconnection)(ext_DisconnectMode RemoveMode, SAFEARRAY * * custom)
	{
		m_AppEvent.EventUnadvise(m_spApp);
		return S_OK;
	}
	STDMETHOD(OnAddInsUpdate)(SAFEARRAY * * custom)
	{
		return S_OK;
	}
	STDMETHOD(OnStartupComplete)(SAFEARRAY * * custom)
	{
		return S_OK;
	}
	STDMETHOD(OnBeginShutdown)(SAFEARRAY * * custom)
	{
		return S_OK;
	}
};

