

#ifndef app_event_h__
#define app_event_h__

// void __stdcall DocumentOpen(IDispatch *pDispDoc)
__declspec(selectany) _ATL_FUNC_INFO OnDocumentOpenInfo = {CC_STDCALL, VT_EMPTY, 1, 
{VT_DISPATCH}};

// void __stdcall NewDocument(IDispatch *pDispDoc )
__declspec(selectany) _ATL_FUNC_INFO OnNewDocumentInfo = {CC_STDCALL, VT_EMPTY, 1, 
{VT_DISPATCH}};

// һ��ר�Ž��ճ����¼�����
class KAppEvent: public IDispEventSimpleImpl<1, KAppEvent, &__uuidof(WPS::ApplicationEvents)>
{
private:
	WPS::_ApplicationPtr m_spApp;
public:
	BEGIN_SINK_MAP(KAppEvent)
		SINK_ENTRY_INFO(1,  __uuidof(WPS::ApplicationEvents), 4, DocumentOpen, &OnDocumentOpenInfo)
		SINK_ENTRY_INFO(1,  __uuidof(WPS::ApplicationEvents), 9, NewDocument, &OnNewDocumentInfo)
	END_SINK_MAP()

	void init(WPS::_Application* pApp)
	{
		m_spApp = pApp;
	}

	void Term()
	{
		m_spApp = NULL;
	}

	// �����¼�
	void EventAdvise(WPS::_Application* pApp)
	{
		DispEventAdvise(pApp);
	}

	// ֹͣ�����¼�
	void EventUnadvise(WPS::_Application* pApp)
	{
		DispEventUnadvise(pApp);
	}
	void __stdcall DocumentOpen (IDispatch *pDispDoc)
	{		
		MessageBox(GetActiveWindow(), L"�����һ���ĵ�", L"��ʾ", MB_TOPMOST);
	}
	void __stdcall NewDocument (IDispatch *pDispDoc )
	{
		MessageBox(GetActiveWindow(), L"���½�һ���ĵ�", L"��ʾ", MB_TOPMOST);
	}
};


#endif // app_event_h__