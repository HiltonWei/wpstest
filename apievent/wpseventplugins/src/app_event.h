#ifndef app_event_h__
#define app_event_h__

//����һ���Զ�����Ϣ����Ϊ��Ҫ�µ�Windows����Ҳʹ����WM_User,΢����Ҫ�����ֵ��100���� 
#define WM_MyMessageOpen (WM_USER+101)
#define WM_MyMessageNew (WM_USER+102)
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
		//FindWindow������������������һ����Ҫ�ҵĴ��ڵ��࣬�ڶ�����Ҫ�ҵĴ��ڵı��⡣��������ʱ��һ�����߶�֪����������Ҫ֪�����е�һ�� 
		MessageBox(GetActiveWindow(), L"�����һ���ĵ�1", L"��ʾ", MB_TOPMOST);
		HWND fHandle = FindWindow(L"VC_WPSSTARTKIT", NULL);  
		PostMessage(fHandle, WM_MyMessageOpen, 100, 200);
	}
	void __stdcall NewDocument (IDispatch *pDispDoc )
	{
		MessageBox(GetActiveWindow(), L"���½�һ���ĵ�", L"��ʾ", MB_TOPMOST);
		HWND fHandle = FindWindow(L"VC_WPSSTARTKIT", NULL);  
		PostMessage(fHandle, WM_MyMessageNew, 100, 200);
	}
};

/*
��SendMessage��PostMessage������
SendMessage����Ϣֱ�ӷ��͵����ڣ������ô˴��ڵ���Ӧ��Ϣ������������Ϣ������������SendMessage�ŷ��أ�SendMessage�����з���ֵ��
PostMessage����Ϣ���͵��봴�����ڵ��߳����������Ϣ���к��������أ�PostMessage����û�з���ֵ��

*/


#endif // app_event_h__