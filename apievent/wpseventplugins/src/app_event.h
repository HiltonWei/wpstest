#ifndef app_event_h__
#define app_event_h__

//定义一个自定义消息，因为需要新的Windows特性也使用了WM_User,微软建议要比这个值大100以上 
#define WM_MyMessageOpen (WM_USER+101)
#define WM_MyMessageNew (WM_USER+102)
// void __stdcall DocumentOpen(IDispatch *pDispDoc)
__declspec(selectany) _ATL_FUNC_INFO OnDocumentOpenInfo = {CC_STDCALL, VT_EMPTY, 1, 
{VT_DISPATCH}};

// void __stdcall NewDocument(IDispatch *pDispDoc )
__declspec(selectany) _ATL_FUNC_INFO OnNewDocumentInfo = {CC_STDCALL, VT_EMPTY, 1, 
{VT_DISPATCH}};

// 一个专门接收程序事件的类
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

	// 接收事件
	void EventAdvise(WPS::_Application* pApp)
	{
		DispEventAdvise(pApp);
	}

	// 停止接收事件
	void EventUnadvise(WPS::_Application* pApp)
	{
		DispEventUnadvise(pApp);
	}
	void __stdcall DocumentOpen (IDispatch *pDispDoc)
	{	
		//FindWindow函数有两个参数，第一个是要找的窗口的类，第二个是要找的窗口的标题。在搜索的时候不一定两者都知道，但至少要知道其中的一个 
		MessageBox(GetActiveWindow(), L"你打开了一个文档1", L"提示", MB_TOPMOST);
		HWND fHandle = FindWindow(L"VC_WPSSTARTKIT", NULL);  
		PostMessage(fHandle, WM_MyMessageOpen, 100, 200);
	}
	void __stdcall NewDocument (IDispatch *pDispDoc )
	{
		MessageBox(GetActiveWindow(), L"你新建一个文档", L"提示", MB_TOPMOST);
		HWND fHandle = FindWindow(L"VC_WPSSTARTKIT", NULL);  
		PostMessage(fHandle, WM_MyMessageNew, 100, 200);
	}
};

/*
附SendMessage与PostMessage的区别：
SendMessage把消息直接发送到窗口，并调用此窗口的相应消息处理函数，等消息处理函数结束后SendMessage才返回，SendMessage函数有返回值；
PostMessage将消息发送到与创建窗口的线程相关联的消息队列后立即返回；PostMessage函数没有返回值。

*/


#endif // app_event_h__