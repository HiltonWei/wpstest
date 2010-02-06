// stdafx.h : include file for standard system include files,
//      or project specific include files that are used frequently,
//      but are changed infrequently
#if !defined(AFX_STDAFX_H__FBC2E757_B8F9_4EFF_BABB_0C20812AC203__INCLUDED_)
#define AFX_STDAFX_H__FBC2E757_B8F9_4EFF_BABB_0C20812AC203__INCLUDED_

#if _MSC_VER > 1000
#pragma once
#endif // _MSC_VER > 1000

#define STRICT
#define _ATL_APARTMENT_THREADED

//ATL support
#include <atlbase.h>
#include <atlcom.h>
#include <comdef.h>

//ATL Windowing support
#include <atlwin.h>

//Advanced Control support
#include<commctrl.h>

//Utilities
#include <string>
#include <windows.h> 
#include <vector>

extern CComModule _Module; //Do not change the name of _Module
using namespace std;

//开发者WPS的安装路径需要修改
// Microsoft Visual C++ will insert additional declarations immediately before the previous line.
#import "D:\Program Files\Kingsoft\WPS Office Personal\office6\ksaddndr.dll" named_guids, raw_interfaces_only
#import "D:\Program Files\Kingsoft\WPS Office Personal\office6\kso10.dll" named_guids, rename_namespace("KSO")
#import "D:\Program Files\Kingsoft\WPS Office Personal\office6\wpscore.dll" rename_namespace("WPS")
#import "D:\Program Files\Kingsoft\WPS Office Personal\office6\etapp.dll" rename_namespace("ET")
#import "D:\Program Files\Kingsoft\WPS Office Personal\office6\wpcre.dll" rename_namespace("WPP")
using namespace KSO;
//using namespace WPS;
//using namespace WPP;
//using namespace ET;

#ifdef _ATL_STATIC_REGISTRY
#include <statreg.h>
#include <statreg.cpp>
#endif

#include <atlimpl.cpp>
#include "resource.h"



//#include "resource.h"

//#pragma comment(lib,"ws2_32.lib")  
//#pragma comment(lib,"Comctl32.lib")

#endif // !defined(AFX_STDAFX_H__FBC2E757_B8F9_4EFF_BABB_0C20812AC203__INCLUDED)