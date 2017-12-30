#pragma once

#include <Windows.h>

#include <Shlwapi.h>
#include <ShlObj.h>
#include <Shobjidl.h>
#include <comdef.h>
#include <comip.h>
#include <Shtypes.h>
_COM_SMARTPTR_TYPEDEF(IContextMenu, __uuidof(IContextMenu));
_COM_SMARTPTR_TYPEDEF(IContextMenu2, __uuidof(IContextMenu2));
_COM_SMARTPTR_TYPEDEF(IContextMenu3, __uuidof(IContextMenu3));
#include <wincred.h>
#include <tchar.h>

#include <cassert>
#include <string>
#include <vector>
#include <set>
#include <map>
#include <algorithm>

#define DASSERT(x) assert(x)
#define TRACE(x) (void)0
#define TRACE1(x,y) (void)0

#define I18N(s) s

#define APPNAME L"OpenShellContextMenu"

#define ID_SHELLMENU_START              32773
#define ID_SHELLMENU_END                33773
#define ID_SHELLCUSTOM_OPENPARENT       33774
#define ID_SHELLCUSTOM_SENDTO_START     33775
#define ID_SHELLCUSTOM_SENDTO_END       34775
#define ID_SHELLCUSTOM_CUT              34776
#define ID_SHELLCUSTOM_COPY             34777

// #define _T(s) L##s
enum {
	WM_APP_TEST = WM_APP+1,
};