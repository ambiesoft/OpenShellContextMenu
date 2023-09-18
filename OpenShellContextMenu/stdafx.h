#pragma once

#include <afx.h>
#include <afxwin.h>
#include <afxole.h>

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
#include <memory>
#include <functional>

#include "../../lsMisc/CreateSimpleWindow.h"
#include "../../lsMisc/GetFilesInfo.h"
#include "../../lsMisc/OpenCommon.h"
#include "../../lsMisc/IsWindowsNT.h"
#include "../../lsMisc/CommandLineParser.h"
#include "../../lsMisc/stdosd/stdosd.h"
#include "../../lsMisc/tstring.h"
#include "../../lsMisc/CHandle.h"
#include "../../lsMisc/WaitWindowClose.h"
#include "../../lsMisc/I18N.h"

#define DASSERT(x) assert(x)

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

typedef std::vector<std::wstring> STRVEC;
