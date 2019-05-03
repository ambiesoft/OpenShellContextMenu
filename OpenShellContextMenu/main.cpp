#include "stdafx.h"



#include "../../lsMisc/CreateSimpleWindow.h"
#include "../../lsMisc/GetFilesInfo.h"
#include "../../lsMisc/OpenCommon.h"
#include "../../lsMisc/IsWindowsNT.h"
// #include "../../lsMisc/stlScopedClear.h"
#include "../../lsMisc/stdwin32/stdwin32.h"
#include "../../lsMisc/CommandLineParser.h"
#include "../../lsMisc/stdosd/stdosd.h"
#include "../../lsMisc/tstring.h"
#include "../../lsMisc/CHandle.h"
#include "../../lsMisc/WaitWindowClose.h"

#include "ContextMenuCB.h"

#pragma comment(lib,"Credui.lib")
#pragma comment(lib,"shlwapi.lib")
#pragma comment(lib,"shell32.lib")

using namespace Ambiesoft;
using namespace Ambiesoft::stdosd;
using namespace stdwin32;
using namespace std;

#define RETURNFALSE do {DASSERT(FALSE);return FALSE;} while(0)

HWND ghMain;

typedef vector<wstring> STRVEC;

typedef HRESULT(__stdcall* pfnSHCreateDefaultContextMenu)(const DEFCONTEXTMENU *pdcm, REFIID riid, void **ppv);
pfnSHCreateDefaultContextMenu fnSHCreateDefaultContextMenu;

CContextMenuCB m_cccb;
IContextMenuPtr m_pcm;
IContextMenu2Ptr m_pcm2;
IContextMenu3Ptr m_pcm3;
int TfxMessageBox(LPCWSTR pMessage)
{
	return MessageBox(ghMain, pMessage, APPNAME, MB_ICONEXCLAMATION);
}
void showpidl(LPITEMIDLIST pidl)
{
	TRACE1("sizeof ITEMIDLIST is %d\r\n", sizeof(ITEMIDLIST));
	TRACE1("sizeof SHITEMID is %d\r\n", sizeof(SHITEMID));

	wstring str;
	if (!pidl)
		str = L"<NULL>";
	else
	{

		for (LPITEMIDLIST p = pidl; p->mkid.cb;)
		{
			TCHAR szT[512];
			wsprintf(szT, L"size is %d : ", p->mkid.cb);
			TRACE(szT);
			for (USHORT i = 0; i < p->mkid.cb; ++i)
			{
				// wsprintf(szT, L"%x", p->mkid.abID[i]);
				wsprintf(szT, L"%x", *((BYTE*)p + i));
				TRACE(szT);
			}

			TRACE(L"\r\n");
			p = (LPITEMIDLIST)(((BYTE*)p) + p->mkid.cb);
		}
	}
}

HRESULT getPIDLsFromPath(const STRVEC& arrayFiles, LPCITEMIDLIST** ppItemIDList)
{
	IShellFolderPtr pSFDesktop;
	SHGetDesktopFolder(&pSFDesktop);

	int nCount = (int)arrayFiles.size();
	LPCITEMIDLIST* pRet = (LPCITEMIDLIST*)calloc(sizeof(LPITEMIDLIST), nCount + 1);
	for (int i = 0; i < nCount; ++i)
	{
		LPITEMIDLIST pidl = NULL;
		pSFDesktop->ParseDisplayName(NULL, 0, (LPWSTR)(LPCWSTR)arrayFiles[i].c_str(), NULL, &pidl, NULL);
		pRet[i] = pidl;
#ifdef _DEBUG
		showpidl(pidl);
#endif
		// get relative pidl via SHBindToParent
		//SHBindToParentEx (pidl, IID_IShellFolder, (void **) &psfFolder, (LPCITEMIDLIST *) &pidlItem);
		//m_pidlArray[i] = CopyPIDL (pidlItem);	// copy relative pidl to pidlArray
		//free (pidlItem);
		//lpMalloc->Free (pidl);		// free pidl allocated by ParseDisplayName
		//psfFolder->Release ();
	}
	*ppItemIDList = pRet;
	return S_OK;
}

int countPidls(LPCITEMIDLIST* pItemIDList)
{
	int ret = 0;
	while (*pItemIDList)
	{
		ret++;
		pItemIDList++;
	}

	return ret;
}

void getExtsFromPath(const STRVEC& arrayFiles, set<wstring>& exts)
{
	int nCount = (int)arrayFiles.size();
	for (int i = 0; i < nCount; ++i)
	{
		size_t nDot = arrayFiles[i].rfind(L'.');// ReverseFind(L'.');
		if (nDot==wstring::npos)
			continue;
		size_t nBS = arrayFiles[i].rfind(L'\\');// ReverseFind(L'\\');
		if (nBS > nDot)
			continue;

		wstring ext(arrayFiles[i].substr(0, arrayFiles[i].size() - nDot));// Right(arrayFiles[i].GetLength() - nDot));
		std::transform(ext.begin(), ext.end(), ext.begin(), ::tolower);
		exts.insert(ext);
	}
}

HRESULT getKeysFromPath(const STRVEC& arrayFiles, HKEY** ppKey, UINT* pnKeyCount)
{
	HKEY buf[16] = { 0 };
	int bufi = 0;

	if (ERROR_SUCCESS == RegOpenKeyEx(
		HKEY_CLASSES_ROOT,
		L"*",
		0,					// ulOptions 
		KEY_READ,
		&buf[bufi]))
	{
		++bufi;
	}

	if (!false)
	{
		set<wstring> exts;
		getExtsFromPath(arrayFiles, exts);

		set<wstring>::iterator it = exts.begin();
		for (; it != exts.end(); ++it)
		{
			if (ERROR_SUCCESS == RegOpenKeyEx(
				HKEY_CLASSES_ROOT,
				it->c_str(),
				0,					// ulOptions 
				KEY_READ,
				&buf[bufi]))
			{
				++bufi;
				if (bufi >= 16)
					break;

				HKEY hKey = buf[bufi - 1];
				DWORD dwType = 0;
				BYTE byteBuff[512];
				DWORD dwBuffSize = sizeof(byteBuff);
				if (ERROR_SUCCESS == RegQueryValueEx(
					hKey,
					NULL,	// default key
					NULL,	// reserved
					&dwType,// type
					byteBuff,// ret
					&dwBuffSize))
				{
					if (dwType == REG_SZ)
					{
						LPCWSTR pFT = (LPCWSTR)byteBuff;

						if (ERROR_SUCCESS == RegOpenKeyEx(
							HKEY_CLASSES_ROOT,
							pFT,
							0,					// ulOptions 
							KEY_READ,
							&buf[bufi]))
						{
							++bufi;
							if (bufi >= 16)
								break;
						}
					}
				}
			}
		}
	}

	*ppKey = (HKEY*)calloc(sizeof(HKEY), 16);
	memcpy(*ppKey, buf, sizeof(buf));
	*pnKeyCount = bufi;
	return S_OK;
}

#define DFM_INVOKECOMMANDEX 12
HRESULT CALLBACK shellcb(
	IShellFolder *psf,
	HWND         hwnd,
	IDataObject  *pdtobj,
	UINT         uMsg,
	WPARAM       wParam,
	LPARAM       lParam)
{
	switch (uMsg)
	{
	case DFM_MERGECONTEXTMENU:
		// Must return S_OK even if we do nothing else or Vista and later
		// won't add standard verbs
		TRACE("DFM_MERGECONTEXTMENU\r\n");
		return S_OK;

	case DFM_INVOKECOMMAND: // Required to invoke default action
		TRACE("DFM_INVOKECOMMAND\r\n");
		return S_FALSE;

		//	case DFM_INVOKECOMMANDEX:
		//		TRACE0("DFM_INVOKECOMMANDEX\r\n");
		//		return S_FALSE;
		//	
		//	case DFM_GETDEFSTATICID: // Required for Windows 7 to pick a default
		//		TRACE0("DFM_GETDEFSTATICID\r\n");
		//		return S_FALSE; 

	default:
		TRACE1("DEFAULT %d\r\n", uMsg);
		return E_NOTIMPL; // Required for Windows 7 to show any menu at all
	}

	return E_NOTIMPL;
}


HRESULT STDMETHODCALLTYPE CContextMenuCB::CallBack(
	/* [unique][in] */
	IShellFolder *psf,
	/* [in] */
	HWND hwndOwner,
	/* [unique][in] */
	IDataObject *pdtobj,
	/* [in] */
	UINT uMsg,
	/* [in] */
	WPARAM wParam,
	/* [in] */
	LPARAM lParam)
{
	return shellcb(psf, hwndOwner, pdtobj, uMsg, wParam, lParam);
}

HRESULT mySHParseDisplayName(
	LPCWSTR          pszName,
	IBindCtx         *pbc,
	LPITEMIDLIST	*ppidl,
	SFGAOF           sfgaoIn,
	SFGAOF           *psfgaoOut)
{
	IShellFolderPtr pDesk;
	if (FAILED(SHGetDesktopFolder(&pDesk)))
		return E_FAIL;

	ULONG eaten = 0;
	if (FAILED(pDesk->ParseDisplayName(
		NULL,
		NULL,
		(LPWSTR)pszName,
		&eaten,
		ppidl,
		psfgaoOut)))
	{
		return E_FAIL;
	}

	return S_OK;
}

HRESULT GetUIObjectOfFile(HWND hwnd, LPCWSTR pszPath, REFIID riid, void **ppv)
{
	*ppv = NULL;
	HRESULT hr;
	LPITEMIDLIST pidl;
	SFGAOF sfgao;
	if (SUCCEEDED(hr = mySHParseDisplayName(
		pszPath,
		NULL,
		&pidl,
		0,
		&sfgao)))
	{
		IShellFolderPtr pSF;
		LPCITEMIDLIST pidlChild;
		if (SUCCEEDED(hr = SHBindToParent(pidl,
			IID_IShellFolder,
			(void**)&pSF,
			&pidlChild)))
		{
			hr = pSF->GetUIObjectOf(hwnd, 1, &pidlChild, riid, NULL, ppv);
		}
		CoTaskMemFree(pidl);
	}
	return hr;
}

wstring removeExt(LPCWSTR p)
{
	LPWSTR pT = stdStrDup(p);
	LPWSTR pRet = wcsrchr(pT, L'.');
	if (!pRet)
		return p;
	*pRet = 0;
	return pT;
}
BOOL createFileMenu(HMENU hSendTo, LPCTSTR pDirectory, map<size_t, wstring>& sendtomap)
{
	FILESINFOW fi;
	if (!GetFilesInfoW(pDirectory, fi))
		RETURNFALSE;

	fi.Sort();

	size_t count = fi.GetCount();

	for (size_t i = 0; i < count; ++i)
	{
		if ((fi[i].dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
			continue;

		if (!InsertMenu(hSendTo, (int)i, MF_BYPOSITION, ID_SHELLCUSTOM_SENDTO_START + i, removeExt(fi[i].cFileName).c_str()))
			RETURNFALSE;

		sendtomap[ID_SHELLCUSTOM_SENDTO_START + i] = fi[i].cFileName;
	}

	return TRUE;
}

wstring dqIfSpace(const wstring& s)
{
	if (s.empty())
		return s;

	if (s[0] == _T('"'))
		return s;

	if (s.find(_T(" ")) == wstring::npos)
		return s;

	return _T("\"") + s + _T("\"");
}

BOOL CreateShellMenu(HMENU hmenu, IContextMenuPtr pcm, map<size_t, wstring>& sendtomap, BOOL bUseMulti)
{
	IContextMenu3Ptr pcm3 = pcm;
	if (pcm3)
	{
		if (FAILED(pcm3->QueryContextMenu(hmenu, 1,
			ID_SHELLMENU_START, ID_SHELLMENU_END,
			CMF_NORMAL)))
		{
			return FALSE;
		}
	}
	else
	{
		IContextMenu2Ptr pcm2 = pcm;
		if (pcm2)
		{
			if (FAILED(pcm2->QueryContextMenu(hmenu, 1,
				ID_SHELLMENU_START, ID_SHELLMENU_END,
				CMF_NORMAL)))
			{
				return FALSE;
			}
		}
		else
		{
			if (FAILED(pcm->QueryContextMenu(hmenu, 1,
				ID_SHELLMENU_START, ID_SHELLMENU_END,
				CMF_NORMAL)))
			{
				return FALSE;
			}
		}
	}

	UINT count = GetMenuItemCount(hmenu);
	int iFirstSep = -1;
	for (UINT i = 0; i < count; ++i)
	{
		MENUITEMINFO mii = { 0 };
		mii.cbSize = sizeof(mii);
		mii.fMask = MIIM_TYPE;
		if (!GetMenuItemInfo(hmenu, i, TRUE, &mii))
			return FALSE;

		if (mii.fType & MF_SEPARATOR)
		{
			iFirstSep = i;
			break;
		}
	}

	if (iFirstSep < 0)
	{
		DASSERT(FALSE);
		return FALSE;
	}



	if (bUseMulti)
	{
		HMENU hSendTo = CreatePopupMenu();
		{

			DASSERT(hSendTo);

			TCHAR szSendToPath[MAX_PATH];
			if (!SHGetSpecialFolderPath(NULL, szSendToPath, CSIDL_SENDTO, FALSE))
				RETURNFALSE;

			PathAddBackslash(szSendToPath);
			if (!createFileMenu(hSendTo, szSendToPath, sendtomap))
				RETURNFALSE;

		}

		iFirstSep++;
		if (!InsertMenu(hmenu, iFirstSep++, MF_BYPOSITION | MF_POPUP, (UINT_PTR)hSendTo, I18N(L"Send To")))
			return FALSE;
		if (!InsertMenu(hmenu, iFirstSep++, MF_BYPOSITION | MF_SEPARATOR, 0, NULL))
			return FALSE;
		if (!InsertMenu(hmenu, iFirstSep++, MF_BYPOSITION, ID_SHELLCUSTOM_CUT, I18N(L"Cut")))
			return FALSE;
		if (!InsertMenu(hmenu, iFirstSep++, MF_BYPOSITION, ID_SHELLCUSTOM_COPY, I18N(L"Copy")))
			return FALSE;
		if (!InsertMenu(hmenu, iFirstSep++, MF_BYPOSITION | MF_SEPARATOR, 0, NULL))
			return FALSE;
	}

	if (!InsertMenu(hmenu, 0, MF_BYPOSITION, ID_SHELLCUSTOM_OPENPARENT, I18N(L"Open &Parent Folder")))
	{
		return FALSE;
	}

	if (!InsertMenu(hmenu, 1, MF_BYPOSITION | MF_SEPARATOR, 0, NULL))
		return FALSE;

	return TRUE;
}

void freepKey(HKEY* pKey)
{
	if (!pKey)
		return;

	for (int i = 0; i < 16; ++i)
	{
		if (!pKey[i])
			break;

		RegCloseKey(pKey[i]);
	}

	free(pKey);
}
_COM_SMARTPTR_TYPEDEF(IShellLinkDataList, IID_IShellLinkDataList);
static BOOL GetShortcutFileInfo(LPCTSTR pszShortcutFile,
	tstring& targetFile,
	tstring& curDir,
	tstring& arg,
	BOOL* pbIsAdmin = NULL)
{
	BOOL bFailed = TRUE;
	HRESULT hr;
	IShellLinkWPtr pShellLink = NULL;
	CoInitialize(NULL);
	TCHAR buffer[MAX_PATH];
	hr = CoCreateInstance(CLSID_ShellLink, NULL, CLSCTX_INPROC_SERVER, IID_IShellLink, (LPVOID*)&pShellLink);
	if (SUCCEEDED(hr) && pShellLink != NULL)
	{
		IPersistFilePtr pPFile = NULL;
		hr = pShellLink->QueryInterface(IID_IPersistFile, (LPVOID*)&pPFile);
		if (SUCCEEDED(hr) && pPFile != NULL)
		{
			hr = pPFile->Load(pszShortcutFile, 0);
			if (SUCCEEDED(hr))
			{
				bFailed = FALSE;
				if (FAILED(pShellLink->GetPath(buffer, _countof(buffer), NULL, 0)))
					return FALSE;
				targetFile = buffer;

				int size = 16;
				do
				{
					//TCHAR* pBuff = (TCHAR*)malloc(size);
					//stlsoft::scoped_handle<void*> ma(pBuff, free);
					unique_ptr<TCHAR[]> buff(new TCHAR[size]);
					if (FAILED(pShellLink->GetArguments(buff.get(), size / sizeof(TCHAR))))
						return FALSE;

					size *= 2;
					//TCHAR* pBuff2 = (TCHAR*)malloc(size);
					//stlsoft::scoped_handle<void*> ma2(pBuff2, free);
					unique_ptr<TCHAR[]> buff2(new TCHAR[size]);
					if (FAILED(pShellLink->GetArguments(buff2.get(), size / sizeof(TCHAR))))
						return FALSE;

					if (lstrcmp(buff.get(), buff2.get()) == 0)
					{
						arg = buff2.get();
						break;
					}
				} while (true);


				if (FAILED(pShellLink->GetWorkingDirectory(buffer, _countof(buffer))))
					return FALSE;

				curDir = buffer;
			}
		}

		if (pbIsAdmin)
		{
			*pbIsAdmin = FALSE;
			IShellLinkDataListPtr pSLDL;
			hr = pShellLink->QueryInterface(IID_IShellLinkDataList, (void**)&pSLDL);
			if (pSLDL)
			{
				DWORD dwFlags = 0;
				hr = pSLDL->GetFlags(&dwFlags);
				if (SUCCEEDED(hr))
				{
					if (SLDF_RUNAS_USER & dwFlags)
					{
						*pbIsAdmin = TRUE;
					}
				}
			}
		}
	}

	CoUninitialize();
	return !bFailed;
}

BOOL OpenCommonShortcutSpecial(HWND hWnd, LPCTSTR pApp, LPCTSTR pCommand = NULL, LPCTSTR pDirectory = NULL)
{
	tstring targetFile;
	tstring curDir;
	tstring arg;
	BOOL bIsAdmin = FALSE;
	if (!GetShortcutFileInfo(pApp,
		targetFile,
		curDir,
		arg,
		&bIsAdmin))
	{
		return FALSE;
	}

	tstring command(arg);
	if (pCommand && pCommand[0])
	{
		command += _T(" ");
		command += pCommand;
	}
	//LPTSTR pCT = _wcsdup(command.c_str());
	// stlsoft::scoped_handle<void*> ma(pCT, free);
	unique_ptr<TCHAR[]> pCT(stdStrDup(command.c_str()));

	if (FALSE) // bIsAdmin)
	{
		HANDLE hToken = NULL;

		TOKEN_PRIVILEGES tp;
		PROCESS_INFORMATION pi;
		STARTUPINFOW si;

		// Initialize structures.
		ZeroMemory(&tp, sizeof(tp));
		ZeroMemory(&pi, sizeof(pi));
		ZeroMemory(&si, sizeof(si));
		si.cb = sizeof(si);


		LPTSTR lpszUsername = _T("Administrator\0");
		LPTSTR lpszDomain = _T(".");		//"bgt\0";
		LPTSTR lpszPassword = _T("\0");

		if (!OpenProcessToken(
			GetCurrentProcess(),
			TOKEN_QUERY |
			TOKEN_ADJUST_PRIVILEGES,
			&hToken))
		{
			return FALSE;
		}

		// Look up the LUID for the TCB Name privilege.
		if (!LookupPrivilegeValue(
			NULL,
			SE_TCB_NAME, //SE_SHUTDOWN_NAME ,//SE_TCB_NAME,
			&tp.Privileges[0].Luid))
		{
			return FALSE;
		}


		tp.PrivilegeCount = 1;
		tp.Privileges[0].Attributes = SE_PRIVILEGE_ENABLED;//SE_PRIVILEGE_ENABLED;
		if (!AdjustTokenPrivileges(
			hToken,
			FALSE,
			&tp,
			0,
			NULL,
			0))
		{
			return FALSE;
		}


		TCHAR  szUserName[256] = { 0 };
		TCHAR  szPassword[256] = { 0 };
		DWORD  dwResult;
		BOOL   bSave = FALSE;
		// HANDLE hToken;

		dwResult = CredUIPromptForCredentials(
			NULL,					// Dialog customize
			TEXT("target-host"),	// displayed on the top
			NULL,					// reserved
			0, // ERROR_ELEVATION_REQUIRED,	// reason for calling this function
			szUserName, sizeof(szUserName) / sizeof(TCHAR), // in-out username
			szPassword, sizeof(szPassword) / sizeof(TCHAR), // in-out password
			&bSave,					// in-out save checkbutton
			CREDUI_FLAGS_EXPECT_CONFIRMATION |
			// CREDUI_FLAGS_GENERIC_CREDENTIALS |
			CREDUI_FLAGS_REQUEST_ADMINISTRATOR |
			CREDUI_FLAGS_USERNAME_TARGET_CREDENTIALS |
			CREDUI_FLAGS_EXCLUDE_CERTIFICATES |
			//			CREDUI_FLAGS_COMPLETE_USERNAME |
			0
			);
		if (dwResult != NO_ERROR)
			return FALSE;

		if (LogonUser(szUserName, NULL, szPassword, LOGON32_LOGON_INTERACTIVE, LOGON32_PROVIDER_DEFAULT, &hToken))
		{
			if (bSave)
				CredUIConfirmCredentials(TEXT("target-host"), TRUE);
			CloseHandle(hToken);
		}
		else
		{
			MessageBox(NULL, TEXT("Logon failed"), NULL, MB_ICONWARNING);
			return FALSE;
		}

		SecureZeroMemory(szPassword, sizeof(szPassword));



		//if(LogonUser(lpszUsername,lpszDomain,lpszPassword,
		//	LOGON32_LOGON_INTERACTIVE,LOGON32_PROVIDER_DEFAULT,&hToken) == 0)
		//{
		//	MyError();
		//}
		//else

		STARTUPINFO sInfo;
		PROCESS_INFORMATION ProcessInfo;
		memset(&sInfo, 0, sizeof(STARTUPINFO));
		sInfo.cb = sizeof(STARTUPINFO);
		sInfo.dwX = CW_USEDEFAULT;
		sInfo.dwY = CW_USEDEFAULT;
		sInfo.dwXSize = CW_USEDEFAULT;
		sInfo.dwYSize = CW_USEDEFAULT;


		if (!CreateProcessAsUser(
			hToken,
			targetFile.c_str(), // _T("c:\\windows\\system32\\notepad.exe"),
			pCT.get(),
			NULL,
			NULL,
			TRUE,
			CREATE_NEW_CONSOLE,
			NULL,
			NULL,
			&sInfo,
			&ProcessInfo))
		{
			return FALSE;
		}
		CloseHandle(pi.hThread);
		CloseHandle(pi.hProcess);
	}
	else
	{


		STARTUPINFO si = { sizeof(STARTUPINFO) };
		PROCESS_INFORMATION pi = { 0 };
		if (!CreateProcess(
			targetFile.c_str(),
			pCT.get(),
			NULL,
			NULL,
			FALSE,
			0,
			NULL,
			pDirectory,
			&si,
			&pi)
			)
		{
			return FALSE;
		}
		CloseHandle(pi.hThread);
		CloseHandle(pi.hProcess);
	}

	return TRUE;
}

BOOL SetFileOntoClipboard(const STRVEC& arFiles, BOOL bCut)
{
	size_t i;
	// COleDataSource ds;

	if (!OpenClipboard(NULL))
		return FALSE;
	if (!EmptyClipboard())
		return FALSE;


	{
		size_t uBuffSize = 0;
		HGLOBAL hgDrop = NULL;


		for (i = 0; i < arFiles.size(); ++i)
		{
			uBuffSize += arFiles[i].size() + 1;
		}

		uBuffSize = sizeof(DROPFILES) + sizeof(TCHAR) * (uBuffSize + 1);

		hgDrop = GlobalAlloc(GHND | GMEM_SHARE, uBuffSize);
		if (NULL == hgDrop)
			return FALSE;

		DROPFILES* pDrop = (DROPFILES*)GlobalLock(hgDrop);
		if (NULL == pDrop)
		{
			GlobalFree(hgDrop);
			return TRUE;
		}

		pDrop->pFiles = sizeof(DROPFILES);

#ifdef _UNICODE
		pDrop->fWide = TRUE;
#endif

		LPTSTR pszBuff = (TCHAR*)(LPBYTE(pDrop) + sizeof(DROPFILES));
		for (i = 0; i < arFiles.size(); ++i)
		{
			lstrcpy(pszBuff, (LPCTSTR)arFiles[i].c_str());
			pszBuff = 1 + wcschr(pszBuff, '\0');
		}

		//	while ( NULL != pos )
		//	{
		//		lstrcpy ( pszBuff, (LPCTSTR) lsDraggedFiles.GetNext ( pos ) );
		//		pszBuff = 1 + _tcschr ( pszBuff, '\0' );
		//	}

		GlobalUnlock(hgDrop);

		if (!SetClipboardData(CF_HDROP, hgDrop))
			return FALSE;

		// Put the data in the data source.
		// FORMATETC etcHDROP = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		// ds.CacheGlobalData ( CF_HDROP, hgDrop, &etcHDROP );

	}


	{
		static UINT CF_PREFERREDDROPEFFECT = RegisterClipboardFormat(CFSTR_PREFERREDDROPEFFECT);
		HGLOBAL hgDE = GlobalAlloc(GHND | GMEM_SHARE, sizeof(DWORD));
		if (NULL == hgDE)
			return FALSE;

		DWORD* pDE = (DWORD*)GlobalLock(hgDE);
		if (NULL == pDE)
		{
			GlobalFree(hgDE);
			return FALSE;
		}

		*pDE = bCut ? DROPEFFECT_MOVE : DROPEFFECT_COPY;

		GlobalUnlock(hgDE);

		if (!SetClipboardData(CF_PREFERREDDROPEFFECT, hgDE))
			return FALSE;

		//FORMATETC etcHDROP = { CF_PREFERREDDROPEFFECT, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		//ds.CacheGlobalData ( CF_PREFERREDDROPEFFECT, hg4, &etcHDROP );
	}

	CloseClipboard();

	return TRUE;
}

void freeppItemIDList(LPCITEMIDLIST* pItemIDList)
{
	if (!pItemIDList)
		return;

	IMallocPtr pMalloc;
	SHGetMalloc(&pMalloc);

	LPCITEMIDLIST* p = pItemIDList;
	while (*p)
	{
		pMalloc->Free((void*)*p);
		p++;
	}
	free(pItemIDList);
}

wstring ddd(const STRVEC& arFiles)
{
	wstring arg;
	for (size_t i = 0; i < arFiles.size(); ++i)
	{
		wstring s = arFiles[i];
		if (s[0] != L'"' && s.find(L' ') != wstring::npos)
			s = L"\"" + s + L"\"";

		arg += s;
		if ((i + 1) != arFiles.size())
			arg += L" ";
	}
	return arg;
}
void ShowShellContextMenu(const STRVEC& arFiles)
{
	BOOL bUseMulti = arFiles.size() > 1;

	LPCITEMIDLIST* pItemIDList = NULL;
	HKEY* pKey = NULL;
	IShellFolderPtr pSFDesktop;
	SHGetDesktopFolder(&pSFDesktop);
	// CShellFolderDS spy;
	HRESULT hr = E_FAIL;
	IContextMenuPtr pcm;

	do
	{
		if (bUseMulti)
		{
			if (FAILED(getPIDLsFromPath(arFiles, &pItemIDList)))
			{
				TfxMessageBox(I18N(L"failed to get pidls"));
				break;;
			}
			DASSERT(arFiles.size() == countPidls(pItemIDList));

			UINT nKeyCount = 0;
			if (FAILED(getKeysFromPath(arFiles, &pKey, &nKeyCount)))
			{
				TfxMessageBox(I18N(L"failed to get pidls keys"));
				break;
			}

			// spy.construct(pSFDesktop, arFiles);
			IShellFolder* pShellFolderDS = NULL;
			// spy.InternalQueryInterface(&IID_IShellFolder, (void**)&pShellFolderDS);
			
			if (fnSHCreateDefaultContextMenu)
			{
				DEFCONTEXTMENU dcm = { 0 };
				dcm.hwnd = ghMain;
				dcm.pcmcb = &m_cccb;
				dcm.pidlFolder = NULL;
				dcm.psf = pShellFolderDS;
				dcm.cidl = countPidls(pItemIDList);
				dcm.apidl = pItemIDList;

				dcm.cKeys = nKeyCount;
				dcm.aKeys = pKey;

				hr = fnSHCreateDefaultContextMenu(&dcm, IID_IContextMenu, (void**)&pcm);
			}
			else
			{
				//				LPITEMIDLIST p;
				//				SHGetSpecialFolderLocation(*this, 
				//					CSIDL_DESKTOP,
				//					&p);


				hr = CDefFolderMenu_Create2(
					NULL,				// An ITEMIDLIST structure for the parent folder. This value can be NULL.
					ghMain,				//
					countPidls(pItemIDList), // The number of ITEMIDLIST structures in the array pointed to by apidl.
					pItemIDList,		// A pointer to an array of ITEMIDLIST structures, one for each item that is selected.
					pShellFolderDS,	// A pointer to the parent folder's IShellFolder interface. 
					// This IShellFolder must support the IDataObject interface. 
					// If it does not, CDefFolderMenu_Create2 fails and returns E_NOINTERFACE. 
					// This value can be NULL.
					shellcb,			// The LPFNDFMCALLBACK callback object. This value can be NULL if the callback object is not needed.
					0,//nKeyCount,			// The number of registry keys in the array pointed to by ahkeys.
					NULL, //pKey,				// A pointer to an array of registry keys that specify the context menu handlers
					// used with the menu's entries. For more information on context menu handlers, 
					// see Creating Context Menu Handlers. This array can contain a maximum of 16 registry keys.
					&pcm				// The address of an IContextMenu interface pointer that, 
					// when this function returns successfully, points to the IContextMenu object that represents the context menu.
					);
			}






			// spy.InternalRelease();

		}
		else
		{
			hr = GetUIObjectOfFile(ghMain,
				arFiles[0].c_str(),
				IID_IContextMenu,
				(void**)&pcm);
		}


		if (SUCCEEDED(hr))
		{
			//HMENU hmenu = CreatePopupMenu();
			//STLSCOPEDFREE(hmenu, HMENU, DestroyMenu);
			CHMenu menu(CreatePopupMenu());
			map<size_t, wstring> sendtomap;
			if (menu && CreateShellMenu(menu, pcm, sendtomap, bUseMulti))
			{
				pcm->QueryInterface(IID_IContextMenu2, (void**)&m_pcm2);
				pcm->QueryInterface(IID_IContextMenu3, (void**)&m_pcm3);
				m_pcm = pcm;

				POINT point;
				GetCursorPos(&point);
				int iCmd = TrackPopupMenuEx(menu, TPM_RETURNCMD,
					point.x, point.y, ghMain, NULL);



				// ((CMainFrame*)theApp.m_pMainWnd)->SetMessageText(L"");
				if (iCmd == ID_SHELLCUSTOM_OPENPARENT)
				{
					OpenFolder(ghMain, arFiles[0].c_str());
				}
				else if (ID_SHELLCUSTOM_SENDTO_START <= iCmd && iCmd <= ID_SHELLCUSTOM_SENDTO_END)
				{
					wstring runfile = sendtomap[iCmd];
					DASSERT(runfile.size() != 0);

					TCHAR szSendToPath[MAX_PATH];
					if (SHGetSpecialFolderPath(NULL, szSendToPath, CSIDL_SENDTO, FALSE))
					{
						PathAddBackslash(szSendToPath);
						lstrcat(szSendToPath, runfile.c_str());

						wstring arg = ddd(arFiles);
						wstring param = dqIfSpace(szSendToPath) + L" " + arg;
						// int len = arg.size();

						if (IsWinVistaOrHigher())
							OpenCommon(ghMain, szSendToPath, arg.c_str(), NULL);
						else
							OpenCommonShortcutSpecial(ghMain, szSendToPath, arg.c_str(), NULL);

					}
				}
				else if (iCmd == ID_SHELLCUSTOM_CUT)
				{
					SetFileOntoClipboard(arFiles, TRUE);
				}
				else if (iCmd == ID_SHELLCUSTOM_COPY)
				{
					SetFileOntoClipboard(arFiles, FALSE);
				}
				else if(iCmd==32792)
				{
					SHELLEXECUTEINFO info = {};
					wstring arg = ddd(arFiles);
					info.cbSize = sizeof info;
					info.lpFile = arg.c_str();
					info.nShow = SW_SHOW;
					info.fMask = SEE_MASK_INVOKEIDLIST;
					info.lpVerb = L"properties";

					ShellExecuteEx(&info);
				}
				else if (iCmd > 0)
				{
					CMINVOKECOMMANDINFOEX info = { 0 };
					info.cbSize = sizeof(info);
					// info.fMask = CMIC_MASK_UNICODE | CMIC_MASK_PTINVOKE;
					info.fMask = CMIC_MASK_PTINVOKE;
					if (GetKeyState(VK_CONTROL) < 0)
					{
						info.fMask |= CMIC_MASK_CONTROL_DOWN;
					}
					if (GetKeyState(VK_SHIFT) < 0)
					{
						info.fMask |= CMIC_MASK_SHIFT_DOWN;
					}
					info.hwnd = ghMain;
					info.lpVerb = MAKEINTRESOURCEA(iCmd - ID_SHELLMENU_START);
					//info.lpVerbW = MAKEINTRESOURCEW(iCmd - ID_SHELLMENU_START);
					info.nShow = SW_SHOWNORMAL;
					info.ptInvoke = point;

					/**
					TCHAR szVerbW[256];
					if(SUCCEEDED(pcm->GetCommandString(iCmd - ID_SHELLMENU_START,
					GCS_VERBW,
					NULL,
					(LPSTR)szVerbW,
					sizeof(szVerbW)/sizeof(szVerbW[0]))))
					{
					info.lpVerbW = szVerbW;
					}
					else
					{
					char szVerb[256];
					if(SUCCEEDED(pcm->GetCommandString(iCmd - ID_SHELLMENU_START,
					GCS_VERBA,
					NULL,
					szVerb,
					sizeof(szVerb)/sizeof(szVerb[0]))))
					{
					info.lpVerb = szVerb;
					}
					}
					**/
					if (m_pcm3)
						m_pcm3->InvokeCommand((LPCMINVOKECOMMANDINFO)&info);
					else
						pcm->InvokeCommand((LPCMINVOKECOMMANDINFO)&info);
				}
				Sleep(10000);

				m_pcm = NULL;
				m_pcm2 = NULL;
				m_pcm3 = NULL;
			}
		}
	} while (0);

	freepKey(pKey);
	freeppItemIDList(pItemIDList);
}

STRVEC gInFiles;

LRESULT CALLBACK MainWndProc(
	_In_ HWND   hwnd,
	_In_ UINT   uMsg,
	_In_ WPARAM wParam,
	_In_ LPARAM lParam
	)
{
	if (m_pcm3)
	{
		LRESULT lres;
		if (SUCCEEDED(m_pcm3->HandleMenuMsg2(uMsg, wParam, lParam, &lres)))
		{
			return lres;
		}
	}
	else if (m_pcm2)
	{
		if (SUCCEEDED(m_pcm2->HandleMenuMsg(uMsg, wParam, lParam)))
		{
			return 0;
		}
	}

	switch (uMsg)
	{
		case WM_CREATE:
		{
			
		}
		break;

		case WM_APP_TEST:
		{
			//STRVEC arFiles;
			//arFiles.push_back(L"C:\\LegacyPrograms\\WhiteBrowser\\WhiteBrowser.exe");
			ShowShellContextMenu(gInFiles);
			PostQuitMessage(0);
		}
		break;
	}
	return DefWindowProc(hwnd, uMsg, wParam, lParam);
}

int CALLBACK wWinMain2(
	_In_ HINSTANCE hInstance,
	_In_ HINSTANCE hPrevInstance,
	_In_ LPWSTR     lpCmdLine,
	_In_ int       nCmdShow
	)
{
	{
		CCommandLineParser parser;
		COption mainOption;
		parser.AddOption(&mainOption);
		parser.Parse();
		wstring unk = parser.getUnknowOptionStrings();
		if (!unk.empty())
		{
			wstring message = I18N(L"Unknown option(s):\n");
			message += unk;
			TfxMessageBox(message.c_str());
			return 1;
		}
		for (size_t i = 0; i < mainOption.getValueCount(); ++i)
		{
			gInFiles.push_back(mainOption.getValue(i));
		}
	}
	if (gInFiles.empty())
	{
		TfxMessageBox(I18N(L"No input path"));
		return 0;
	}
#ifdef NDEBUG
	if (gInFiles.size() > 1)
	{
		TfxMessageBox(I18N(L"Currently only one file is acceptable."));
		return 0;
	}
#endif
	if (FAILED(OleInitialize(NULL)))
	{
		TfxMessageBox(L"OleInitialize failed.");
		return 1;
	}

	HMODULE hShell = LoadLibrary(L"Shell32.dll");
	if (hShell)
	{
		fnSHCreateDefaultContextMenu = (pfnSHCreateDefaultContextMenu)GetProcAddress(hShell, "SHCreateDefaultContextMenu");
		FreeLibrary(hShell);
	}

	ghMain = CreateSimpleWindow(
		nullptr,
		L"OpenShellContextMenu_WindowClass",
		nullptr,
		MainWndProc,
		0,
		WS_EX_TOOLWINDOW,
		nullptr);
	CHWnd hfreer(ghMain);

	ShowWindow(ghMain, SW_SHOW);
	PostMessage(ghMain, WM_APP_TEST, 0, 0);
	MSG msg;
	BOOL bRet;
	while ((bRet = GetMessage(&msg, NULL, 0, 0)) != 0)
	{
		if (bRet == -1)
		{
			// handle the error and possibly exit
			return GetLastError();
		}
		else
		{
			TranslateMessage(&msg);
			DispatchMessage(&msg);
		}
	}

	return 0;
}

int CALLBACK wWinMain(
	_In_ HINSTANCE hInstance,
	_In_ HINSTANCE hPrevInstance,
	_In_ LPWSTR     lpCmdLine,
	_In_ int       nCmdShow
)
{
	int ret = wWinMain2(hInstance, hPrevInstance, lpCmdLine, nCmdShow);
	WaitWindowClose();
	return ret;
}
