#include "stdafx.h"
#include "ShellFolderDS.h"


BEGIN_INTERFACE_MAP(CShellFolderDS, CCmdTarget)
    INTERFACE_PART(CShellFolderDS, IID_IShellFolder, ShellFolderDS)
END_INTERFACE_MAP()



ULONG STDMETHODCALLTYPE CShellFolderDS::XShellFolderDS::AddRef(void)
{
	METHOD_PROLOGUE(CShellFolderDS, ShellFolderDS)
	return pThis->m_iSF->AddRef();
}
ULONG STDMETHODCALLTYPE CShellFolderDS::XShellFolderDS::Release(void)
{
	METHOD_PROLOGUE(CShellFolderDS, ShellFolderDS)
	return pThis->m_iSF->Release();
}
HRESULT STDMETHODCALLTYPE CShellFolderDS::XShellFolderDS::QueryInterface(REFIID iid, LPVOID* ppvObj)
{
	METHOD_PROLOGUE(CShellFolderDS, ShellFolderDS)
	return pThis->m_iSF->QueryInterface(iid, ppvObj);
}


// IShellFolder methods ----
HRESULT STDMETHODCALLTYPE CShellFolderDS::XShellFolderDS::GetUIObjectOf( // this method we "override"
      /* [in] */ HWND hwndOwner,
      /* [in] */ UINT cidl,
      /* [size_is][in] */ LPCITEMIDLIST *apidl,
      /* [in] */ REFIID riid,
      /* [unique][out][in] */ UINT *rgfReserved,
      /* [iid_is][out] */ void **ppv)
{

	METHOD_PROLOGUE(CShellFolderDS, ShellFolderDS)


	if(InlineIsEqualGUID(riid, IID_IDataObject))
	{
		// i ignore the pidl array supplied and return the good data object
		return pThis->m_ds.InternalQueryInterface(&riid,ppv);//  riid, ppv);
	}
	else 
	{
		// i've seen IDropTarget requests, let base handle them
		return pThis->m_iSF->GetUIObjectOf(hwndOwner, cidl, apidl, riid, rgfReserved, ppv);
	}
}





HRESULT STDMETHODCALLTYPE CShellFolderDS::XShellFolderDS::ParseDisplayName( 
    /* [in] */ HWND hwnd,
    /* [in] */ LPBC pbc,
    /* [string][in] */ LPOLESTR pszDisplayName,
    /* [out] */ ULONG *pchEaten,
    /* [out] */ LPITEMIDLIST *ppidl,
    /* [unique][out][in] */ ULONG *pdwAttributes)
{
	METHOD_PROLOGUE(CShellFolderDS, ShellFolderDS)
	return pThis->m_iSF->ParseDisplayName(hwnd,pbc,pszDisplayName,pchEaten,ppidl,pdwAttributes);
}

HRESULT STDMETHODCALLTYPE CShellFolderDS::XShellFolderDS::EnumObjects( 
    /* [in] */ HWND hwnd,
    /* [in] */ SHCONTF grfFlags,
    /* [out] */ IEnumIDList **ppenumIDList)
{
	METHOD_PROLOGUE(CShellFolderDS, ShellFolderDS)
	return pThis->m_iSF->EnumObjects(hwnd,grfFlags,ppenumIDList);
}

HRESULT STDMETHODCALLTYPE CShellFolderDS::XShellFolderDS::BindToObject( 
    /* [in] */ LPCITEMIDLIST pidl,
    /* [in] */ LPBC pbc,
    /* [in] */ REFIID riid,
    /* [iid_is][out] */ void **ppv)
{
	METHOD_PROLOGUE(CShellFolderDS, ShellFolderDS)
	return pThis->m_iSF->BindToObject(pidl,pbc,riid,ppv);
}

HRESULT STDMETHODCALLTYPE CShellFolderDS::XShellFolderDS::BindToStorage( 
    /* [in] */ LPCITEMIDLIST pidl,
    /* [in] */ LPBC pbc,
    /* [in] */ REFIID riid,
    /* [iid_is][out] */ void **ppv)
{
	METHOD_PROLOGUE(CShellFolderDS, ShellFolderDS)
	return pThis->m_iSF->BindToStorage(pidl,pbc,riid,ppv);
}

HRESULT STDMETHODCALLTYPE CShellFolderDS::XShellFolderDS::CompareIDs( 
    /* [in] */ LPARAM lParam,
    /* [in] */ LPCITEMIDLIST pidl1,
    /* [in] */ LPCITEMIDLIST pidl2) 
{
	METHOD_PROLOGUE(CShellFolderDS, ShellFolderDS)
	return pThis->m_iSF->CompareIDs(lParam,pidl1,pidl2);
}

HRESULT STDMETHODCALLTYPE CShellFolderDS::XShellFolderDS::CreateViewObject( 
    /* [in] */ HWND hwndOwner,
    /* [in] */ REFIID riid,
    /* [iid_is][out] */ void **ppv) 
{
	METHOD_PROLOGUE(CShellFolderDS, ShellFolderDS)
	return pThis->m_iSF->CreateViewObject(hwndOwner,riid,ppv);
}

HRESULT STDMETHODCALLTYPE CShellFolderDS::XShellFolderDS::GetAttributesOf( 
    /* [in] */ UINT cidl,
    /* [size_is][in] */ LPCITEMIDLIST *apidl,
    /* [out][in] */ SFGAOF *rgfInOut) 
{
	METHOD_PROLOGUE(CShellFolderDS, ShellFolderDS)
	return pThis->m_iSF->GetAttributesOf( cidl,apidl,rgfInOut);
}



HRESULT STDMETHODCALLTYPE CShellFolderDS::XShellFolderDS::GetDisplayNameOf( 
    /* [in] */ LPCITEMIDLIST pidl,
    /* [in] */ SHGDNF uFlags,
    /* [out] */ STRRET *pName) 
{
	METHOD_PROLOGUE(CShellFolderDS, ShellFolderDS)
	return pThis->m_iSF->GetDisplayNameOf( pidl,uFlags,pName);
}

HRESULT STDMETHODCALLTYPE CShellFolderDS::XShellFolderDS::SetNameOf( 
    /* [in] */ HWND hwnd,
    /* [in] */ LPCITEMIDLIST pidl,
    /* [string][in] */ LPCOLESTR pszName,
    /* [in] */ SHGDNF uFlags,
    /* [out] */ LPITEMIDLIST *ppidlOut) 
{
	METHOD_PROLOGUE(CShellFolderDS, ShellFolderDS)
	return pThis->m_iSF->SetNameOf(hwnd,pidl,pszName,uFlags,ppidlOut);
}


























void getFD(LPCTSTR pFile, FILEDESCRIPTOR* pFD)
{
	pFD->dwFlags = FD_ATTRIBUTES;

	pFD->dwFileAttributes = GetFileAttributes(pFile);
	HANDLE h = CreateFile(
		pFile, 
		GENERIC_READ,
		FILE_SHARE_READ,
		NULL,
		OPEN_EXISTING,
		FILE_ATTRIBUTE_NORMAL,
		NULL);
	if(h != INVALID_HANDLE_VALUE)
	{
		FILETIME ftCT,ftLA,ftLW;
		if(GetFileTime(h, &ftCT, &ftLA, &ftLW))
		{
			pFD->dwFlags |= FD_CREATETIME | FD_ACCESSTIME | FD_WRITESTIME ;
			pFD->ftCreationTime = ftCT;
			pFD->ftLastAccessTime = ftLA;
			pFD->ftLastWriteTime = ftLW;
		}

		DWORD dwSizeHigh=0;
		DWORD dwSizeLow=0;
		dwSizeLow = GetFileSize(h, &dwSizeHigh);
		if( !(dwSizeLow==INVALID_FILE_SIZE && GetLastError() != NO_ERROR) )
		{
			pFD->dwFlags |= FD_FILESIZE;
			pFD->nFileSizeLow = dwSizeLow;
			pFD->nFileSizeHigh = dwSizeHigh;
		}

		CloseHandle(h);
	}

	lstrcpyn(pFD->cFileName, pFile, sizeof(pFD->cFileName)/sizeof(pFD->cFileName[0]));
	
}


size_t getPidlSize(LPITEMIDLIST pidl)
{
	size_t ret = sizeof(pidl->mkid.cb); // last 0
	for(LPITEMIDLIST p = pidl; p->mkid.cb ; )
	{
		ret += p->mkid.cb;
		p = (LPITEMIDLIST)( ((BYTE*)p) + p->mkid.cb );
	}
	return ret;
}
void CShellFolderDS::construct(IShellFolderPtr sf, const STRVEC& arFiles)
{
	// sf->AddRef();
	ASSERT(!m_iSF);
	m_iSF = sf;

	{
		size_t uBuffSize=0;
		HGLOBAL hgDrop = NULL;
		CStringList lsDraggedFiles;
		for(size_t ti=0 ; ti < arFiles.size() ; ++ti)
		{
			lsDraggedFiles.AddTail(arFiles[ti].c_str());
			uBuffSize += arFiles[ti].size() + 1;
		}

		// Add 1 extra for the final null char, and the size of the DROPFILES struct.
		uBuffSize = sizeof(DROPFILES) + sizeof(TCHAR) * (uBuffSize + 1);

		// Allocate memory from the heap for the DROPFILES struct.
		hgDrop = GlobalAlloc ( GHND | GMEM_SHARE, uBuffSize );

		if ( NULL == hgDrop )
			return;

		DROPFILES* pDrop = (DROPFILES*) GlobalLock ( hgDrop );

		if ( NULL == pDrop )
		{
			GlobalFree ( hgDrop );
			return;
		}

		memset(pDrop, 0, uBuffSize);
		// Fill in the DROPFILES struct.

		pDrop->pFiles = sizeof(DROPFILES);

#ifdef _UNICODE
		// If we're compiling for Unicode, set the Unicode flag in the struct to
		// indicate it contains Unicode strings.

		pDrop->fWide = TRUE;
#endif

		// Copy all the filenames into memory after the end of the DROPFILES struct.

		POSITION pos = lsDraggedFiles.GetHeadPosition();
		LPTSTR pszBuff = (TCHAR*) (LPBYTE(pDrop) + sizeof(DROPFILES));

		while ( NULL != pos )
		{
			lstrcpy ( pszBuff, (LPCTSTR) lsDraggedFiles.GetNext ( pos ) );
			pszBuff = 1 + _tcschr ( pszBuff, '\0' );
		}

		GlobalUnlock ( hgDrop );

		// Put the data in the data source.
		FORMATETC etcHDROP = { CF_HDROP, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		m_ds.CacheGlobalData ( CF_HDROP, hgDrop, &etcHDROP );
	}



	BOOL bAdd=FALSE;
	size_t nFileCount = arFiles.size();
	if(bAdd)
	{
		static UINT CF_FILEDESCRIPTOR = RegisterClipboardFormat(CFSTR_FILEDESCRIPTOR);
		size_t nFGDsize = sizeof(UINT) + sizeof(FILEDESCRIPTOR )*nFileCount;

		HGLOBAL hgFGD = GlobalAlloc ( GHND | GMEM_SHARE, nFGDsize );
		if(!hgFGD)
		{
			ASSERT(FALSE);
			return;
		}
		FILEGROUPDESCRIPTOR* pFGD = (FILEGROUPDESCRIPTOR*)GlobalLock(hgFGD);
		if(!pFGD)
		{
			ASSERT(FALSE);
			return;
		}
		pFGD->cItems = (UINT)nFileCount;
		for(size_t i=0 ; i < nFileCount ; ++i)
		{
			FILEDESCRIPTOR fd = {0};
			getFD(arFiles[i].c_str(), &fd);
			pFGD->fgd[i] = fd;
		}
		GlobalUnlock(hgFGD);
		FORMATETC etcFILEDESCRIPTOR = { (WORD)CF_FILEDESCRIPTOR, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		m_ds.CacheGlobalData ( CF_FILEDESCRIPTOR, hgFGD, &etcFILEDESCRIPTOR );
	}


	if(bAdd)
	{
		static UINT CF_FILECONTENTS = RegisterClipboardFormat(CFSTR_FILECONTENTS);
		for(size_t i=0 ; i < nFileCount ; ++i)
		{
			size_t size = 10*1000*1000;
			HGLOBAL hgF = GlobalAlloc ( GHND | GMEM_SHARE, size );
			if(!hgF)
				return;
			BYTE* pF = (BYTE*)GlobalLock(hgF);
			if(!pF)
				return;
			CFile file(arFiles[i].c_str(),CFile::modeRead);
			file.Read((void*)pF, (UINT)size);
			GlobalUnlock(hgF);
			FORMATETC etcF = { (WORD)CF_FILECONTENTS, NULL, DVASPECT_CONTENT, (LONG)i, TYMED_HGLOBAL };
			m_ds.CacheGlobalData ( CF_FILECONTENTS, hgF, &etcF );
		}
	}

	if(bAdd)
	{
#define GetPIDLFolder(pida) (LPCITEMIDLIST)(((LPBYTE)pida)+(pida)->aoffset[0])
#define GetPIDLItem(pida, i) (LPCITEMIDLIST)(((LPBYTE)pida)+(pida)->aoffset[i+1])

		static UINT CF_SHELLIDLIST = RegisterClipboardFormat(CFSTR_SHELLIDLIST);

		LPITEMIDLIST* pAllPidls = new LPITEMIDLIST[nFileCount];
		IShellFolderPtr pDesktop;
		SHGetDesktopFolder(&pDesktop);
		for(size_t i=0 ; i < nFileCount ; ++i)
		{
			LPITEMIDLIST pIDL = NULL;
			if(FAILED(pDesktop->ParseDisplayName(
				NULL,
				NULL,
				(LPTSTR)arFiles[i].c_str(),
				NULL,
				&pIDL,
				NULL)))
			{
				ASSERT(FALSE);
				return;
			}
			pAllPidls[i] = pIDL;
		}

		size_t offsetareasize = sizeof(UINT)*(1+nFileCount); // count of offsets 1=parent
		size_t size = sizeof(UINT); // cidl
		size += offsetareasize;
		for(size_t i=0 ; i < nFileCount ; ++i)
		{
			size += getPidlSize(pAllPidls[i]);
		}

		HGLOBAL hgSIDL = GlobalAlloc ( GHND | GMEM_SHARE, size );
		if(!hgSIDL)
		{
			return;
		}
		CIDA * pCIDA  = (CIDA *)GlobalLock(hgSIDL);
		if(!pCIDA)
		{
			return;
		}


		BYTE* pOffset =  ( ((BYTE*)pCIDA) + sizeof(UINT) );
		BYTE* pContent = ( ((BYTE*)pOffset) + offsetareasize );

		pOffset[0] = 0; // desktop
		UINT contoffset = 0;
		for(size_t i=0 ; i < nFileCount ; ++i)
		{
			size_t sizePidl = getPidlSize(pAllPidls[i]);
			memcpy(pContent+contoffset, &pAllPidls[i], sizePidl);
			
			UINT u = (UINT)(pContent + contoffset - (BYTE*)pCIDA);
			pOffset[1+i] = u;

			contoffset += (UINT)sizePidl;
		}

		for(size_t i=0 ; i < nFileCount ; ++i)
		{
			CoTaskMemFree(pAllPidls[i]);
		}
		delete[] pAllPidls;


		GlobalUnlock(hgSIDL);
		FORMATETC etcSIDL = { (WORD)CF_SHELLIDLIST, NULL, DVASPECT_CONTENT, -1, TYMED_HGLOBAL };
		m_ds.CacheGlobalData ( CF_SHELLIDLIST, hgSIDL, &etcSIDL );
	}
}


