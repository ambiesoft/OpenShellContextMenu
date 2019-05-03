#pragma once

//class CShellDataSouce: public COleDataSource
//{
//};

class CShellFolderDS : public CCmdTarget
{
public:
   CShellFolderDS()
   {
	   m_iSF = NULL;
   }
	void construct(IShellFolderPtr sf, const STRVEC& arFiles);
	virtual ~CShellFolderDS()
	{
		//if(m_iSF)
		//{
		//	m_iSF->Release(); 
		//	m_iSF=NULL;
		//}
	}


public:
    DECLARE_INTERFACE_MAP()

    BEGIN_INTERFACE_PART(ShellFolderDS, IShellFolder)
        virtual HRESULT STDMETHODCALLTYPE ParseDisplayName( 
            /* [in] */ HWND hwnd,
            /* [in] */ LPBC pbc,
            /* [string][in] */ LPOLESTR pszDisplayName,
            /* [out] */ ULONG *pchEaten,
            /* [out] */ LPITEMIDLIST *ppidl,
            /* [unique][out][in] */ ULONG *pdwAttributes);
        
        virtual HRESULT STDMETHODCALLTYPE EnumObjects( 
            /* [in] */ HWND hwnd,
            /* [in] */ SHCONTF grfFlags,
            /* [out] */ IEnumIDList **ppenumIDList);
        
        virtual HRESULT STDMETHODCALLTYPE BindToObject( 
            /* [in] */ LPCITEMIDLIST pidl,
            /* [in] */ LPBC pbc,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppv);
        
        virtual HRESULT STDMETHODCALLTYPE BindToStorage( 
            /* [in] */ LPCITEMIDLIST pidl,
            /* [in] */ LPBC pbc,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppv);
        
        virtual HRESULT STDMETHODCALLTYPE CompareIDs( 
            /* [in] */ LPARAM lParam,
            /* [in] */ LPCITEMIDLIST pidl1,
            /* [in] */ LPCITEMIDLIST pidl2);
        
        virtual HRESULT STDMETHODCALLTYPE CreateViewObject( 
            /* [in] */ HWND hwndOwner,
            /* [in] */ REFIID riid,
            /* [iid_is][out] */ void **ppv);
        
        virtual HRESULT STDMETHODCALLTYPE GetAttributesOf( 
            /* [in] */ UINT cidl,
            /* [size_is][in] */ LPCITEMIDLIST *apidl,
            /* [out][in] */ SFGAOF *rgfInOut);
        
        virtual HRESULT STDMETHODCALLTYPE GetUIObjectOf( 
            /* [in] */ HWND hwndOwner,
            /* [in] */ UINT cidl,
            /* [size_is][in] */ LPCITEMIDLIST *apidl,
            /* [in] */ REFIID riid,
            /* [unique][out][in] */ UINT *rgfReserved,
            /* [iid_is][out] */ void **ppv);
        
        virtual HRESULT STDMETHODCALLTYPE GetDisplayNameOf( 
            /* [in] */ LPCITEMIDLIST pidl,
            /* [in] */ SHGDNF uFlags,
            /* [out] */ STRRET *pName);
        
        virtual HRESULT STDMETHODCALLTYPE SetNameOf( 
            /* [in] */ HWND hwnd,
            /* [in] */ LPCITEMIDLIST pidl,
            /* [string][in] */ LPCOLESTR pszName,
            /* [in] */ SHGDNF uFlags,
            /* [out] */ LPITEMIDLIST *ppidlOut);
    END_INTERFACE_PART(ShellFolderDS)



protected:
   IShellFolderPtr m_iSF; // the real underwriter
   COleDataSource m_ds; // that's mine with proper HDROP

public:
	void qqq(void** ppv)
	{
		m_ds.InternalQueryInterface(&IID_IDataObject, ppv);
	}
};
 
