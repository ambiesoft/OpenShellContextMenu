#pragma once
class CContextMenuCB : public IContextMenuCB
{
	virtual HRESULT STDMETHODCALLTYPE QueryInterface(
		REFIID riid,
		void **ppvObject)
	{
		if (IsEqualIID(IID_IContextMenuCB, riid))
		{
			*ppvObject = this;
			return S_OK;
		}
		return E_NOINTERFACE;
	}
	virtual ULONG STDMETHODCALLTYPE AddRef(void){ return 1; }
	virtual ULONG STDMETHODCALLTYPE Release(void){ return 1; }

	virtual HRESULT STDMETHODCALLTYPE CallBack(
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
		LPARAM lParam);
};