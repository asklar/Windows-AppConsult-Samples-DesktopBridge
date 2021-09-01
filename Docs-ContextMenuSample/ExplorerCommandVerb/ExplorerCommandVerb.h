// THIS CODE AND INFORMATION IS PROVIDED "AS IS" WITHOUT WARRANTY OF
// ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT LIMITED TO
// THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS FOR A
// PARTICULAR PURPOSE.
//
// Copyright (c) Microsoft Corporation. All rights reserved

// ExplorerCommand handlers are an inproc verb implementation method that can provide
// dynamic behavior including computing the name of the command, its icon and its visibility state.
// only use this verb implemetnation method if you are implementing a command handler on
// the commands module and need the same functionality on a context menu.
//
// each ExplorerCommand handler needs to have a unique COM object, run uuidgen to
// create new CLSID values for your handler. a handler can implement multiple
// different verbs using the information provided via IInitializeCommand (the verb name).
// your code can switch off those different verb names or the properties provided
// in the property bag

#include "Dll.h"
#include <iostream>
#include <Windows.h>
#include <winrt/base.h>

class CExplorerCommandVerb;
IEnumExplorerCommand* CEnumExplorerCommand_CreateInstance(CExplorerCommandVerb* e);

class CExplorerCommandVerb : public IExplorerCommand,
	public IInitializeCommand,
	public IObjectWithSite,
	public IObjectWithSelection,
	public IShellExtInit
{
	static int num;
	int id;
public:
	CExplorerCommandVerb(const CExplorerCommandVerb&) = delete;
	CExplorerCommandVerb(CExplorerCommandVerb&&) = delete;
	
	static constexpr auto const c_szVerbDisplayName = L"See legacy context menus...";

	CExplorerCommandVerb(bool hasSubitems) : _cRef(1), _punkSite(NULL), _hwnd(NULL)
	{
		id = num++;
		_hasSubItems = hasSubitems;
		DllAddRef();
		SHStrDup(c_szVerbDisplayName, &_title);
	}

	CExplorerCommandVerb(PWSTR title) : _cRef(1)
	{
		id = num++;
		_hasSubItems = false;
		DllAddRef();
		SHStrDup(title, &_title);
	}
	// IUnknown
	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
	{
		static const QITAB qit[] =
		{
			QITABENT(CExplorerCommandVerb, IExplorerCommand),       // required
			QITABENT(CExplorerCommandVerb, IInitializeCommand),     // optional
			QITABENT(CExplorerCommandVerb, IObjectWithSite),        // optional
			QITABENT(CExplorerCommandVerb, IObjectWithSelection),
			QITABENT(CExplorerCommandVerb, IShellExtInit),
			{ 0 },
		};
		return QISearch(this, qit, riid, ppv);
	}

	IFACEMETHODIMP_(ULONG) AddRef()
	{
		return InterlockedIncrement(&_cRef);
	}

	IFACEMETHODIMP_(ULONG) Release()
	{
		long cRef = InterlockedDecrement(&_cRef);
		if (!cRef)
		{
			delete this;
		}
		return cRef;
	}

	IFACEMETHODIMP SetSelection(IShellItemArray* sia) {
		_psia.copy_from(sia);
		return S_OK;
	}

	IFACEMETHODIMP GetSelection(
		REFIID riid,
		void** ppv
	) {
		if (_psia) {
			return _psia->QueryInterface(riid, ppv);
		}

		//IShellItemArray* ows{ nullptr };
		//if (SUCCEEDED(IUnknown_QueryService(_punkSite, SID_SFolderView, IID_PPV_ARGS(&ows)))) {
		//	return ows->QueryInterface(riid, ppv);
		//}
		//
		//IShellFolderView* fv{ nullptr };
		//if (SUCCEEDED(IUnknown_QueryService(_punkSite, SID_SFolderView, IID_PPV_ARGS(&fv)))) {
		//	PCUITEMID_CHILD* children;
		//	UINT items{ 0 };
		//	if (SUCCEEDED(fv->GetSelectedObjects(&children, &items))) {
		//		return E_ACCESSDENIED;
		//	}
		//}
		//
		return E_NOTIMPL;
	}

	IFACEMETHODIMP Initialize(PCIDLIST_ABSOLUTE /*pidlFolder*/, IDataObject* pdtobj, HKEY /*progId*/) {
		return SHCreateShellItemArrayFromDataObject(pdtobj, IID_PPV_ARGS(&this->_psia));
	}

	// IExplorerCommand
	IFACEMETHODIMP GetTitle(IShellItemArray* /* psiItemArray */, LPWSTR* ppszName);

	IFACEMETHODIMP GetIcon(IShellItemArray* /* psiItemArray */, LPWSTR* ppszIcon)
	{
		// the icon ref ("dll,-<resid>") is provied here, in this case none is provieded
		*ppszIcon = NULL;
		return E_NOTIMPL;
	}

	IFACEMETHODIMP GetToolTip(IShellItemArray* /* psiItemArray */, LPWSTR* ppszInfotip)
	{
		// tooltip provided here, in this case none is provieded
		*ppszInfotip = NULL;
		return E_NOTIMPL;
	}

	IFACEMETHODIMP GetCanonicalName(GUID* pguidCommandName)
	{
		*pguidCommandName = __uuidof(this);
		return S_OK;
	}

	// compute the visibility of the verb here, respect "fOkToBeSlow" if this is slow (does IO for example)
	// when called with fOkToBeSlow == FALSE return E_PENDING and this object will be called
	// back on a background thread with fOkToBeSlow == TRUE
	IFACEMETHODIMP GetState(IShellItemArray*  psiItemArray , BOOL /*fOkToBeSlow*/, EXPCMDSTATE* pCmdState)
	{
		*pCmdState = ECS_ENABLED;
		_psia = nullptr;
		HRESULT hr = psiItemArray->QueryInterface(IID_PPV_ARGS(&_psia));
		return hr;
	}

	IFACEMETHODIMP Invoke(IShellItemArray* psiItemArray, IBindCtx* pbc);

	IFACEMETHODIMP GetFlags(EXPCMDFLAGS* pFlags)
	{
		*pFlags = ECF_DEFAULT |ECF_HASSUBCOMMANDS;
		return S_OK;
	}

	IFACEMETHODIMP EnumSubCommands(IEnumExplorerCommand** ppEnum)
	{
		if (_hasSubItems) {
			*ppEnum = CEnumExplorerCommand_CreateInstance(this);
			return S_OK;
		}
		else {
			*ppEnum = nullptr;
			return E_NOTIMPL;
		}
	}

	// IInitializeCommand
	IFACEMETHODIMP Initialize(PCWSTR /* pszCommandName */, IPropertyBag* /* ppb */)
	{
		// the verb name is in pszCommandName, this handler can vary its behavior
		// based on the command name (implementing different verbs) or the
		// data stored under that verb in the registry can be read via ppb
		return S_OK;
	}

	// IObjectWithSite
	IFACEMETHODIMP SetSite(IUnknown* punkSite)
	{
		_punkSite.copy_from(punkSite);
		return S_OK;
	}

	IFACEMETHODIMP GetSite(REFIID riid, void** ppv)
	{
		*ppv = NULL;
		return _punkSite ? _punkSite->QueryInterface(riid, ppv) : E_FAIL;
	}

private:

	~CExplorerCommandVerb()
	{
		_punkSite = nullptr;
		DllRelease();
	}

	bool _hasSubItems;
	wil::unique_cotaskmem_string _title;

	long _cRef;
	winrt::com_ptr<IUnknown> _punkSite{ nullptr };
	HWND _hwnd{ 0 };
	//IStream* _pstmShellItemArray{ nullptr };
	winrt::com_ptr<IShellItemArray> _psia{ nullptr };
};


