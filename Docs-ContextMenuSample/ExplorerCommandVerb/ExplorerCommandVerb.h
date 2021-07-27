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

IEnumExplorerCommand* CEnumExplorerCommand_CreateInstance();

class CExplorerCommandVerb : public IExplorerCommand,
	public IInitializeCommand,
	public IObjectWithSite
{
public:
	CExplorerCommandVerb(bool hasSubitems) : _cRef(1), _punkSite(NULL), _hwnd(NULL), _pstmShellItemArray(NULL)
	{
		_hasSubItems = hasSubitems;
		DllAddRef();
	}

	// IUnknown
	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
	{
		static const QITAB qit[] =
		{
			QITABENT(CExplorerCommandVerb, IExplorerCommand),       // required
			QITABENT(CExplorerCommandVerb, IInitializeCommand),     // optional
			QITABENT(CExplorerCommandVerb, IObjectWithSite),        // optional
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
	IFACEMETHODIMP GetState(IShellItemArray* /* psiItemArray */, BOOL /*fOkToBeSlow*/, EXPCMDSTATE* pCmdState)
	{
		//HRESULT hr;
		//if (fOkToBeSlow)
		//{
		//    Sleep(4 * 1000);    // simulate expensive work
		//    *pCmdState = ECS_ENABLED;
		//    hr = S_OK;
		//}
		//else
		//{
		//    *pCmdState = ECS_DISABLED;
		//    // returning E_PENDING requests that a new instance of this object be called back
		//    // on a background thread so that it can do work that might be slow
		//    hr = E_PENDING;
		//}
		*pCmdState = ECS_ENABLED;

		HRESULT hr = S_OK;
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
			*ppEnum = CEnumExplorerCommand_CreateInstance();
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
		SetInterface(&_punkSite, punkSite);
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
		SafeRelease(&_punkSite);
		SafeRelease(&_pstmShellItemArray);
		DllRelease();
	}

	DWORD _ThreadProc();
	bool _hasSubItems;
	static DWORD __stdcall s_ThreadProc(void* pv)
	{
		CExplorerCommandVerb* pecv = (CExplorerCommandVerb*)pv;
		const DWORD ret = pecv->_ThreadProc();
		pecv->Release();
		return ret;
	}

	long _cRef;
	IUnknown* _punkSite;
	HWND _hwnd;
	IStream* _pstmShellItemArray;
};


