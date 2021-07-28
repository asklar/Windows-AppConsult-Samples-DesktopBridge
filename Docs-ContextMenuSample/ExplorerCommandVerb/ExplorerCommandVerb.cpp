#include "ExplorerCommandVerb.h"

static WCHAR const c_szVerbDisplayName[] = L"See legacy context menus...";
static WCHAR const c_szVerbName[] = L"Sample.ExplorerCommandVerb";

DWORD CExplorerCommandVerb::_ThreadProc()
{
	IShellItemArray* psia;
	HRESULT hr = CoGetInterfaceAndReleaseStream(_pstmShellItemArray, IID_PPV_ARGS(&psia));
	_pstmShellItemArray = NULL;
	if (SUCCEEDED(hr))
	{
		
		DWORD count;
		psia->GetCount(&count);

		IShellItem2* psi;
		hr = GetItemAt(psia, 0, IID_PPV_ARGS(&psi));
		if (SUCCEEDED(hr))
		{
			PWSTR pszName;
			hr = psi->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, &pszName);
			if (SUCCEEDED(hr))
			{
				WCHAR szMsg[128];
				StringCchPrintf(szMsg, ARRAYSIZE(szMsg), L"%d item(s), first item is named %s", count, pszName);

				MessageBox(_hwnd, szMsg, L"ExplorerCommand Sample Verb", MB_OK);

				CoTaskMemFree(pszName);
			}

			psi->Release();
		}
		psia->Release();
	}

	return 0;
}

IFACEMETHODIMP CExplorerCommandVerb::Invoke(IShellItemArray* psia, IBindCtx* /* pbc */)
{
	IUnknown_GetWindow(_punkSite, &_hwnd);
	
	HRESULT hr = CoMarshalInterThreadInterfaceInStream(__uuidof(psia), psia, &_pstmShellItemArray);
	if (SUCCEEDED(hr))
	{
		AddRef();
		if (!SHCreateThread(s_ThreadProc, this, CTF_COINIT_STA | CTF_PROCESS_REF, NULL))
		{
			Release();
		}
	}
	return S_OK;
}

static WCHAR const c_szProgID[] = L"txtfile";

HRESULT __cdecl CExplorerCommandVerb_RegisterUnRegister(bool fRegister)
{
	CRegisterExtension re(__uuidof(CExplorerCommandVerb));

	HRESULT hr;
	if (fRegister)
	{
		hr = re.RegisterInProcServer(c_szVerbDisplayName, L"Apartment");
		if (SUCCEEDED(hr))
		{
			// register this verb on .txt files ProgID
			hr = re.RegisterExplorerCommandVerb(c_szProgID, c_szVerbName, c_szVerbDisplayName);
			if (SUCCEEDED(hr))
			{
				hr = re.RegisterVerbAttribute(c_szProgID, c_szVerbName, L"NeverDefault");
			}
		}
	}
	else
	{
		// best effort
		hr = re.UnRegisterVerb(c_szProgID, c_szVerbName);
		hr = re.UnRegisterObject();
	}
	return hr;
}

HRESULT __cdecl CExplorerCommandVerb_CreateInstance(REFIID riid, void** ppv)
{
	*ppv = NULL;
	CExplorerCommandVerb* pVerb = new (std::nothrow) CExplorerCommandVerb(true);
	HRESULT hr = pVerb ? S_OK : E_OUTOFMEMORY;
	if (SUCCEEDED(hr))
	{
		pVerb->QueryInterface(riid, ppv);
		pVerb->Release();
	}
	return hr;
}

IFACEMETHODIMP CExplorerCommandVerb::GetTitle(IShellItemArray* /* psiItemArray */, LPWSTR* ppszName)
{
	// the verb name can be computed here, in this example it is static
	return SHStrDup(c_szVerbDisplayName, ppszName);
}

int CExplorerCommandVerb::num = 0;