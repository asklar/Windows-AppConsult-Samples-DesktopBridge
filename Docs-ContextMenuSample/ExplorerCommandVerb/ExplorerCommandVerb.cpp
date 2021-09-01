#include "ExplorerCommandVerb.h"
#include <wil/resource.h>

static WCHAR const c_szVerbName[] = L"Sample.ExplorerCommandVerb";


DWORD _ThreadProc(void* pv)
{
	winrt::com_ptr<IStream> pstm;
	pstm.attach(reinterpret_cast<IStream*>(pv));

	winrt::com_ptr<IShellItemArray> psia;
	HRESULT hr = CoGetInterfaceAndReleaseStream(pstm.get(), IID_PPV_ARGS(&psia));

	if (SUCCEEDED(hr))
	{

		DWORD count;
		psia->GetCount(&count);

		winrt::com_ptr<IShellItem2> psi;
		hr = GetItemAt(psia.get(), 0, IID_PPV_ARGS(&psi));
		if (SUCCEEDED(hr))
		{
			wil::unique_cotaskmem_string pszName;
			
			hr = psi->GetDisplayName(SIGDN_DESKTOPABSOLUTEPARSING, &pszName);
			if (SUCCEEDED(hr))
			{
				WCHAR szMsg[128];
				StringCchPrintf(szMsg, ARRAYSIZE(szMsg), L"%d item(s), first item is named %s", count, pszName.get());

				MessageBox(nullptr, szMsg, L"ExplorerCommand Sample Verb", MB_OK);
			}
		}
	}

	return 0;
}


IFACEMETHODIMP CExplorerCommandVerb::Invoke(IShellItemArray* psia, IBindCtx* /* pbc */)
{
	IUnknown_GetWindow(_punkSite.get(), &_hwnd);
	
	winrt::com_ptr<IStream> pstm;
	HRESULT hr = CoMarshalInterThreadInterfaceInStream(__uuidof(psia), psia, pstm.put());
	if (SUCCEEDED(hr))
	{
		AddRef();
		if (!SHCreateThread(_ThreadProc, pstm.detach(), CTF_COINIT_STA | CTF_PROCESS_REF, NULL))
		{
			Release();
		}
	}
	return S_OK;
}

static WCHAR const c_szProgID[] = L"*";// L"txtfile";

HRESULT __cdecl CExplorerCommandVerb_RegisterUnRegister(bool fRegister)
{
	CRegisterExtension re(__uuidof(CExplorerCommandVerb));

	HRESULT hr;
	if (fRegister)
	{
		hr = re.RegisterInProcServer(CExplorerCommandVerb::c_szVerbDisplayName, L"Apartment");
		if (SUCCEEDED(hr))
		{
			// register this verb on .txt files ProgID
			hr = re.RegisterExplorerCommandVerb(c_szProgID, c_szVerbName, CExplorerCommandVerb::c_szVerbDisplayName);
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
	return SHStrDup(_title.get(), ppszName);
}

int CExplorerCommandVerb::num = 0;