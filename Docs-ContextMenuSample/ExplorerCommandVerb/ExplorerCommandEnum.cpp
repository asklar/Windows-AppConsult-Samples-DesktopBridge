#include "Dll.h"
#include <Windows.h>

#include "ExplorerCommandVerb.h"

struct CEnumExplorerCommand : public IEnumExplorerCommand {
	CEnumExplorerCommand(CExplorerCommandVerb* menu) : _cRef(1)
	{
		DllAddRef();

		_menu = menu;
		menu->AddRef();
	}

	CExplorerCommandVerb* _menu{ nullptr };
	// IUnknown
	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
	{
		static const QITAB qit[] =
		{
			QITABENT(CEnumExplorerCommand, IEnumExplorerCommand),       // required
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

	IFACEMETHODIMP Clone(IEnumExplorerCommand** ppEnum) {
		*ppEnum = nullptr;
		return E_NOTIMPL;
	}

	IFACEMETHODIMP Next(
		ULONG            celt,
		IExplorerCommand** pUICommand,
		ULONG* pceltFetched
	);

	IFACEMETHODIMP Reset() { return E_NOTIMPL; }

	IFACEMETHODIMP Skip(ULONG /*cElt*/) { return E_NOTIMPL; }

private:
	long _cRef;
	int idx = 0;

	IShellItemArray* GetItems() {
		IObjectWithSelection* sel{ nullptr };
		IShellItemArray* psia{ nullptr };
		if (SUCCEEDED(_menu->QueryInterface(IID_PPV_ARGS(&sel)))) {
			sel->GetSelection(IID_PPV_ARGS(&psia));
		}
		return psia;
	}

	HKEY progId{ (HKEY)0 };
};


IFACEMETHODIMP CEnumExplorerCommand::Next(
	ULONG            /*celt*/,
	IExplorerCommand** pUICommand,
	ULONG* pceltFetched
) {

	DWORD count{ 0 };
	auto psia = GetItems();
	if (psia && SUCCEEDED(psia->GetCount(&count))) {
		for (DWORD i = 0; i < count; i++) {
			IShellItem* psi;
			if (SUCCEEDED(psia->GetItemAt(i, &psi))) {
				PWSTR name{ nullptr };
				if (SUCCEEDED(psi->GetDisplayName(SIGDN_FILESYSPATH, &name))) {
					OutputDebugStringW(name);
					CoTaskMemFree(name);
				}
				psi->Release();
			}
		}
	}

	if (idx++ == 0) {
		*pceltFetched = 1;
		*pUICommand = new CExplorerCommandVerb(false);
		return S_OK;
	}
	else {
		return S_FALSE;
	}

}

IEnumExplorerCommand* CEnumExplorerCommand_CreateInstance(CExplorerCommandVerb* e) {
	return new CEnumExplorerCommand(e);
}
