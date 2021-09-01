#include "Dll.h"
#include <Windows.h>

#include "ExplorerCommandVerb.h"

struct CEnumExplorerCommand : public IEnumExplorerCommand {
	CEnumExplorerCommand(CExplorerCommandVerb* menu) : _cRef(1)
	{
		DllAddRef();
		_menu.copy_from(menu);
	}

	winrt::com_ptr<CExplorerCommandVerb> _menu{ nullptr };


	// IUnknown
	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
	{
		static const QITAB qit[] =
		{
			QITABENT(CEnumExplorerCommand, IEnumExplorerCommand),       // required
			QITABENT(CEnumExplorerCommand, IShellExtInit),
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
	DWORD _idx = 0;

	winrt::com_ptr<IShellItemArray> GetItems() {
		winrt::com_ptr<IObjectWithSelection> sel;
		if (SUCCEEDED(_menu->QueryInterface(IID_PPV_ARGS(&sel)))) {
			winrt::com_ptr<IShellItemArray> psia;
			sel->GetSelection(IID_PPV_ARGS(&psia));
			return psia;
		}

		return nullptr;
	}
};


IFACEMETHODIMP CEnumExplorerCommand::Next(
	ULONG            /*celt*/,
	IExplorerCommand** pUICommand,
	ULONG* pceltFetched
) {

	*pceltFetched = 0;
	*pUICommand = nullptr;
	DWORD count{ 0 };
	auto psia = GetItems();
	if (psia && SUCCEEDED(psia->GetCount(&count))) {
		if (_idx < count) {
			*pceltFetched = 1;
			*pUICommand = reinterpret_cast<IExplorerCommand*>(CoTaskMemAlloc(sizeof(IExplorerCommand*) * 1));
			winrt::com_ptr<IShellItem> psi;
			if (SUCCEEDED(psia->GetItemAt(_idx++, psi.put()))) {
				wil::unique_cotaskmem_string name;
				if (SUCCEEDED(psi->GetDisplayName(SIGDN_FILESYSPATH, &name))) {
					OutputDebugStringW(name.get());
					OutputDebugStringW(L"\n");
					pUICommand[0] = new CExplorerCommandVerb(name.get());
					return S_OK;
				}
			}
		}
	}

	return E_FAIL;

}

IEnumExplorerCommand* CEnumExplorerCommand_CreateInstance(CExplorerCommandVerb* e) {
	return new CEnumExplorerCommand(e);
}
