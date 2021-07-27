#include "Dll.h"
#include <Windows.h>

#include "ExplorerCommandVerb.h"

struct CEnumExplorerCommand : public IEnumExplorerCommand {
	CEnumExplorerCommand() : _cRef(1)
	{
		DllAddRef();
	}

	// IUnknown
	IFACEMETHODIMP QueryInterface(REFIID riid, void** ppv)
	{
		static const QITAB qit[] =
		{
			QITABENT(CEnumExplorerCommand , IEnumExplorerCommand),       // required
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
};


IFACEMETHODIMP CEnumExplorerCommand::Next(
	ULONG            /*celt*/,
	IExplorerCommand** pUICommand,
	ULONG* pceltFetched
) {
	if (idx++ == 0) {
		*pceltFetched = 1;
		*pUICommand = new CExplorerCommandVerb(false);
		return S_OK;
	}
	else {
		return S_FALSE;
	}
}

IEnumExplorerCommand* CEnumExplorerCommand_CreateInstance() {
	return new CEnumExplorerCommand();
}
