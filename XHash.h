// XHash.h : Declaration of the CXHash

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"
#include "atlib.h"

// CXHash

class ATL_NO_VTABLE CXHash : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXHash, &CLSID_XHash>,
	public IDispatchImpl<IXHash, &IID_IXHash, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CXHash() : m_iAlgo(-1)
	{}

DECLARE_REGISTRY_RESOURCEID(IDR_XHASH)


BEGIN_COM_MAP(CXHash)
	COM_INTERFACE_ENTRY(IXHash)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
	}

public:
	STDMETHOD(get_Name)(BSTR *pVal);
	STDMETHOD(get_HashSize)(short *pVal);
	STDMETHOD(Create)(BSTR bstrAlgo = L"MD5");
	STDMETHOD(Update)(VARIANT varData);
	STDMETHOD(Final)(VARIANT varData, VARIANT *retVal);

private:
	BYTE m_ctx[256];
	int m_iAlgo;
};

OBJECT_ENTRY_AUTO(__uuidof(XHash), CXHash)
