// XRsa.h : Declaration of the CXRsa

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"
#include "atlib.h"


// CXRsa

class ATL_NO_VTABLE CXRsa : 
	public CComObjectRootEx<CComSingleThreadModel>,
	public CComCoClass<CXRsa, &CLSID_XRsa>,
	public IDispatchImpl<IXRsa, &IID_IXRsa, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CXRsa() : m_pRSA(NULL), m_nPadding(1)
	{
	}

DECLARE_REGISTRY_RESOURCEID(IDR_XRSA)

BEGIN_COM_MAP(CXRsa)
	COM_INTERFACE_ENTRY(IXRsa)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}
	
	void FinalRelease() 
	{
		free();
	}

public:
	STDMETHOD(get_IsPrivateKey)(VARIANT_BOOL *pVal);
	STDMETHOD(get_Key)(VARIANT *pVal);
	STDMETHOD(get_KeySize)(short *pVal);
	STDMETHOD(get_Padding)(short *pVal);
	STDMETHOD(put_Padding)(short newVal);
	STDMETHOD(get_PrivateKey)(VARIANT *pVal);
	STDMETHOD(put_PrivateKey)(VARIANT newVal);
	STDMETHOD(get_PublicKey)(VARIANT *pVal);
	STDMETHOD(put_PublicKey)(VARIANT newVal);
	STDMETHOD(Decrypt)(VARIANT varData, VARIANT *pVal);
	STDMETHOD(Encrypt)(VARIANT varData, VARIANT *pVal);
	STDMETHOD(GenerateKey)(VARIANT varSize);

private:
	void *m_pRSA;
	short m_nPadding;

	void free(void);
};

OBJECT_ENTRY_AUTO(__uuidof(XRsa), CXRsa)
