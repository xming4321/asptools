// XDictionary.h : Declaration of the CXDictionary

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"
#include "atlib.h"


// CXDictionary

class ATL_NO_VTABLE CXDictionary : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXDictionary, &CLSID_XDictionary>,
	public IDispatchImpl<IXDictionary, &IID_IXDictionary, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CXDictionary()
	{
		m_pUnkMarshaler = NULL;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_XDICTIONARY)


BEGIN_COM_MAP(CXDictionary)
	COM_INTERFACE_ENTRY(IXDictionary)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
END_COM_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()
	DECLARE_GET_CONTROLLING_UNKNOWN()

	HRESULT FinalConstruct()
	{
		return CoCreateFreeThreadedMarshaler(
			GetControllingUnknown(), &m_pUnkMarshaler.p);
	}

	void FinalRelease()
	{
		m_pUnkMarshaler.Release();
	}

	CComPtr<IUnknown> m_pUnkMarshaler;

public:
	STDMETHOD(get_Item)(VARIANT key, VARIANT* pVal);
	STDMETHOD(put_Item)(VARIANT key, VARIANT newVal);
	STDMETHOD(putref_Item)(VARIANT key, VARIANT newVal);
	STDMETHOD(Count)(LONG* pVal);
	STDMETHOD(Remove)(VARIANT key);
	STDMETHOD(RemoveAll)(void);
	STDMETHOD(Add)(VARIANT key, VARIANT value);
	STDMETHOD(Exists)(VARIANT key, VARIANT_BOOL* pVal);
	STDMETHOD(Join)(VARIANT strKeyDelimiter, VARIANT strDelimiter, BSTR* pVal);
	STDMETHOD(Split)(BSTR bstrExpression, VARIANT strKeyDelimiter, VARIANT strDelimiter);
	STDMETHOD(get_Keys)(VARIANT* pVal);
	STDMETHOD(get_Items)(VARIANT* pVal);
	STDMETHOD(get__NewEnum)(IUnknown** pVal);

public:
	void FillEnum(CAtlArray<VARIANT>& arrayVariant);

private:
	CCriticalSection m_cs;
	CRBMap<CXVariant, CComVariant> m_mapItems;

	HRESULT putItem(VARIANT* pvarKey, VARIANT* pvar);
};

OBJECT_ENTRY_AUTO(__uuidof(XDictionary), CXDictionary)
