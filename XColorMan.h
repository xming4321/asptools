// XColorMan.h : Declaration of the CXColorMan

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"


// CXColorMan

class ATL_NO_VTABLE CXColorMan : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXColorMan, &CLSID_XColorMan>,
	public IDispatchImpl<IXColorMan, &IID_IXColorMan, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CXColorMan() :
		r(0), g(0), b(0),
		h(0), s(0), l(0)
	{}

DECLARE_REGISTRY_RESOURCEID(IDR_XCOLORMAN)


BEGIN_COM_MAP(CXColorMan)
	COM_INTERFACE_ENTRY(IXColorMan)
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

	STDMETHOD(get_R)(SHORT* pVal);
	STDMETHOD(put_R)(SHORT newVal);
	STDMETHOD(get_G)(SHORT* pVal);
	STDMETHOD(put_G)(SHORT newVal);
	STDMETHOD(get_B)(SHORT* pVal);
	STDMETHOD(put_B)(SHORT newVal);
	STDMETHOD(get_H)(SHORT* pVal);
	STDMETHOD(put_H)(SHORT newVal);
	STDMETHOD(get_S)(SHORT* pVal);
	STDMETHOD(put_S)(SHORT newVal);
	STDMETHOD(get_L)(SHORT* pVal);
	STDMETHOD(put_L)(SHORT newVal);
	STDMETHOD(get_RGB)(LONG* pVal);
	STDMETHOD(put_RGB)(LONG newVal);
	STDMETHOD(get_RGBString)(BSTR* pVal);
	STDMETHOD(put_RGBString)(BSTR newVal);

private:
	void toHSL();
	void toRGB();

private:
	short r, g, b;
	short h, s, l;
};

OBJECT_ENTRY_AUTO(__uuidof(XColorMan), CXColorMan)
