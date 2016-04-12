// XEncoding.h : Declaration of the CXEncoding

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"
#include "atlib.h"


// CXEncoding

class ATL_NO_VTABLE CXEncoding : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXEncoding, &CLSID_XEncoding>,
	public IDispatchImpl<IXEncoding, &IID_IXEncoding, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CXEncoding()
	{
		m_pUnkMarshaler = NULL;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_XENCODING)


BEGIN_COM_MAP(CXEncoding)
	COM_INTERFACE_ENTRY(IXEncoding)
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
	STDMETHOD(Base32Decode)(BSTR base32String, VARIANT *retVal);
	STDMETHOD(Base32Encode)(VARIANT varData, BSTR *retVal);

	STDMETHOD(Base64Decode)(BSTR base64String, VARIANT *retVal);
	STDMETHOD(Base64Encode)(VARIANT varData, BSTR *retVal);

	STDMETHOD(HtmlDecode)(BSTR HtmlString, BSTR* retVal);
	STDMETHOD(HtmlFormat)(BSTR TextString, VARIANT varFormat, BSTR* retVal);

	STDMETHOD(JSEncode)(BSTR TextString, BSTR *retVal);

	STDMETHOD(UrlDecode)(BSTR urlString, BSTR *retVal);
	STDMETHOD(UrlEncode)(BSTR urlString, VARIANT bDBCS, BSTR *retVal);

	STDMETHOD(QuotedDecode)(BSTR QuotedString, VARIANT nCodePage, BSTR* retVal);
	STDMETHOD(QuotedEncode)(BSTR txtString, VARIANT nCodePage, BSTR* retVal);

	static CXString UrlDecode(LPCSTR urlString, UINT len);
	static CXString UrlDecode(CXStringA str)
	{
		return UrlDecode(str, str.GetLength());
	}

	static CXStringA UrlEncode(LPCSTR urlString, UINT len, BOOL bEncodeDBMS = FALSE);
	static CXStringA UrlEncode(CXStringA str, BOOL bEncodeDBMS = FALSE)
	{
		return UrlEncode(str, str.GetLength(), bEncodeDBMS);
	}

	STDMETHOD(GBK2GB)(BSTR gbkStr, BSTR *gbStr);

	STDMETHOD(BinToStr)(VARIANT varData, VARIANT varCodePage, BSTR *retVal);
	STDMETHOD(StrToBin)(BSTR strData, VARIANT varCodePage, VARIANT *retVal);

	STDMETHOD(GMTDecode)(BSTR GMTString, DATE* retVal);
	STDMETHOD(GMTEncode)(DATE varDate, BSTR* retVal);

	STDMETHOD(HexDecode)(BSTR HexString, VARIANT *retVal);
	STDMETHOD(HexEncode)(VARIANT varData, BSTR *retVal);

	STDMETHOD(XCodeEncode)(BSTR codeString, BSTR key, BSTR *retVal);
	STDMETHOD(XCodeDecode)(BSTR codeString, VARIANT key, BSTR *retVal);

	STDMETHOD(XmlEncode)(BSTR txtString, BSTR* retVal);
	STDMETHOD(XmlFilter)(BSTR txtString, BSTR* retVal);

	STDMETHOD(GetKeyword)(BSTR txtString, VARIANT attr, VARIANT* pVal);

	STDMETHOD(FixString)(BSTR txtString, short nLen, BSTR* pVal);
	STDMETHOD(FixStringLen)(BSTR txtString, LONG* pVal);

	STDMETHOD(FindString)(VARIANT varStart, BSTR txtString, VARIANT varFind, LONG* pVal);

	STDMETHOD(FixSpace)(BSTR txtString, BSTR* pVal);

	STDMETHOD(FormatMessage)(BSTR bstrFmtString, 
		VARIANT bstr1, VARIANT bstr2, VARIANT bstr3, VARIANT bstr4, VARIANT bstr5, VARIANT bstr6, VARIANT bstr7, VARIANT bstr8,
		VARIANT bstra, VARIANT bstrb, VARIANT bstrc, VARIANT bstrd, VARIANT bstre, VARIANT bstrf, VARIANT bstrg, VARIANT bstrh,
		VARIANT bstri, VARIANT bstrj, VARIANT bstrk, VARIANT bstrl, VARIANT bstrm, VARIANT bstrn, VARIANT bstro, VARIANT bstrp,
		VARIANT bstrq, VARIANT bstrr, VARIANT bstrs, VARIANT bstrt, VARIANT bstru, VARIANT bstrv, VARIANT bstrw, VARIANT bstrx,
		VARIANT bstry, VARIANT bstrz,
		VARIANT bstr9, BSTR *bstrMessage);
};

OBJECT_ENTRY_AUTO(__uuidof(XEncoding), CXEncoding)
