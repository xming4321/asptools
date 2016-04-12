// XHttp.h : Declaration of the CXHttp

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"
#include "atlib.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CXHttp

class ATL_NO_VTABLE CXHttp :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXHttp, &CLSID_XHttp>,
	public IDispatchImpl<IXHttp, &IID_IXHttp, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CXHttp()
	{
		m_pUnkMarshaler = NULL;

		WSADATA wsd;
		WSAStartup(MAKEWORD(2,2),&wsd);
	}

DECLARE_REGISTRY_RESOURCEID(IDR_XHTTP)


BEGIN_COM_MAP(CXHttp)
	COM_INTERFACE_ENTRY(IXHttp)
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

	STDMETHOD(Get)(BSTR url, VARIANT varPost, BSTR* retVal);

private:
	SOCKET getConnection(LPSTR strHost);
	void putConnection(LPSTR strHost, SOCKET s);

private:
	CCriticalSection m_cs;
	CRBMultiMap<CStringA, SOCKET> m_mapConnPool;

};

OBJECT_ENTRY_AUTO(__uuidof(XHttp), CXHttp)
