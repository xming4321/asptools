// XSession.h : Declaration of the CXSession

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"
#include "XSessionMan.h"
#include <activscp.h>


// CXSession

class ATL_NO_VTABLE CXSession : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXSession, &CLSID_XSession>,
	public IDispatchImpl<IXSession, &IID_IXSession, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>,
	public IActiveScriptSite
{
public:

DECLARE_REGISTRY_RESOURCEID(IDR_XSESSION)


BEGIN_COM_MAP(CXSession)
	COM_INTERFACE_ENTRY(IXSession)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY(IActiveScriptSite)
END_COM_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()
	DECLARE_GET_CONTROLLING_UNKNOWN()
	
	HRESULT FinalConstruct()
	{
		m_bIsNewSession = FALSE;
		return S_OK;
	}

	void FinalRelease();

public:
	// IActiveScriptSite
	STDMETHOD(GetLCID)(LCID *plcid);
	STDMETHOD(GetItemInfo)(LPCOLESTR pstrName, DWORD dwReturnMask, IUnknown **ppiunkItem, ITypeInfo **ppti);
	STDMETHOD(GetDocVersionString)(BSTR *pbstrVersion);
	STDMETHOD(OnScriptTerminate)(const VARIANT *pvarResult, const EXCEPINFO *pexcepinfo);
	STDMETHOD(OnStateChange)(SCRIPTSTATE ssScriptState);
	STDMETHOD(OnScriptError)(IActiveScriptError *pscripterror);
	STDMETHOD(OnEnterScript)(void);
	STDMETHOD(OnLeaveScript)(void);

// IXSession
public:
	//Active Server Pages Methods
	STDMETHOD(OnStartPage)(IUnknown* IUnk);
	STDMETHOD(OnEndPage)();

public:
	STDMETHOD(get_Item)(VARIANT key, VARIANT* pVal);
	STDMETHOD(put_Item)(VARIANT key, VARIANT newVal);
	STDMETHOD(putref_Item)(VARIANT key, VARIANT newVal);
	STDMETHOD(get_LastAccessTime)(DATE* pVal);
	STDMETHOD(get_LastModifyTime)(DATE* pVal);
	STDMETHOD(get_AccessTimes)(long* pVal);
	STDMETHOD(get_Contents)(IXDictionary** pVal);
	STDMETHOD(get_Count)(LONG* pVal);
	STDMETHOD(get_SessionID)(BSTR* pVal);
	STDMETHOD(get_isActived)(VARIANT_BOOL* pVal);
	STDMETHOD(get_isLocal)(VARIANT_BOOL* pVal);
	STDMETHOD(isNewSession)(VARIANT_BOOL* pVal);
	STDMETHOD(get_Application)(IDispatch** pVal);
	STDMETHOD(Remove)(VARIANT key);
	STDMETHOD(RemoveAll)(void);
	STDMETHOD(GetSession)(BSTR bstrID);
	STDMETHOD(Suspend)(void);
	STDMETHOD(Update)(void);
	STDMETHOD(Abandon)(void);
	STDMETHOD(Execute)(BSTR strEvent, VARIANT v1, VARIANT v2, VARIANT v3, VARIANT v4, VARIANT v5, VARIANT v6, VARIANT v7, VARIANT v8, VARIANT* pVal);
	STDMETHOD(Rnd)(long* pVal);

	void Attach(CXSessionMan::CXSessionInst* pInst, IRequest* spRequest, IResponse* spResponse)
	{
		m_pInstance = pInst;
		m_spRequest = spRequest;
		m_spResponse = spResponse;
	}

	void Detach()
	{
		m_pInstance.Release();
		m_spRequest.Release();
		m_spResponse.Release();
	}

	int LoadGlobal();
	void ClearGlobal();
	HRESULT scriptError(HRESULT hr, LPCWSTR strEvent, EXCEPINFO* pei);
	HRESULT DoEvent(LPCWSTR strEvent);
	HRESULT FireEvent(LPCWSTR strEvent, VARIANT *pvar, int size);
	HRESULT SyncEvent(BYTE * ptr, int size);

private:
	CXComPtr<CXSessionMan::CXSessionInst> m_pInstance;
	BOOL m_bIsNewSession;
	DWORD m_nErrorLine;

	CComPtr<IRequest> m_spRequest;
	CComPtr<IResponse> m_spResponse;
	CXComPtr<IActiveScript> m_pActiveScript;

public:
	CComDispatchDriver m_pDispGlobal;
};

OBJECT_ENTRY_AUTO(__uuidof(XSession), CXSession)
