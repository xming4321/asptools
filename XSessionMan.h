// XSessionMan.h : Declaration of the CXSessionMan

#pragma once
#include "resource.h"       // main symbols
#include <asptlb.h>         // Active Server Pages Definitions

#include "asptools.h"
#include "atlib.h"
#include "XDictionary.h"
#include "XRecords.h"


// CXSessionMan

class ATL_NO_VTABLE CXSessionMan : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXSessionMan, &CLSID_XSessionMan>,
	public IDispatchImpl<IXSessionMan, &IID_IXSessionMan, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CXSessionMan()
	{
		m_pUnkMarshaler = NULL;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_XSESSIONMAN)


BEGIN_COM_MAP(CXSessionMan)
	COM_INTERFACE_ENTRY(IXSessionMan)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
END_COM_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()
	DECLARE_GET_CONTROLLING_UNKNOWN()

	HRESULT FinalConstruct();
	void FinalRelease();

	CComPtr<IUnknown> m_pUnkMarshaler;

public:
	class CXSessionInst : public CXDispatch<IXSession>
	{
	public:
		CXSessionInst();
		~CXSessionInst()
		{
			int i;

			for(i = 0; i < (int)m_arrayVariant.GetCount(); i ++)
				::VariantClear(&m_arrayVariant[i]);

			for(i = 0; i < (int)m_arrayIndexVariant.GetCount(); i ++)
				::VariantClear(&m_arrayIndexVariant[i]);
		}

	public:
		//Active Server Pages Methods
		STDMETHOD(OnStartPage)(IUnknown* IUnk)
		{	return E_NOTIMPL;}
		STDMETHOD(OnEndPage)()
		{	return E_NOTIMPL;}

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
		STDMETHOD(Remove)(VARIANT key);
		STDMETHOD(RemoveAll)(void);

		STDMETHOD(isNewSession)(VARIANT_BOOL* pVal)
		{	return E_NOTIMPL;}
		STDMETHOD(get_Application)(IDispatch** pVal)
		{	return E_NOTIMPL;}
		STDMETHOD(GetSession)(BSTR bstrID)
		{	return E_NOTIMPL;}
		STDMETHOD(Suspend)(void)
		{	return E_NOTIMPL;}
		STDMETHOD(Update)(void)
		{	return E_NOTIMPL;}
		STDMETHOD(Abandon)(void)
		{	return E_NOTIMPL;}
		STDMETHOD(Execute)(BSTR strEvent, VARIANT v1, VARIANT v2, VARIANT v3, VARIANT v4, VARIANT v5, VARIANT v6, VARIANT v7, VARIANT v8, VARIANT* pVal)
		{	return E_NOTIMPL;}
		STDMETHOD(Rnd)(long* pVal)
		{	return E_NOTIMPL;}

		void put_IndexItem(int i, VARIANT& var);
		void put_Item(int i, VARIANT& var);

	public:
		HRESULT NewSession(IRequest* spRequest, IResponse* spResponse);
		HRESULT UpdateDate(IRequest* spRequest, IResponse* spResponse);
		void Reserve();
		void Suspend(BOOL bEvent = TRUE);
		void Offline(BOOL bEvent = TRUE);
		void UpdateLocal();

		void SetSessionID(__int64 id);
		BOOL isLocal();

		HRESULT DoEvent(LPCWSTR strEvent, IRequest* spRequest, IResponse* spResponse);
		void DoEvent(LPCWSTR strEvent);

		HRESULT readBuffer(CAtlArray<BYTE>& buf);
		HRESULT syncBuffer(BYTE* buf, int size);

		__int64 m_date, m_dateSync, m_dateModify;
		BOOL m_bIsLocal;

	private:
		POSITION m_posLocalOnline, m_posSyceOnline, m_posReserveSession, m_posMap;

		CAtlArray<POSITION> m_arrayPos;
		POSITION m_posPY1, m_posPY2;
		char m_chPY1, m_chPY2;
		CXComPtr<CXDictionary> m_pContents;
		__int64 m_nSessionID;
		long m_nAccessTimes;

	private:
		CAtlArray<VARIANT> m_arrayVariant;
		CXComPtr<CXFields> m_pFields;
		CAtlArray<VARIANT> m_arrayIndexVariant;
	};

public:
	class CXSessionList : public CAtlList< CXComPtr<CXSessionInst> >
	{
	public:
		CXSessionList( UINT nBlockSize = 1024 ) throw() : CAtlList< CXComPtr<CXSessionInst> >( nBlockSize )
		{}
	};

	class CXIndex : public CRBMultiMap<CXVariant, CXComPtr<CXSessionInst> >
	{
	public:
		CXIndex( UINT nBlockSize = 1024 ) throw() : CRBMultiMap<CXVariant, CXComPtr<CXSessionInst> >( nBlockSize )
		{}
	};

	class CXMapSession : public CRBMap<__int64, CXComPtr<CXSessionMan::CXSessionInst> >
	{
	public:
		CXMapSession( UINT nBlockSize = 1024 ) throw() : CRBMap<__int64, CXComPtr<CXSessionMan::CXSessionInst> >( nBlockSize )
		{}
	};

public:
	STDMETHOD(AddField)(BSTR key, SHORT type, VARIANT varIndex);
	STDMETHOD(StartBroadcast)(BSTR strNetwork, BSTR strCastIP, int port);
	STDMETHOD(ListUser)(BSTR leadChar, IXList** pVal);
	STDMETHOD(CountUser)(BSTR leadChar, long* pVal);
	STDMETHOD(ListWhere)(BSTR key, BSTR op, VARIANT v1, VARIANT v2, IXList** pVal);
	STDMETHOD(CountWhere)(BSTR key, BSTR op, VARIANT v1, VARIANT v2, long* pVal);
	STDMETHOD(ExistsWhere)(BSTR key, BSTR op, VARIANT v1, VARIANT v2, VARIANT_BOOL* pVal);
	STDMETHOD(UpdateWhere)(BSTR key, VARIANT newVal, BSTR whereKey, BSTR op, VARIANT v1, VARIANT v2);
	STDMETHOD(FireEvent)(BSTR strEvent, VARIANT v1, VARIANT v2, VARIANT v3, VARIANT v4, VARIANT v5, VARIANT v6, VARIANT v7, VARIANT v8);
	STDMETHOD(SetTimer)(LONG Val);
	STDMETHOD(ClearSession)(void);
	STDMETHOD(CountOnline)(LONG* pVal);
	STDMETHOD(CountLocal)(LONG* pVal);
	STDMETHOD(CountSession)(LONG* pVal);
	STDMETHOD(get_OnlineTimeout)(LONG* pVal);
	STDMETHOD(put_OnlineTimeout)(LONG newVal);
	STDMETHOD(get_SuspendTimeout)(LONG* pVal);
	STDMETHOD(put_SuspendTimeout)(LONG newVal);
	STDMETHOD(get_SessionTimeout)(LONG* pVal);
	STDMETHOD(put_SessionTimeout)(LONG newVal);
	STDMETHOD(get_SessionDomain)(BSTR* pVal);
	STDMETHOD(put_SessionDomain)(BSTR newVal);
	STDMETHOD(get_SessionKey)(BSTR* pVal);
	STDMETHOD(put_SessionKey)(BSTR newVal);
	STDMETHOD(get_ComputerName)(BSTR* pVal);
	STDMETHOD(get_Version)(BSTR* pVal);
	STDMETHOD(get_Debug)(int* pVal);
	STDMETHOD(Execute)(BSTR strEvent, VARIANT v1, VARIANT v2, VARIANT v3, VARIANT v4, VARIANT v5, VARIANT v6, VARIANT v7, VARIANT v8, VARIANT* pVal);

private:
	static DWORD WINAPI staticBackLoopProc(LPVOID lpThreadParameter);
	static void WINAPI staticBackClearOnline();
	static DWORD WINAPI staticBackClearProc(LPVOID lpThreadParameter);
	static DWORD WINAPI staticBackSyncProc(LPVOID lpThreadParameter);

private:
	HANDLE m_hBackLoop, m_hBackClear, m_hBackSync;
};

OBJECT_ENTRY_AUTO(__uuidof(XSessionMan), CXSessionMan)
