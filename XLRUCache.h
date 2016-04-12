// XLRUCache.h : Declaration of the CXLRUCache

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"
#include "atlib.h"


// CXLRUCache

class ATL_NO_VTABLE CXLRUCache : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXLRUCache, &CLSID_XLRUCache>,
	public IDispatchImpl<IXLRUCache, &IID_IXLRUCache, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CXLRUCache()
	{
		m_pUnkMarshaler = NULL;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_XLRUCACHE)


BEGIN_COM_MAP(CXLRUCache)
	COM_INTERFACE_ENTRY(IXLRUCache)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
END_COM_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()
	DECLARE_GET_CONTROLLING_UNKNOWN()

	HRESULT FinalConstruct()
	{
		m_nSize = 1024;
		m_nTimeout = -1;
		getCount = 1;
		hitCount = 1;

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
	STDMETHOD(GetItem)(VARIANT key, VARIANT* pVal);
	STDMETHOD(Pop)(VARIANT key, VARIANT* pVal);
	STDMETHOD(Count)(LONG* pVal);
	STDMETHOD(Remove)(VARIANT key);
	STDMETHOD(RemoveAll)(void);
	STDMETHOD(Add)(VARIANT key, VARIANT value);
	STDMETHOD(Exists)(VARIANT key, VARIANT_BOOL* pVal);
	STDMETHOD(get_MaxSize)(LONG* pVal);
	STDMETHOD(put_MaxSize)(LONG newVal);
	STDMETHOD(get_Timeout)(LONG* pVal);
	STDMETHOD(put_Timeout)(LONG newVal);
	STDMETHOD(get_MaxTime)(LONG* pVal);
	STDMETHOD(get_HitRate)(float* pVal);
	STDMETHOD(get__NewEnum)(IUnknown** pVal);

public:
	void FillEnum(CAtlArray<VARIANT>& arrayVariant);

private:
	void CheckTimeout();
	void removeNode(POSITION pos);

private:
	class LRUNode
	{
	public:
		POSITION posAccess;
		POSITION posTimeout;
		CComVariant var;
		__int64 timeout;
	};

	CCriticalSection m_cs;
	CRBMap<CXVariant, LRUNode> m_mapItems;
	CAtlList<POSITION> m_listAccess;
	CAtlList<POSITION> m_listTimeout;
	int m_nSize;
	int m_nTimeout;

	__int64 getCount, hitCount;

	HRESULT putItem(VARIANT* pvarKey, VARIANT* pvar);
};

OBJECT_ENTRY_AUTO(__uuidof(XLRUCache), CXLRUCache)
