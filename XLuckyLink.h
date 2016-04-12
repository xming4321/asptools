// XLuckyLink.h : Declaration of the CXLuckyLink

#pragma once
#include "atlib.h"       // main symbols
#include "resource.h"       // main symbols

#include "asptools.h"

#include <math.h>


// CXLuckyLink

class ATL_NO_VTABLE CXLuckyLink :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXLuckyLink, &CLSID_XLuckyLink>,
	public IDispatchImpl<IXLuckyLink, &IID_IXLuckyLink, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
private:
	class CXTree;
	class CXNode : public CXObject
	{
	public:
		CXNode()
		{
			m_nLife = 0;
			m_nAct = 0;
			m_pParent = NULL;
			m_pos = NULL;
			m_nID = 0;
		}

	public:
		CXString m_strValue;
		double m_nLife;

		int m_nID;

		CXTree* m_pParent;
		POSITION m_pos;

		unsigned __int64 m_nAct;
	};

	class CXTree : public CXObject
	{
	public:
		CXTree()
		{
			m_nLifes = 0;
			m_nSelfLifes = 0;
			m_nLifesAct = 0;
			m_nNodeCount = 0;
			m_pParent = NULL;
			m_nAct = 0;

			m_nCountAct = 0;
			m_nLifesAct = 0;
		}

	public:
		void HalfLife(double life);
		void UpdateLife(int count);
		CXTree* GetNode(LPCWSTR mark, BOOL bCreate = FALSE);
		void RemoveAll()
		{
			m_nLifes = 0;
			m_nSelfLifes = 0;
			m_nLifesAct = 0;
			m_nNodeCount = 0;
			m_pParent = NULL;
			m_nAct = 0;

			m_nCountAct = 0;
			m_nLifesAct = 0;

			m_aSubTrees.RemoveAll();
			m_aNodes.RemoveAll();
		}

	public:
		CXNode* GetItem(__int64 nAct);

	private:
		void inline checkAct(__int64 nAct)
		{
			if(nAct != m_nAct)
			{
				m_nCountAct = m_nNodeCount;
				m_nLifesAct = m_nLifes;
				m_nAct = nAct;
			}
		}

		UINT inline GetCount(__int64 nAct)
		{
			checkAct(nAct);
			return m_nCountAct;
		}

		double inline GetLifes(__int64 nAct)
		{
			checkAct(nAct);
			return m_nLifesAct;
		}

	public:
		CAtlArray< CXComPtr<CXTree> > m_aSubTrees;
		CAtlList< CXComPtr<CXNode> > m_aNodes;

		CXTree* m_pParent;

		double m_nLifes;
		double m_nSelfLifes;

		UINT m_nNodeCount;

		UINT m_nCountAct;
		double m_nLifesAct;
		unsigned __int64 m_nAct;
	};

public:
	CXLuckyLink()
	{
		m_baseTime = GetVariantTime();
		m_nhalfLifeTimes = 24;
		m_nAct = 0;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_XLUCKYLINK)


BEGIN_COM_MAP(CXLuckyLink)
	COM_INTERFACE_ENTRY(IXLuckyLink)
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
	STDMETHOD(get_halfLife)(int* pVal);
	STDMETHOD(put_halfLife)(int newVal);
	STDMETHOD(get_Count)(int* pVal);
	STDMETHOD(AddItem)(BSTR mark, double life, DATE time, int id, BSTR value);
	STDMETHOD(BatchAdd)(BSTR batchText);
	STDMETHOD(RemoveItem)(int id);
	STDMETHOD(RemoveAll)();
	STDMETHOD(GetList)(BSTR mark, int nCount, VARIANT seed, VARIANT* pVal);

private:
	HRESULT AddOneItem(LPCWSTR mark, double life, DATE time, int id, LPCWSTR value);

	int m_nhalfLifeTimes;
	CXTree m_treeRoot;
	DATE m_baseTime;

	CRBMap< int, CXComPtr<CXNode> > m_mapNodes;

	unsigned __int64 m_nAct;

	CCriticalSection m_cs;
};

OBJECT_ENTRY_AUTO(__uuidof(XLuckyLink), CXLuckyLink)
