// XGeoLinks.h : Declaration of the CXGeoLinks

#pragma once
#include "resource.h"       // main symbols

#include "atlib.h"       // main symbols
#include "asptools.h"
#include "XGeoTools.h"


#if defined(_WIN32_WCE) && !defined(_CE_DCOM) && !defined(_CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA)
#error "Single-threaded COM objects are not properly supported on Windows CE platform, such as the Windows Mobile platforms that do not include full DCOM support. Define _CE_ALLOW_SINGLE_THREADED_OBJECTS_IN_MTA to force ATL to support creating single-thread COM object's and allow use of it's single-threaded COM object implementations. The threading model in your rgs file was set to 'Free' as that is the only threading model supported in non DCOM Windows CE platforms."
#endif



// CXGeoLinks

class ATL_NO_VTABLE CXGeoLinks :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXGeoLinks, &CLSID_XGeoLinks>,
	public IDispatchImpl<IXGeoLinks, &IID_IXGeoLinks, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CXGeoLinks()
	{
		m_pUnkMarshaler = NULL;
		m_root.CreateObject();

		m_nhalfLifeTimes = 10;
		m_nhalfLifeRange = 1;
		m_baseTime = GetVariantTime();
	}

DECLARE_REGISTRY_RESOURCEID(IDR_XGEOLINKS)


BEGIN_COM_MAP(CXGeoLinks)
	COM_INTERFACE_ENTRY(IXGeoLinks)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
END_COM_MAP()

	DECLARE_PROTECT_FINAL_CONSTRUCT()
	DECLARE_GET_CONTROLLING_UNKNOWN()

	HRESULT FinalConstruct()
	{
		if(!m_root)
			return E_OUTOFMEMORY;

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
	STDMETHOD(get_halfRange)(int* pVal);
	STDMETHOD(put_halfRange)(int newVal);
	STDMETHOD(get_Count)(int* pVal);
	STDMETHOD(Modify)(IXGeoLinks** pVal);
	STDMETHOD(Commit)(void);
	STDMETHOD(AddItem)(int id, int cityid, double Latitude, double Longitude, int type, BSTR mark, BSTR tag, double life, DATE time, BSTR value);
	STDMETHOD(Remove)(int id);
	STDMETHOD(RemoveAll)();
	STDMETHOD(BatchAdd)(BSTR batchText);
	STDMETHOD(GetList)(int cityid, double Latitude, double Longitude, BSTR mark, BSTR tag, BSTR key, int nStart, int nCount, BSTR Types, int mode, VARIANT seed, VARIANT* pVal);
	STDMETHOD(GetCityList)(int cityid, BSTR mark, BSTR tag, int nStart, int nCount, BSTR Types, VARIANT* pVal);
	STDMETHOD(GetTagList)(int cityid, BSTR mark, int nStart, int nCount, BSTR Types, VARIANT* pVal);
	STDMETHOD(GetHotTag)(int cityid, BSTR mark, BSTR tag, int nStart, int nCount, VARIANT* pVal);

private:
	class CXTree;
	class CXNode : public CXObject
	{
	public:
		CXNode()
		{
			m_nID = 0;
			m_nCityID = 0;
			m_nlat = 0;
			m_nlng = 0;
			m_strMark[0] = 0;
			m_nLife = 0;
		}

	public:
		CXString m_strValue;
		WCHAR m_strMark[16];
		double m_nLife;
		int m_nID, m_nCityID;
		CAtlArray<CXString> m_strTags;
		double m_nlat, m_nlng;
	};

	class CXRoot;

	class CXNodeIndex
	{
	public:
		CXNode* m_pNode;
		int m_cityid;
		double m_nLife;
	};

	class CXTagIndex
	{
	public:
		CXString m_strTag;
		double m_nLife;
	};

	class CXTree
	{
	public:
		CXTree()
		{
			ZeroMemory(&m_aSubTrees, sizeof(m_aSubTrees));
		}

		~CXTree();

	protected:
		void CXGeoLinks::CXTree::GetTreeNodes(LPCWSTR tag, CAtlArray<CXString>* keys, int rate, int mode, CAtlArray<CXNodeIndex>& nodes, int nCount);

	public:
		void AddAreaItem(CXNode* pNode, LPCWSTR mark);
		void GetNodes(LPCWSTR mark, LPCWSTR tag, CAtlArray<CXString>* keys, int mode, int *rates, CAtlArray<CXNodeIndex>& nodes, int nCount);
		void FillNodes(CAtlList< CXComPtr<CXNode> >* pList, CAtlArray<CXString>* keys, int mode, int rate, CAtlArray<CXNodeIndex>& nodes, int nCount);

		void FillTagNodes(CAtlArray<CXNodeIndex>& basenodes, CAtlArray<CXNodeIndex>& nodes, int rate, int nCount);
		void GetTagNodes(LPCWSTR mark, CAtlArray<CXNodeIndex>& basenodes, CAtlArray<CXNodeIndex>& nodes, int nCount, int *rates);
		void GetHotTags(LPCWSTR mark, LPCWSTR tag, CRBMap<CXString, int>& nodes, double& n);
		void CountHotTags(LPCWSTR tag, CRBMap<CXString, int>& nodes, double& n);
		void FillHotTags(CAtlList< CXComPtr<CXNode> >* pList, LPCWSTR tag, CRBMap<CXString, int>& nodes, double& n);

	public:
		CXTree* m_aSubTrees[26];
		CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* > m_aTagNodes;
		CRBMap< CXString, CAtlList< CXComPtr<CXNode> >* > m_aTagNodes1;
	};

	class CXRoot : public CXObject
	{
	public:
		CXRoot()
		{
			ZeroMemory(m_mapCitys, sizeof(m_mapCitys));
		}

		~CXRoot()
		{
			int i;

			for(i = 0; i < CITYCOUNT; i ++)
				if(m_mapCitys[i])delete m_mapCitys[i];
		}

	public:
		void GetNodes(int cityid, double Latitude, double Longitude, LPCWSTR mark, LPCWSTR tag, CAtlArray<CXString>* keys, int mode, int nCount, int *rates, CAtlArray<CXNodeIndex>& nodes);
		void GetTagNodes(int cityid, LPCWSTR mark, int nCount, int *rates, CAtlArray<CXNodeIndex>& nodes);
		void GetCityNodes(int cityid, LPCWSTR mark, LPCWSTR tag, int nCount, int *rates, CAtlArray<CXNodeIndex>& nodes);
		void GetHotTags(int cityid, LPCWSTR mark, LPCWSTR tag, int nCount, CAtlArray<CXTagIndex>& nodes);
		void AddAreaItem(CXNode* pNode);

	private:
		CXTree* m_mapCitys[CITYCOUNT];
	};

private:
	void AddOneItem(int id, int cityid, double Latitude, double Longitude, LPCWSTR mark, LPCWSTR tag, double life, DATE time, LPCWSTR value);

private:
	CXComPtr<CXRoot> m_root;
	CXComPtr<CXGeoLinks> m_pUpdate;

	CCriticalSection m_cs;
	CRBMap< int, CXComPtr<CXNode> > m_mapNodes;

	int m_nhalfLifeTimes;
	int m_nhalfLifeRange;
	DATE m_baseTime;
};

OBJECT_ENTRY_AUTO(__uuidof(XGeoLinks), CXGeoLinks)
