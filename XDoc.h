// XDoc.h : Declaration of the CXDoc

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"
#include "XList.h"
#include "XFolderMan.h"
#include "XRecords.h"
#include "XFileStream.h"
#include "XDocItem.h"
#include "XTypeManager.h"

#define R(v)		{hr = StreamStub.ReadStream(v);if(FAILED(hr))return hr;}
#define W(v)		{hr = StreamStub.WriteStream(v);if(FAILED(hr))return hr;}

// CXDoc

class ATL_NO_VTABLE CXDoc : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXDoc, &CLSID_XDoc>,
	public IDispatchImpl<IXDoc, &IID_IXDoc, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:

DECLARE_REGISTRY_RESOURCEID(IDR_XDOC)


BEGIN_COM_MAP(CXDoc)
	COM_INTERFACE_ENTRY(IXDoc)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		m_nDocFormat = -1;

		m_listDocs.CreateObject();

		m_bExpand = FALSE;
		m_bCompressed = TRUE;
		m_bTransMode = TRUE;

		return S_OK;
	}

	void FinalRelease() 
	{
	}

public: 
	STDMETHOD(get_Item)(BSTR key, VARIANT* pVal);
	STDMETHOD(put_Item)(BSTR key, VARIANT newVal);
	STDMETHOD(Count)(LONG* pVal);
	STDMETHOD(get__NewEnum)(IUnknown** pVal);
	STDMETHOD(Create)(BSTR ver);
	STDMETHOD(Load)(BSTR strPath);
	STDMETHOD(Open)(BSTR strPath, VARIANT varNew);
	STDMETHOD(Save)(void);
	STDMETHOD(Close)(void);
	STDMETHOD(ConvertFormat)(BSTR newFmt);
	STDMETHOD(get_FormatName)(BSTR* pVal);
	STDMETHOD(get_Compressed)(VARIANT_BOOL* pVal);
	STDMETHOD(put_Compressed)(VARIANT_BOOL newVal);
	STDMETHOD(get_TransMode)(VARIANT_BOOL* pVal);
	STDMETHOD(put_TransMode)(VARIANT_BOOL newVal);
	STDMETHOD(Remove)(BSTR key);
	STDMETHOD(Exists)(BSTR key, VARIANT_BOOL* pVal);
	STDMETHOD(get_docs)(IXList** pVal);
	STDMETHOD(Append)(BSTR user, LONG id, SHORT emote, SHORT rate, SHORT hot, BSTR host, BSTR ip, IXDocItem** pVal);
	STDMETHOD(RemoveDoc)(long i);
	STDMETHOD(AddRecordset)(BSTR key, VARIANT varFmt, IXRecords** pVal);
	STDMETHOD(CountHost)(long* pval);
	STDMETHOD(CountIP)(long* pval);
	STDMETHOD(CountUser)(long* pval);
	STDMETHOD(CountRate)(long* pval);
	STDMETHOD(updateTime)(long nStart, long nEnd, DATE* pval);
	STDMETHOD(HotRank)(long timeline, long timeline1, double* pval);
	STDMETHOD(Expand)(void);

public:
	void FillEnum(CAtlArray<VARIANT>& arrayVariant);
	HRESULT SetFormat(__int64 nDocFormat);

private:
	CXComPtr< CXList< CXComPtr<CXDocItem> > > m_listDocs;
	BOOL m_bCompressed;
	BOOL m_bTransMode;
	BOOL m_bExpand;

	__int64 m_nDocFormat;
	CXComPtr<CXRecord> m_pBaseItems;
	CRBMap<CXString, CComVariant> m_mapItems;

	UINT m_nEndOfFile;

	CXComPtr<CXFileStream> m_pBinFile;
	CXComPtr<CXMemStream> m_pBufFile;

private:
	HRESULT OpenFile(BSTR strFile, int nFlags);
	HRESULT ReadDoc(BOOL bClose = FALSE);
	HRESULT WriteDoc();
};

OBJECT_ENTRY_AUTO(__uuidof(XDoc), CXDoc)

