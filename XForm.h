// XForm.h : Declaration of the CXForm

#pragma once

#include "resource.h"       // main symbols
#include <asptlb.h>         // Active Server Pages Definitions

#include "asptools.h"
#include "XUploadData.h"

// CXForm

class ATL_NO_VTABLE CXForm : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXForm, &CLSID_XForm>,
	public IDispatchImpl<IXForm, &IID_IXForm, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
DECLARE_REGISTRY_RESOURCEID(IDR_XFORM)


BEGIN_COM_MAP(CXForm)
	COM_INTERFACE_ENTRY(IXForm)
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

// IXForm
public:
	//Active Server Pages Methods
	STDMETHOD(OnStartPage)(IUnknown* IUnk);
	STDMETHOD(OnEndPage)();
	STDMETHOD(get_Item)(BSTR key, IXUploadList** pVariantReturn);
	STDMETHOD(get__NewEnum)(IUnknown** ppEnumReturn);
	STDMETHOD(get_Count)(long* cStrRet);
	STDMETHOD(Exists)(BSTR key, VARIANT_BOOL* pExists);

public:
	void FillEnum(CAtlArray<VARIANT>& arrayVariant);

private:
	HRESULT ParseUrlEncodeString(LPCSTR pstr, UINT nSize);
	HRESULT ParseUploadString(LPCSTR pstr, UINT nSize);

private:
	void AddValue(CXString strKey, const CXComPtr<CXUploadData>& pData)
	{
		CXComPtr<CXUploadList> pList;

		if(!m_mapForm.Lookup(strKey, pList))
		{
			pList.CreateObject();
			m_mapForm.SetAt(strKey, pList);
		}

		pList->AddValue(pData);
	}

private:
	CRBMap<CXKeyString, CXComPtr<CXUploadList> > m_mapForm;
};

OBJECT_ENTRY_AUTO(__uuidof(XForm), CXForm)
