#pragma once

#include "resource.h"       // main symbols
#include "asptools.h"
#include "atlib.h"

class CXUploadData : public CXDispatchFree<IXUploadData>
{
public:
	CXUploadData(void) : m_nSize(0)
	{}

public:
	// IUploadData
    STDMETHOD(get_Item)(VARIANT *pVariantReturn);
	STDMETHOD(get_FileName)(BSTR *pStrReturn);
	STDMETHOD(get_Size)(long* pVal);
	STDMETHOD(get_ContentType)(BSTR *pStrReturn);

public:
	CComVariant m_varData;
	CXString m_strFileName;
	CXString m_strContentType;
	long m_nSize;
};

class CXUploadList : public CXDispatchFree<IXUploadList>
{
public:
	// IUploadData
    STDMETHOD(get_Item)(VARIANT i, VARIANT *pVariantReturn);
	STDMETHOD(get_Count)(long *cStrRet);
	STDMETHOD(get__NewEnum)(IUnknown **ppEnumReturn);
	STDMETHOD(get_FileName)(BSTR *pStrReturn);
	STDMETHOD(get_Size)(long* pVal);
	STDMETHOD(get_ContentType)(BSTR *pStrReturn);

public:
	void FillEnum(CAtlArray<VARIANT>& arrayVariant)
	{
		int i;

		for (i = 0; i < (int)m_arrayData.GetCount(); i ++)
		{
			VARIANT var = {VT_EMPTY};

			*(CComVariant*)&var = m_arrayData[i];
			arrayVariant.Add(var);
		}
	}

	void AddValue(const CXComPtr<CXUploadData>& pData)
	{
		m_arrayData.Add(pData);
	}

private:
	CAtlArray< CXComPtr<CXUploadData> > m_arrayData;
};
