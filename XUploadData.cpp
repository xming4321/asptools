#include "StdAfx.h"
#include "XUploadData.h"


STDMETHODIMP CXUploadData::get_Item(VARIANT *pVariantReturn)
{
	return VariantCopy(pVariantReturn, &m_varData);
}

STDMETHODIMP CXUploadData::get_FileName(BSTR *pStrReturn)
{
	*pStrReturn = m_strFileName.AllocSysString();
	return S_OK;
}

STDMETHODIMP CXUploadData::get_Size(long* pVal)
{
	*pVal = m_nSize;
	return S_OK;
}

STDMETHODIMP CXUploadData::get_ContentType(BSTR *pStrReturn)
{
	*pStrReturn = m_strContentType.AllocSysString();
	return S_OK;
}

static CXUploadData s_EmptyData;

STDMETHODIMP CXUploadList::get_Item(VARIANT i, VARIANT *pVariantReturn)
{
	if(i.vt == VT_ERROR)
	{
		if(m_arrayData.GetCount() == 1)
			return m_arrayData[0]->get_Item(pVariantReturn);
		else if(m_arrayData.GetCount() > 1)
			return DISP_E_PARAMNOTOPTIONAL;

		return S_OK;
	}

	long n = varGetNumbar(i);

	if(n > 0 && n <= (int)m_arrayData.GetCount())
		*(CComVariant*)pVariantReturn = m_arrayData[n - 1];
	else return DISP_E_BADINDEX;

	return S_OK;
}

STDMETHODIMP CXUploadList::get_Count(long *cStrRet)
{
	*cStrRet = (long)m_arrayData.GetCount();

	return S_OK;
}

STDMETHODIMP CXUploadList::get__NewEnum(IUnknown **ppEnumReturn)
{
	return getNewEnum(this, ppEnumReturn);
}

STDMETHODIMP CXUploadList::get_FileName(BSTR *pStrReturn)
{
	if(m_arrayData.GetCount() == 1)
		return m_arrayData[0]->get_FileName(pStrReturn);
	else if(m_arrayData.GetCount() == 0)
		return S_OK;

	return E_NOTIMPL;
}

STDMETHODIMP CXUploadList::get_Size(long* pVal)
{
	if(m_arrayData.GetCount() == 1)
		return m_arrayData[0]->get_Size(pVal);
	else if(m_arrayData.GetCount() == 0)
	{
		return S_OK;
		*pVal = 0;
	}

	return E_NOTIMPL;
}

STDMETHODIMP CXUploadList::get_ContentType(BSTR *pStrReturn)
{
	if(m_arrayData.GetCount() == 1)
		return m_arrayData[0]->get_ContentType(pStrReturn);
	else if(m_arrayData.GetCount() == 0)
		return S_OK;

	return E_NOTIMPL;
}

