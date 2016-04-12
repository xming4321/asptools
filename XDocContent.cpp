// XDocContent.cpp : Implementation of CXDocContent

#include "stdafx.h"
#include "XDocContent.h"


// CXDocContent


STDMETHODIMP CXDocContent::get_Time(DATE* pVal)
{
	*pVal = m_date;

	return S_OK;
}

STDMETHODIMP CXDocContent::put_Time(DATE newVal)
{
	m_date = newVal;

	return S_OK;
}

STDMETHODIMP CXDocContent::get_Label(BSTR* pVal)
{
	*pVal = m_strLabel.AllocSysString();

	return S_OK;
}

STDMETHODIMP CXDocContent::put_Label(BSTR newVal)
{
	m_strLabel = newVal;

	return S_OK;
}

STDMETHODIMP CXDocContent::get_Text(BSTR* pVal)
{
	*pVal = m_strText.AllocSysString();

	return S_OK;
}

STDMETHODIMP CXDocContent::put_Text(BSTR newVal)
{
	m_strText = newVal;

	return S_OK;
}

