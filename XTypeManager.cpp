// XTypeManager.cpp : Implementation of CXTypeManager

#include "stdafx.h"
#include "XTypeManager.h"

CAtlArray< CXComPtr<CXFields> > s_aFormats;
CAtlArray< __int64 > s_aFormatHashs;
CAtlArray< CXString > s_aFormatName;

// CXTypeManager

CXTypeManager::~CXTypeManager()
{
	s_aFormats.RemoveAll();
	s_aFormatHashs.RemoveAll();
	s_aFormatName.RemoveAll();
}

__int64 CXTypeManager::GetFormatHash(BSTR index)
{
	for(int i = 0; i < (int)s_aFormatName.GetCount(); i ++)
		if(!s_aFormatName[i].CompareNoCase(index))
			return s_aFormatHashs[i];

	return -1;
}

CXFields* CXTypeManager::GetFormat(__int64 hash)
{
	int i, count;

	count = (int)s_aFormats.GetCount();

	for(i = 0; i < count; i ++)
		if(s_aFormatHashs[i] == hash)
			return s_aFormats[i];

	return NULL;
}

BSTR CXTypeManager::GetFormatName(__int64 hash)
{
	int i, count;

	count = (int)s_aFormats.GetCount();

	for(i = 0; i < count; i ++)
		if(s_aFormatHashs[i] == hash)
			return s_aFormatName[i].AllocSysString();

	return ::SysAllocString(L"");
}

STDMETHODIMP CXTypeManager::CXClass::AddField(BSTR key, VARIANT type)
{
	HRESULT hr;

	if(type.vt == VT_BSTR)
	{
		__int64 t = CXTypeManager::GetFormatHash(type.bstrVal);

		if(t == -1)
			return E_INVALIDARG;
		hr = s_aFormats[m_nIndex]->AddField(key, VT_DISPATCH, t);
	}else hr = s_aFormats[m_nIndex]->AddField(key, (short)varGetNumbar(type));
	if(FAILED(hr))return hr;

	s_aFormatHashs[m_nIndex] = s_aFormats[m_nIndex]->GetHash();

    return S_OK;
}

STDMETHODIMP CXTypeManager::CXClass::AddRecordset(BSTR key, VARIANT type)
{
	HRESULT hr;

	if(type.vt == VT_BSTR)
	{
		__int64 t = CXTypeManager::GetFormatHash(type.bstrVal);

		if(t == -1)
			return E_INVALIDARG;
		hr = s_aFormats[m_nIndex]->AddRecordset(key, t);
	}else return E_INVALIDARG;

	if(FAILED(hr))return hr;

	s_aFormatHashs[m_nIndex] = s_aFormats[m_nIndex]->GetHash();

    return S_OK;
}

STDMETHODIMP CXTypeManager::newClass(BSTR Ver, IXClass** pVal)
{
	CXComPtr<CXClass> pClass;
	int i;

	for(i = 0; i < (int)s_aFormatName.GetCount(); i ++)
		if(!s_aFormatName[i].CompareNoCase(Ver))
			break;

	if(i == s_aFormatName.GetCount())
	{	
		CXComPtr<CXFields> pFields;

		s_aFormatName.Add();
		s_aFormatName[i] = Ver;

		pFields.CreateObject();
		s_aFormats.Add(pFields);
		s_aFormatHashs.Add(0);
	}

	pClass.CreateObject();
	pClass->m_nIndex = i;

	*pVal = pClass.Detach();

	return S_OK;
}

STDMETHODIMP CXTypeManager::newRecordset(VARIANT varFmt, IXRecords** pVal)
{
	CXComPtr<CXRecords> pRecords;

	pRecords.CreateObject();

	CXString strFmt;

	varGetString(varFmt, strFmt);

	if(!strFmt.IsEmpty())
	{
		HRESULT hr = pRecords->Create((BSTR)(LPCWSTR)strFmt);
		if(FAILED(hr))return hr;
	}

	return pRecords.QueryInterface(pVal);
}
