// XDictionary.cpp : Implementation of CXDictionary

#include "stdafx.h"
#include "XDictionary.h"


// CXDictionary

STDMETHODIMP CXDictionary::get_Item(VARIANT key, VARIANT* pVal)
{
	VARIANT* pkey = &key;
	BOOL bFound;

	while(pkey && pkey->vt == (VT_VARIANT | VT_BYREF))
		pkey = pkey->pvarVal;

	if(!pkey || (pkey->vt & VT_ARRAY))return E_INVALIDARG;
	if(pkey->vt == VT_ERROR)return DISP_E_PARAMNOTOPTIONAL;

	m_cs.Enter();
	bFound = m_mapItems.Lookup(*(CXVariant*)pkey, *(CComVariant*)pVal);
	m_cs.Leave();

	return bFound ? S_OK : S_FALSE;
}

HRESULT CXDictionary::putItem(VARIANT* pvarKey, VARIANT* pvar)
{
	VARIANT varTempKey = {VT_EMPTY};

	while(pvarKey && pvarKey->vt == (VT_VARIANT | VT_BYREF))
		pvarKey = pvarKey->pvarVal;

	if(!pvarKey || (pvarKey->vt & VT_ARRAY) || !pvar)
		return E_INVALIDARG;

	if(pvarKey->vt == VT_ERROR)return DISP_E_PARAMNOTOPTIONAL;

	if(pvarKey->vt == VT_UNKNOWN || pvarKey->vt == VT_DISPATCH)
	{
		HRESULT hr;
		CComDispatchDriver pDisp;

		if(pvarKey->punkVal == NULL)
			return DISP_E_UNKNOWNINTERFACE;

		hr = pvarKey->punkVal->QueryInterface(&pDisp);
		if(FAILED(hr))return hr;

		hr = pDisp.GetProperty(0, &varTempKey);
		if(FAILED(hr))return hr;

		pvar = &varTempKey;
	}

	if(pvarKey->vt == VT_UNKNOWN || pvarKey->vt == VT_DISPATCH)
	{
		VariantClear(&varTempKey);
		return DISP_E_TYPEMISMATCH;
	}

	m_cs.Enter();

	CRBMap<CXVariant, CComVariant>::CPair* pPair = m_mapItems.Lookup(*(CXVariant*)pvarKey);

	if(pPair == NULL)
	{
		static VARIANT varTemp;

		pPair = (CRBMap<CXVariant, CComVariant>::CPair*)m_mapItems.SetAt(*(CXVariant*)pvarKey, varTemp);
	}

	VariantCopyInd(&pPair->m_value, pvar);

	m_cs.Leave();

	if(varTempKey.vt != VT_EMPTY)VariantClear(&varTempKey);;

	return S_OK;
}

STDMETHODIMP CXDictionary::put_Item(VARIANT key, VARIANT newVal)
{
	VARIANT *pvar = &newVal;
	CComVariant varTemp;

	while(pvar && pvar->vt == (VT_VARIANT | VT_BYREF))
		pvar = pvar->pvarVal;

	if(pvar && (pvar->vt == VT_UNKNOWN || pvar->vt == VT_DISPATCH))
	{
		HRESULT hr;
		CComDispatchDriver pDisp;

		if(pvar->punkVal == NULL)
			return DISP_E_UNKNOWNINTERFACE;

		hr = pvar->punkVal->QueryInterface(&pDisp);
		if(FAILED(hr))return hr;

		varTemp.Clear();
		hr = pDisp.GetProperty(0, &varTemp);
		if(FAILED(hr))return hr;

		pvar = &varTemp;
	}

	return putItem(&key, pvar);
}

STDMETHODIMP CXDictionary::putref_Item(VARIANT key, VARIANT newVal)
{
	VARIANT *pvar = &newVal;

	while(pvar && pvar->vt == (VT_VARIANT | VT_BYREF))
		pvar = pvar->pvarVal;

	if(pvar && (pvar->vt == VT_UNKNOWN || pvar->vt == VT_DISPATCH) && pvar->punkVal)
	{
		CComQIPtr<IMarshal> pMarshal;
		CLSID clsid = {0};

		if(pMarshal = pvar->punkVal)
			pMarshal->GetUnmarshalClass(IID_IUnknown, &pvar->punkVal, MSHCTX_INPROC, NULL, MSHLFLAGS_NORMAL, &clsid);

		if(clsid != CLSID_FreeThreadedMarshaler)
			return E_INVALIDARG;
	}

	return putItem(&key, pvar);
}

STDMETHODIMP CXDictionary::Count(LONG* pVal)
{
	m_cs.Enter();
	*pVal = (long)m_mapItems.GetCount();
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXDictionary::Remove(VARIANT key)
{
	VARIANT* pkey = &key;

	while(pkey && pkey->vt == (VT_VARIANT | VT_BYREF))
		pkey = pkey->pvarVal;

	if(!pkey || (pkey->vt & VT_ARRAY))return E_INVALIDARG;
	if(pkey->vt == VT_ERROR)return DISP_E_PARAMNOTOPTIONAL;

	m_cs.Enter();
	m_mapItems.RemoveKey(*(CXVariant*)pkey);
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXDictionary::RemoveAll(void)
{
	m_cs.Enter();
	m_mapItems.RemoveAll();
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXDictionary::Add(VARIANT key, VARIANT value)
{
	return putItem(&key, &value);
}

STDMETHODIMP CXDictionary::Exists(VARIANT key, VARIANT_BOOL* pVal)
{
	VARIANT* pkey = &key;

	while(pkey && pkey->vt == (VT_VARIANT | VT_BYREF))
		pkey = pkey->pvarVal;

	if(!pkey || (pkey->vt & VT_ARRAY))return E_INVALIDARG;
	if(pkey->vt == VT_ERROR)return DISP_E_PARAMNOTOPTIONAL;

	m_cs.Enter();
	*pVal = m_mapItems.Lookup(*(CXVariant*)pkey) ? VARIANT_TRUE : VARIANT_FALSE;
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXDictionary::Join(VARIANT strKeyDelimiter, VARIANT strDelimiter, BSTR* pVal)
{
	HRESULT hr;
	CXString strValue, strKeyDel, strDel;
	POSITION pos;
	CRBMap<CXVariant, CComVariant>::CPair* pPair;

	varGetString(strKeyDelimiter, strKeyDel, L":");
	varGetString(strDelimiter, strDel,L",");

	m_cs.Enter();

	pos = m_mapItems.GetHeadPosition();

	while(pos)
	{
		VARIANT var = {VT_EMPTY};

		pPair = (CRBMap<CXVariant, CComVariant>::CPair*)m_mapItems.GetNext(pos);

		if(strValue[0])strValue.Append(strDel);

		if(pPair->m_key.vt == VT_BSTR)
			strValue.Append(pPair->m_key.bstrVal);
		else
		{
			CComVariant var;

			hr = VariantChangeType(&var, (VARIANTARG*)&pPair->m_key, VARIANT_ALPHABOOL, VT_BSTR);
			if(FAILED(hr))
			{
				m_cs.Leave();
				return hr;
			}

			strValue.Append(var.bstrVal);
		}

		strValue.Append(strKeyDel);

		if(pPair->m_value.vt == VT_BSTR)
			strValue.Append(pPair->m_value.bstrVal);
		else
		{
			CComVariant var;

			hr = VariantChangeType(&var, &pPair->m_value, VARIANT_ALPHABOOL, VT_BSTR);
			if(FAILED(hr))
			{
				m_cs.Leave();
				return hr;
			}

			strValue.Append(var.bstrVal);
		}
	}

	m_cs.Leave();

	*pVal = strValue.AllocSysString();

	return S_OK;
}

inline BSTR findstr(BSTR bstr, int nCount, LPCWSTR strKey)
{
	BSTR bstrTemp;
	UINT nCountTemp;
	LPCWSTR strKeyTemp;

	while(nCount > 0)
	{
		bstrTemp = bstr;
		nCountTemp = nCount;
		strKeyTemp = strKey;

		while(nCountTemp && *strKeyTemp && *bstrTemp == *strKeyTemp)
		{
			nCountTemp --;
			strKeyTemp ++;
			bstrTemp ++;
		}

		if(!*strKeyTemp)return bstr;

		bstr ++;
		nCount --;
	}

	return bstr;
}

STDMETHODIMP CXDictionary::Split(BSTR bstrExpression, VARIANT strKeyDelimiter, VARIANT strDelimiter)
{
	int nCount, strKeyDelimiterLength, strDelimiterLength;
	BSTR pstr;
	CXVariant varKey;
	VARIANT varValue;
	CRBMap<CXVariant, CComVariant>::CPair* pPair;
	static VARIANT varEmpty;
	CXString strKeyDel, strDel;

	varGetString(strKeyDelimiter, strKeyDel, L":");
	varGetString(strDelimiter, strDel,L",");

	strKeyDelimiterLength = strKeyDel.GetLength();
	strDelimiterLength = strDel.GetLength();

	varKey.vt = VT_BSTR;
	VariantInit(&varValue);
	varValue.vt = VT_BSTR;

	nCount = ::SysStringLen(bstrExpression);

	m_cs.Enter();

	while(nCount > 0)
	{
		pstr = findstr(bstrExpression, nCount, strKeyDel);
		varKey.bstrVal = ::SysAllocStringLen(bstrExpression, pstr - bstrExpression);
		nCount -= strKeyDelimiterLength + pstr - bstrExpression;
		bstrExpression = pstr + strKeyDelimiterLength;

		pstr = findstr(bstrExpression, nCount, strDel);
		varValue.bstrVal = ::SysAllocStringLen(bstrExpression, pstr - bstrExpression);
		nCount -= strDelimiterLength + pstr - bstrExpression;
		bstrExpression = pstr + strDelimiterLength;

		pPair = m_mapItems.Lookup(varKey);

		if(pPair == NULL)
			pPair = (CRBMap<CXVariant, CComVariant>::CPair*)m_mapItems.SetAt(varKey, varEmpty);
		else VariantClear(&pPair->m_value);

		SysFreeString(varKey.bstrVal);

		CopyMemory(&pPair->m_value, &varValue, sizeof(VARIANT));
	}

	m_cs.Leave();

	varKey.vt = VT_EMPTY;

	return S_OK;
}

STDMETHODIMP CXDictionary::get_Keys(VARIANT* pVal)
{
	CComSafeArray<VARIANT> bstrArray;
	VARIANT* pVar;
	HRESULT hr;
	POSITION pos;
	CRBMap<CXVariant, CComVariant>::CPair* pPair;
	int i = 0;

	m_cs.Enter();

	hr = bstrArray.Create((ULONG)m_mapItems.GetCount());
	if(FAILED(hr))
	{
		m_cs.Leave();
		return hr;
	}

	pVar = (VARIANT*)bstrArray.m_psa->pvData;
	pos = m_mapItems.GetHeadPosition();

	while(pos)
	{
		VARIANT var = {VT_EMPTY};

		pPair = (CRBMap<CXVariant, CComVariant>::CPair*)m_mapItems.GetNext(pos);
		hr = VariantCopy(&pVar[i ++], (VARIANT*)&pPair->m_key);
		if(FAILED(hr))
		{
			m_cs.Leave();
			bstrArray.Destroy();
			return hr;
		}
	}

	m_cs.Leave();

	pVal->vt = VT_ARRAY | VT_VARIANT;
	pVal->parray = bstrArray.Detach();

	return S_OK;
}

STDMETHODIMP CXDictionary::get_Items(VARIANT* pVal)
{
	CComSafeArray<VARIANT> bstrArray;
	VARIANT* pVar;
	HRESULT hr;
	POSITION pos;
	CRBMap<CXVariant, CComVariant>::CPair* pPair;
	int i = 0;

	m_cs.Enter();

	hr = bstrArray.Create((ULONG)m_mapItems.GetCount());
	if(FAILED(hr))
	{
		m_cs.Leave();
		return hr;
	}

	pVar = (VARIANT*)bstrArray.m_psa->pvData;
	pos = m_mapItems.GetHeadPosition();

	while(pos)
	{
		VARIANT var = {VT_EMPTY};

		pPair = (CRBMap<CXVariant, CComVariant>::CPair*)m_mapItems.GetNext(pos);
		hr = VariantCopy(&pVar[i ++], (VARIANT*)&pPair->m_value);
		if(FAILED(hr))
		{
			m_cs.Leave();
			bstrArray.Destroy();
			return hr;
		}
	}

	m_cs.Leave();

	pVal->vt = VT_ARRAY | VT_VARIANT;
	pVal->parray = bstrArray.Detach();

	return S_OK;
}

STDMETHODIMP CXDictionary::get__NewEnum(IUnknown** pVal)
{
	return getNewEnum(this, pVal);
}

void CXDictionary::FillEnum(CAtlArray<VARIANT>& arrayVariant)
{
	POSITION pos;
	CRBMap<CXVariant, CComVariant>::CPair* pPair;

	m_cs.Enter();

	pos = m_mapItems.GetHeadPosition();

	while(pos)
	{
		VARIANT var = {VT_EMPTY};

		pPair = (CRBMap<CXVariant, CComVariant>::CPair*)m_mapItems.GetNext(pos);
		VariantCopy(&var, (VARIANT*)&pPair->m_key);
		arrayVariant.Add(var);
	}

	m_cs.Leave();
}

