// XLRUCache.cpp : Implementation of CXLRUCache

#include "stdafx.h"
#include "XLRUCache.h"


// CXLRUCache

void CXLRUCache::removeNode(POSITION pos)
{
	CRBMap<CXVariant, LRUNode>::CPair* pPair = (CRBMap<CXVariant, LRUNode>::CPair*)pos;

	m_listAccess.RemoveAt(pPair->m_value.posAccess);
	m_listTimeout.RemoveAt(pPair->m_value.posTimeout);
	m_mapItems.RemoveAt(pPair);
}

void CXLRUCache::CheckTimeout()
{
	if(m_nTimeout != -1)
	{
		CRBMap<CXVariant, LRUNode>::CPair* pPair;

		while(m_listTimeout.GetCount() > 0)
		{
			pPair = (CRBMap<CXVariant, LRUNode>::CPair*)m_listTimeout.GetHead();
			if(pPair->m_value.timeout + (m_nTimeout * 1000) > GetLocalFileTime())
				break;

			removeNode(pPair);
		}
	}

	while((long)m_listAccess.GetCount() > m_nSize)
		removeNode(m_listAccess.GetHead());
}

STDMETHODIMP CXLRUCache::get_Item(VARIANT key, VARIANT* pVal)
{
	VARIANT* pkey = &key;
	BOOL bFound = FALSE;

	while(pkey && pkey->vt == (VT_VARIANT | VT_BYREF))
		pkey = pkey->pvarVal;

	if(!pkey || (pkey->vt & VT_ARRAY))return E_INVALIDARG;
	if(pkey->vt == VT_ERROR)return DISP_E_PARAMNOTOPTIONAL;

	m_cs.Enter();
	CheckTimeout();

	CRBMap<CXVariant, LRUNode>::CPair* pPair;

	if(pPair = m_mapItems.Lookup(*(CXVariant*)pkey))
	{
		bFound = TRUE;

		m_listAccess.RemoveAt(pPair->m_value.posAccess);

		VariantCopyInd(pVal, &pPair->m_value.var);

		pPair->m_value.posAccess = m_listAccess.AddTail(pPair);
	}

	m_cs.Leave();

	getCount ++;
	if(bFound)
		hitCount ++;

	return bFound ? S_OK : S_FALSE;
}

HRESULT CXLRUCache::putItem(VARIANT* pvarKey, VARIANT* pvar)
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
	CheckTimeout();

	CRBMap<CXVariant, LRUNode>::CPair* pPair = m_mapItems.Lookup(*(CXVariant*)pvarKey);

	if(pPair == NULL)
	{
		static LRUNode varTemp;

		pPair = (CRBMap<CXVariant, LRUNode>::CPair*)m_mapItems.SetAt(*(CXVariant*)pvarKey, varTemp);
	}else
	{
		m_listAccess.RemoveAt(pPair->m_value.posAccess);
		m_listTimeout.RemoveAt(pPair->m_value.posTimeout);
	}

	VariantCopyInd(&pPair->m_value.var, pvar);
	pPair->m_value.posAccess = m_listAccess.AddTail(pPair);
	pPair->m_value.posTimeout = m_listTimeout.AddTail(pPair);

	pPair->m_value.timeout = GetLocalFileTime();

	m_cs.Leave();

	if(varTempKey.vt != VT_EMPTY)VariantClear(&varTempKey);

	return S_OK;
}

STDMETHODIMP CXLRUCache::put_Item(VARIANT key, VARIANT newVal)
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

STDMETHODIMP CXLRUCache::putref_Item(VARIANT key, VARIANT newVal)
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

STDMETHODIMP CXLRUCache::GetItem(VARIANT key, VARIANT* pVal)
{
	VARIANT* pkey = &key;
	BOOL bFound = FALSE;

	while(pkey && pkey->vt == (VT_VARIANT | VT_BYREF))
		pkey = pkey->pvarVal;

	if(!pkey || (pkey->vt & VT_ARRAY))return E_INVALIDARG;
	if(pkey->vt == VT_ERROR)return DISP_E_PARAMNOTOPTIONAL;

	m_cs.Enter();
	CheckTimeout();

	CRBMap<CXVariant, LRUNode>::CPair* pPair;

	if(pPair = m_mapItems.Lookup(*(CXVariant*)pkey))
	{
		bFound = TRUE;
		VariantCopyInd(pVal, &pPair->m_value.var);
	}

	m_cs.Leave();

	getCount ++;
	if(bFound)
		hitCount ++;

	return bFound ? S_OK : S_FALSE;
}

STDMETHODIMP CXLRUCache::Pop(VARIANT key, VARIANT* pVal)
{
	VARIANT* pkey = &key;
	BOOL bFound = FALSE;

	while(pkey && pkey->vt == (VT_VARIANT | VT_BYREF))
		pkey = pkey->pvarVal;

	if(!pkey || (pkey->vt & VT_ARRAY))return E_INVALIDARG;
	if(pkey->vt == VT_ERROR)return DISP_E_PARAMNOTOPTIONAL;

	m_cs.Enter();
	CheckTimeout();

	CRBMap<CXVariant, LRUNode>::CPair* pPair;

	if(pPair = m_mapItems.Lookup(*(CXVariant*)pkey))
	{
		bFound = TRUE;
		VariantCopyInd(pVal, &pPair->m_value.var);

		removeNode(pPair);
	}

	m_cs.Leave();

	getCount ++;
	if(bFound)
		hitCount ++;

	return bFound ? S_OK : S_FALSE;
}

STDMETHODIMP CXLRUCache::Count(LONG* pVal)
{
	m_cs.Enter();
	CheckTimeout();
	*pVal = (long)m_mapItems.GetCount();
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXLRUCache::Remove(VARIANT key)
{
	VARIANT* pkey = &key;

	while(pkey && pkey->vt == (VT_VARIANT | VT_BYREF))
		pkey = pkey->pvarVal;

	if(!pkey || (pkey->vt & VT_ARRAY))return E_INVALIDARG;
	if(pkey->vt == VT_ERROR)return DISP_E_PARAMNOTOPTIONAL;

	m_cs.Enter();
	CheckTimeout();

	CRBMap<CXVariant, LRUNode>::CPair* pPair;

	if(pPair = m_mapItems.Lookup(*(CXVariant*)pkey))
		removeNode(pPair);

	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXLRUCache::RemoveAll(void)
{
	m_cs.Enter();
	m_mapItems.RemoveAll();
	m_listAccess.RemoveAll();
	m_listTimeout.RemoveAll();
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXLRUCache::Add(VARIANT key, VARIANT value)
{
	return putItem(&key, &value);
}

STDMETHODIMP CXLRUCache::Exists(VARIANT key, VARIANT_BOOL* pVal)
{
	VARIANT* pkey = &key;

	while(pkey && pkey->vt == (VT_VARIANT | VT_BYREF))
		pkey = pkey->pvarVal;

	if(!pkey || (pkey->vt & VT_ARRAY))return E_INVALIDARG;
	if(pkey->vt == VT_ERROR)return DISP_E_PARAMNOTOPTIONAL;

	m_cs.Enter();
	CheckTimeout();
	*pVal = m_mapItems.Lookup(*(CXVariant*)pkey) ? VARIANT_TRUE : VARIANT_FALSE;
	m_cs.Leave();

	getCount ++;
	if(*pVal == VARIANT_TRUE)
		hitCount ++;

	return S_OK;
}

STDMETHODIMP CXLRUCache::get_MaxSize(LONG* pVal)
{
	m_cs.Enter();
	*pVal = m_nSize;
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXLRUCache::put_MaxSize(LONG newVal)
{
	if(newVal > 0)
	{
		m_cs.Enter();

		m_nSize = newVal;
		CheckTimeout();

		m_cs.Leave();
	}

	return S_OK;
}

STDMETHODIMP CXLRUCache::get_Timeout(LONG* pVal)
{
	m_cs.Enter();
	*pVal = m_nTimeout;
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXLRUCache::put_Timeout(LONG newVal)
{
	m_cs.Enter();

	if(newVal > 0)m_nTimeout = newVal;
	else m_nTimeout = -1;

	CheckTimeout();

	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXLRUCache::get_MaxTime(LONG* pVal)
{
	CRBMap<CXVariant, LRUNode>::CPair* pPair;

	m_cs.Enter();
	CheckTimeout();

	if(m_listTimeout.GetCount() > 0)
	{
		pPair = (CRBMap<CXVariant, LRUNode>::CPair*)m_listTimeout.GetHead();
		*pVal = (long)(GetLocalFileTime() - pPair->m_value.timeout);
	}else *pVal = 0;

	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXLRUCache::get_HitRate(float* pVal)
{
	*pVal = (float)((double)hitCount / getCount);

	return S_OK;
}

STDMETHODIMP CXLRUCache::get__NewEnum(IUnknown** pVal)
{
	return getNewEnum(this, pVal);
}

void CXLRUCache::FillEnum(CAtlArray<VARIANT>& arrayVariant)
{
	POSITION pos;
	CRBMap<CXVariant, LRUNode>::CPair * pPair;

	m_cs.Enter();
	CheckTimeout();

	pos = m_listTimeout.GetTailPosition();

	while(pos)
	{
		VARIANT var = {VT_EMPTY};

		pPair = (CRBMap<CXVariant, LRUNode>::CPair*)m_listTimeout.GetPrev(pos);
		VariantCopy(&var, (VARIANT*)&pPair->m_key);
		arrayVariant.Add(var);
	}

	m_cs.Leave();
}

