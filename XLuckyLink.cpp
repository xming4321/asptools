// XLuckyLink.cpp : Implementation of CXLuckyLink

#include "stdafx.h"
#include "XLuckyLink.h"


// CXLuckyLink

STDMETHODIMP CXLuckyLink::get_halfLife(int* pVal)
{
	*pVal = m_nhalfLifeTimes;

	return S_OK;
}

STDMETHODIMP CXLuckyLink::put_halfLife(int newVal)
{
	if(newVal > 0)
		m_nhalfLifeTimes = newVal;

	return S_OK;
}

CXLuckyLink::CXTree* CXLuckyLink::CXTree::GetNode(LPCWSTR mark, BOOL bCreate)
{
	WCHAR ch = *mark;
	CXTree* pNode = this;

	while(ch)
	{
		if((ch >= 'a' && ch <='z') || (ch >= 'A' && ch <='Z'))
		{
			mark ++;
			ch = (ch & 0x1f) - 1;

			if(bCreate)
			{
				if(pNode->m_aSubTrees.IsEmpty())
					pNode->m_aSubTrees.SetCount(26);

				if(!pNode->m_aSubTrees[ch])
				{
					pNode->m_aSubTrees[ch].CreateObject();
					pNode->m_aSubTrees[ch]->m_pParent = pNode;
				}
			}else
			{
				if(pNode->m_aSubTrees.IsEmpty())
					return NULL;

				if(!pNode->m_aSubTrees[ch])
					return NULL;
			}

			pNode = pNode->m_aSubTrees[ch];
			ch = *mark;
		}else return NULL;
	}

	return pNode;
}

void CXLuckyLink::CXTree::UpdateLife(int count)
{
	CXTree *pNode = this;
	POSITION myPos = m_aNodes.GetHeadPosition();
	UINT i;

	m_nSelfLifes = 0;

	while(myPos)
		m_nSelfLifes += m_aNodes.GetNext(myPos)->m_nLife;

	while(pNode)
	{
		pNode->m_nNodeCount += count;
		pNode->m_nLifes = pNode->m_nSelfLifes;

		for(i = 0; i < pNode->m_aSubTrees.GetCount(); i ++)
			if(pNode->m_aSubTrees[i])pNode->m_nLifes += pNode->m_aSubTrees[i]->m_nLifes;

		pNode = pNode->m_pParent;
	}
}

void CXLuckyLink::CXTree::HalfLife(double life)
{
	UINT i;

	for(i = 0; i < m_aSubTrees.GetCount(); i ++)
		if(m_aSubTrees[i])
			m_aSubTrees[i]->HalfLife(life);

	POSITION myPos = m_aNodes.GetHeadPosition();

	while(myPos)
		m_aNodes.GetNext(myPos)->m_nLife /= life;

	m_nLifes /= life;
}

STDMETHODIMP CXLuckyLink::get_Count(int* pVal)
{
	*pVal = m_treeRoot.m_nNodeCount;
	return S_OK;
}

HRESULT CXLuckyLink::AddOneItem(LPCWSTR mark, double life, DATE time, int id, LPCWSTR value)
{
	CXTree* pNode;
	CXComPtr<CXNode> pItem;

	if(pNode = m_treeRoot.GetNode(mark, TRUE))
	{
		if(life >= 0.0)
		{
			pItem.CreateObject();
			life = life * pow(2, (time - m_baseTime) * 24 / m_nhalfLifeTimes);
			pItem->m_nLife = life;
			pItem->m_strValue = value;
			pItem->m_nID = id;

			CRBMap< int, CXComPtr<CXNode> >::CPair* pPair;

			pPair = (CRBMap< int, CXComPtr<CXNode> >::CPair*)m_mapNodes.Lookup(id);

			if(pPair)
			{
				pPair->m_value->m_pParent->m_aNodes.RemoveAt(pPair->m_value->m_pos);
				pPair->m_value->m_pParent->UpdateLife(-1);

				pPair->m_value = pItem;
			}else m_mapNodes.SetAt(id, pItem);

			pItem->m_pParent = pNode;
			pItem->m_pos = pNode->m_aNodes.AddTail(pItem);

			pNode->UpdateLife(1);
		}else
		{
			CRBMap< int, CXComPtr<CXNode> >::CPair* pPair;

			pPair = (CRBMap< int, CXComPtr<CXNode> >::CPair*)m_mapNodes.Lookup(id);

			if(pPair)
			{
				pPair->m_value->m_pParent->m_aNodes.RemoveAt(pPair->m_value->m_pos);
				pPair->m_value->m_pParent->UpdateLife(-1);

				m_mapNodes.RemoveAt(pPair);
			}
		}
	}else return DISP_E_BADCALLEE;

	return S_OK;
}

STDMETHODIMP CXLuckyLink::AddItem(BSTR mark, double life, DATE time, int id, BSTR value)
{
	HRESULT hr;

	m_cs.Enter();
	hr = AddOneItem(mark, life, time, id, value);
	m_cs.Leave();

	return hr;
}

void inline getSplitString(BSTR& batchText, CXString& str)
{
	BSTR strTemp = batchText;
	WCHAR ch = 0;

	while((ch = *batchText) && (ch != 13))
		batchText ++;

	str.SetString(strTemp, batchText - strTemp);

	if(*batchText)batchText ++;
}

STDMETHODIMP CXLuckyLink::BatchAdd(BSTR batchText)
{
	HRESULT hr = S_OK;
	CXString mark, value, str;
	double life;
	DATE time;
	int id;

	m_cs.Enter();

	while(*batchText)
	{
		getSplitString(batchText, mark);
		getSplitString(batchText, str);
		life = _wtof(str);
		getSplitString(batchText, str);
		hr = ::VarDateFromStr((OLECHAR*)(LPCWSTR)str, 0, 0, &time);
		if(FAILED(hr))break;
		getSplitString(batchText, str);
		id = _wtol(str);
		getSplitString(batchText, value);

		if(id && !value.IsEmpty() && !mark.IsEmpty())
		{
			hr = AddOneItem(mark, life, time, id, value);
			if(FAILED(hr))break;
		}
	}
	m_cs.Leave();

	return hr;
}

CXLuckyLink::CXNode* CXLuckyLink::CXTree::GetItem(__int64 nAct)
{
	double nLifeTemp, nLife;
	UINT nCountTemp, nCount;
	UINT i;
	CXNode* p;

	if(nLife = GetLifes(nAct))
	{
		nLifeTemp = rand32() * nLife / 0x100000000;

		for(i = 0; i < m_aSubTrees.GetCount(); i ++)
		{
			if(m_aSubTrees[i])
			{
				if(m_aSubTrees[i]->GetCount(nAct) && m_aSubTrees[i]->GetLifes(nAct) > nLifeTemp)
				{
					p = m_aSubTrees[i]->GetItem(nAct);
					m_nCountAct --;
					m_nLifesAct -= p->m_nLife;

					return p;
				}else nLifeTemp -= m_aSubTrees[i]->GetLifes(nAct);
			}
		}

		POSITION myPos = m_aNodes.GetHeadPosition();

		while(myPos)
		{
			p = m_aNodes.GetNext(myPos);

			if(p->m_nAct != nAct)
			{
				if(p->m_nLife > nLifeTemp)
				{
					p->m_nAct = nAct;

					m_nCountAct --;
					m_nLifesAct -= p->m_nLife;

					return p;
				}else nLifeTemp -= p->m_nLife;
			}
		}
	}

	if(nCount = GetCount(nAct))
	{
		nCountTemp = (UINT)(rand32() * (__int64)nCount / 0x100000000);

		for(i = 0; i < m_aSubTrees.GetCount(); i ++)
		{
			if(m_aSubTrees[i])
			{
				if(m_aSubTrees[i]->GetCount(nAct) > nCountTemp)
				{
					p = m_aSubTrees[i]->GetItem(nAct);
					m_nCountAct --;
					m_nLifesAct -= p->m_nLife;

					return p;
				}else nCountTemp -= m_aSubTrees[i]->GetCount(nAct);
			}
		}

		POSITION myPos = m_aNodes.GetHeadPosition();

		while(myPos)
		{
			p = m_aNodes.GetNext(myPos);

			if(p->m_nAct != nAct)
			{
				if(!nCountTemp)
				{
					p->m_nAct = nAct;

					m_nCountAct --;
					m_nLifesAct -= p->m_nLife;

					return p;
				}else nCountTemp --;
			}
		}
	}

	return NULL;
}

STDMETHODIMP CXLuckyLink::RemoveItem(int id)
{
	CRBMap< int, CXComPtr<CXNode> >::CPair* pPair;

	m_cs.Enter();
	if(pPair = (CRBMap< int, CXComPtr<CXNode> >::CPair*)m_mapNodes.Lookup(id))
	{
		pPair->m_value->m_pParent->m_aNodes.RemoveAt(pPair->m_value->m_pos);
		pPair->m_value->m_pParent->UpdateLife(-1);

		m_mapNodes.RemoveAt(pPair);
	}
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXLuckyLink::RemoveAll()
{
	m_cs.Enter();
	m_treeRoot.RemoveAll();
	m_mapNodes.RemoveAll();
	m_cs.Leave();
	return S_OK;
}

STDMETHODIMP CXLuckyLink::GetList(BSTR mark, int nCount, VARIANT seed, VARIANT* pVal)
{
	if(nCount <= 0)return DISP_E_BADCALLEE;

	CComSafeArray<VARIANT> bstrArray;
	VARIANT* pVar;
	HRESULT hr;
	CAtlArray< CXComPtr<CXNode> > aNodes;
	CXTree* pNode;
	CXNode* pItem;
	UINT i;

	if(i = varGetNumbar(seed))
		srand(i);

	m_cs.Enter();
	if(pNode = m_treeRoot.GetNode(mark))
	{
		if((UINT)nCount > pNode->m_nNodeCount)
			nCount = pNode->m_nNodeCount;

		m_nAct ++;

		for(i = 0; i < (UINT)nCount; i ++)
			if(pItem = pNode->GetItem(m_nAct))
				aNodes.Add(pItem);
			else break;
	}
	m_cs.Leave();

	hr = bstrArray.Create(aNodes.GetCount());
	if(FAILED(hr))return hr;

	pVar = (VARIANT*)bstrArray.m_psa->pvData;

	for(i = 0; i < aNodes.GetCount(); i ++)
	{
		pVar->vt = VT_BSTR;
		pVar->bstrVal = aNodes[i]->m_strValue.AllocSysString();
		pVar ++;
	}

	pVal->vt = VT_ARRAY | VT_VARIANT;
	pVal->parray = bstrArray.Detach();

	return S_OK;
}

