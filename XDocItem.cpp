// XDocItem.cpp : Implementation of CXDocItem

#include "stdafx.h"
#include "XDoc.h"
#include "XDocItem.h"

// CXDocItem


HRESULT CXDocItem::Save(IStream* pFile)
{
	CXStream StreamStub(pFile);
	HRESULT hr;

	W(m_strUserName);
	W(m_nUserID);
	W(m_nEmote);
	W(m_nRate);
	W(m_strHost);
	W(m_strIP);
	W(m_nisDel);
	W(m_nHot);
	W(m_dateDel);
	W(m_strDelUser);

	int count, i;

	count = (int)m_Contents->GetCount();
	W(count);

	for(i = 0; i < count; i ++)
	{
		CXComPtr<CXDocContent> pContent = m_Contents->GetValue(i);

		W(pContent->m_date);
		W(pContent->m_strLabel);
		W(pContent->m_strText);
	}

	count = (int)m_Attachments->GetCount();
	W(count);

	for(i = 0; i < count; i ++)
	{
		W(m_Attachments->GetValue(i));
	}

	return S_OK;
}

HRESULT CXDocItem::Load(IStream* pFile)
{
	CXStream StreamStub(pFile);
	HRESULT hr;

	R(m_strUserName);
	R(m_nUserID);
	R(m_nEmote);
	R(m_nRate);
	R(m_strHost);
	R(m_strIP);
	R(m_nisDel);
	R(m_nHot);
	R(m_dateDel);
	R(m_strDelUser);

	if(m_nHot == 0)
		m_nHot = 10;

	int count, i;

	R(count);

	for(i = 0; i < count; i ++)
	{
		CXComPtr<CXDocContent> pContent;

		pContent.CreateObject();

		R(pContent->m_date);
		R(pContent->m_strLabel);
		R(pContent->m_strText);

		m_Contents->AddValue(pContent);
	}

	R(count);

	for(i = 0; i < count; i ++)
	{
		CXString strTemp;

		R(strTemp);

		m_Attachments->AddValue(strTemp);
	}

	return S_OK;
}

STDMETHODIMP CXDocItem::get_UserName(BSTR* pVal)
{
	*pVal = m_strUserName.AllocSysString();

	return S_OK;
}

STDMETHODIMP CXDocItem::get_UserID(LONG* pVal)
{
	*pVal = m_nUserID;

	return S_OK;
}

STDMETHODIMP CXDocItem::get_Emote(SHORT* pVal)
{
	*pVal = m_nEmote;

	return S_OK;
}

STDMETHODIMP CXDocItem::get_Rate(SHORT* pVal)
{
	*pVal = m_nRate;

	return S_OK;
}

STDMETHODIMP CXDocItem::get_Host(BSTR* pVal)
{
	*pVal = m_strHost.AllocSysString();

	return S_OK;
}

STDMETHODIMP CXDocItem::get_IP(BSTR* pVal)
{
	*pVal = m_strIP.AllocSysString();

	return S_OK;
}

STDMETHODIMP CXDocItem::Delete(BSTR strUser, VARIANT varSystem)
{
	m_nisDel = (char)varGetNumbar(varSystem, -1);
	m_dateDel = GetVariantTime();
	m_strDelUser = strUser;

	return S_OK;
}

STDMETHODIMP CXDocItem::Restore(void)
{
	m_nisDel = 0;
	m_dateDel = GetVariantTime();
	m_strDelUser.Empty();

	return S_OK;
}

STDMETHODIMP CXDocItem::isDel(short* pval)
{
	*pval = m_nisDel;

	return S_OK;
}

STDMETHODIMP CXDocItem::get_Hot(short* pval)
{
	*pval = m_nHot;

	return S_OK;
}

STDMETHODIMP CXDocItem::get_delTime(DATE* pVal)
{
	*pVal = m_dateDel;

	return S_OK;
}

STDMETHODIMP CXDocItem::get_delUser(BSTR* pVal)
{
	*pVal = m_strDelUser.AllocSysString();

	return S_OK;
}

STDMETHODIMP CXDocItem::get_updateTime(DATE* pVal)
{
	DATE d;
	CXComPtr<CXDocContent> pContent = m_Contents->GetValue((int)m_Contents->GetCount() - 1);

	d = pContent->m_date;
	if( d < m_dateDel)
		d = m_dateDel;

	*pVal = d;

	return S_OK;
}

STDMETHODIMP CXDocItem::newContent(BSTR strText, DATE date, VARIANT varLabel)
{
	CXComPtr<CXDocContent> pContent;
	CXString strLabel;

	varGetString(varLabel, strLabel);

	pContent.CreateObject();
	pContent->SetText(strText, date, strLabel);

	m_Contents->AddValue(pContent);

	return S_OK;
}

STDMETHODIMP CXDocItem::newAttachment(BSTR filename)
{
	m_Attachments->AddValue(filename);

	return S_OK;
}

STDMETHODIMP CXDocItem::get_Contents(IXList** pVal)
{
	return m_Contents.QueryInterface(pVal);
}

STDMETHODIMP CXDocItem::get_Attachments(IXList** pVal)
{
	return m_Attachments.QueryInterface(pVal);
}

