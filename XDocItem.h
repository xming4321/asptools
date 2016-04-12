// XDocItem.h : Declaration of the CXDocItem

#pragma once
#include "resource.h"       // main symbols
#include "asptools.h"
#include "atlib.h"

#include "XList.h"
#include "XFileStream.h"
#include "XDocContent.h"

// CXDocItem

class CXDocItem : public CXDispatch<IXDocItem>
{
public:
	CXDocItem() :
		m_nUserID(0),
		m_nEmote(0),
		m_nRate(0),
		m_nHot(0),
		m_nisDel(0),
		m_dateDel(0)
	{
		m_Contents.CreateObject();
		m_Attachments.CreateObject();
	}

public:
	STDMETHOD(get_UserName)(BSTR* pVal);
	STDMETHOD(get_UserID)(LONG* pVal);
	STDMETHOD(get_Emote)(SHORT* pVal);
	STDMETHOD(get_Rate)(SHORT* pVal);
	STDMETHOD(get_Host)(BSTR* pVal);
	STDMETHOD(get_IP)(BSTR* pVal);
	STDMETHOD(Delete)(BSTR strUser, VARIANT varSystem);
	STDMETHOD(Restore)(void);
	STDMETHOD(isDel)(short* pval);
	STDMETHOD(get_Hot)(short* pval);
	STDMETHOD(get_delTime)(DATE* pVal);
	STDMETHOD(get_delUser)(BSTR* pVal);
	STDMETHOD(get_updateTime)(DATE* pVal);
	STDMETHOD(newContent)(BSTR strText, DATE date, VARIANT varLabel);
	STDMETHOD(newAttachment)(BSTR filename);
	STDMETHOD(get_Contents)(IXList** pVal);
	STDMETHOD(get_Attachments)(IXList** pVal);

	void SetValue(BSTR user, LONG id, SHORT emote, SHORT rate, SHORT hot, BSTR host, BSTR ip)
	{
		if(hot < 1)hot = 1;

		m_strUserName = user;
		m_nUserID = id;
		m_nEmote = (BYTE)emote;
		m_nRate = (char)rate;
		m_nHot = (char)hot;
		m_strHost = host;
		m_strIP = ip;
	}

	HRESULT Save(IStream* pFile);
	HRESULT Load(IStream* pFile);

	BOOL isDel()
	{
		return m_nisDel;
	}

	BOOL HotNumber()
	{
		return m_nHot;
	}

	DATE updateTime()
	{
		CXComPtr<CXDocContent> pContent = m_Contents->GetValue((int)m_Contents->GetCount() - 1);
		return pContent->m_date;
	}

public:
	CXString m_strUserName;
	LONG m_nUserID;
	BYTE m_nEmote;
	char m_nRate;
	CXString m_strHost;
	CXString m_strIP;
	unsigned char m_nisDel;
	unsigned char m_nHot;
	DATE m_dateDel;
	CXString m_strDelUser;

	CXComPtr< CXList< CXComPtr<CXDocContent> > > m_Contents;
	CXComPtr< CXList<CXString> > m_Attachments;
};

