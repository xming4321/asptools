// XDocContent.h : Declaration of the CXDocContent

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"
#include "atlib.h"

// CXDocContent

class CXDocContent : public CXDispatch<IXDocContent>
{
public:
	STDMETHOD(get_Time)(DATE* pVal);
	STDMETHOD(put_Time)(DATE newVal);
	STDMETHOD(get_Label)(BSTR* pVal);
	STDMETHOD(put_Label)(BSTR newVal);
	STDMETHOD(get_Text)(BSTR* pVal);
	STDMETHOD(put_Text)(BSTR newVal);

	void SetText(BSTR strText, DATE date, CXString& strLabel)
	{
		m_strText = strText;
		m_strLabel = strLabel;
		m_date = date;
	}

public:
	DATE m_date;
	CXString m_strLabel;
	CXString m_strText;
};
