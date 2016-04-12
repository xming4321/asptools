// XEnKeyword.cpp : Implementation of CXEncoding

#include "stdafx.h"
#include "XEncoding.h"

extern "C" {
#include "scws\scws.h"
}

extern CXStringA s_pathRuntime;

STDMETHODIMP CXEncoding::GetKeyword(BSTR txtString, VARIANT attr, VARIANT* pVal)
{
	HRESULT hr;
	CComSafeArray<VARIANT> bstrArray;
	VARIANT* pVar;
	CXStringA strString(txtString, ::SysStringLen(txtString));
	scws_t s_scws;
	scws_top_t s_ttws, p;
	int count;
	CXStringA strAttr;

	varGetString(attr, strAttr);

	s_scws = scws_new();
	if(!s_scws)return DISP_E_EXCEPTION;

	scws_set_charset(s_scws, "gbk");
	scws_set_dict(s_scws, s_pathRuntime + "dict.xdb", SCWS_XDICT_XDB);
	scws_set_rule(s_scws, s_pathRuntime + "rules.ini");

	scws_send_text(s_scws, strString, strString.GetLength());

	s_ttws = scws_get_tops(s_scws, 20, strAttr.IsEmpty() ? NULL : (LPSTR)(LPCSTR)strAttr);

	p = s_ttws;
	count = 0;
	while(p)
	{
		count ++;
		p = p->next;
	}

	SAFEARRAYBOUND Bounds[2] = {{2, 0}, {count, 0}};

	hr = bstrArray.Create(Bounds, 2);
	if(FAILED(hr))
	{
		scws_free_tops(s_ttws);
		scws_free(s_scws);

		return hr;
	}

	if(count)
	{
		pVar = (VARIANT*)bstrArray.m_psa->pvData;

		p = s_ttws;
		count = 0;
		while(p)
		{
			pVar[count * 2].vt = VT_BSTR;
			pVar[count * 2].bstrVal = CXString(p->word).AllocSysString();
			pVar[count * 2 + 1].vt = VT_I4;
			pVar[count * 2 + 1].lVal = p->times;

			count ++;
			p = p->next;
		}

		pVal->vt = VT_ARRAY | VT_VARIANT;
		pVal->parray = bstrArray.Detach();

	}

	scws_free_tops(s_ttws);
	scws_free(s_scws);

	return S_OK;
}
