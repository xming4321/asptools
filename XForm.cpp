// XForm.cpp : Implementation of CXForm

#include "stdafx.h"
#include "XForm.h"
#include "XEncoding.h"
#include "XMemStream.h"

// CXForm

#define FORM_BLOCK_SIZE		4096

STDMETHODIMP CXForm::OnStartPage (IUnknown* pUnk)  
{
	if(!pUnk)
		return S_FALSE;

	CComPtr<IScriptingContext> spContext;
	HRESULT hr;

	// Get the IScriptingContext Interface
	hr = pUnk->QueryInterface(__uuidof(IScriptingContext), (void **)&spContext);
	if(FAILED(hr))
		return hr;

	CComPtr<IServer> piServer;
	hr = spContext->get_Server(&piServer);
	if(FAILED(hr))return hr;

	hr = piServer->put_ScriptTimeout(2 * 60 * 60);
	if(FAILED(hr))return hr;

	CComPtr<IRequest> piRequest;
	hr = spContext->get_Request(&piRequest);
	if(FAILED(hr))return hr;

	long nSize = 0;

	hr = piRequest->get_TotalBytes(&nSize);
	if(FAILED(hr))return hr;

	if(!nSize)return S_OK;

	CComVariant varLen, varData;
	CXMemStream mBuffer;

	varLen = FORM_BLOCK_SIZE;
	while(nSize)
	{
		CXVarPtr vPtr;

		if(nSize < FORM_BLOCK_SIZE)
		{
			varLen = nSize;
			nSize = 0;
		}else nSize -= FORM_BLOCK_SIZE;

		try
		{
			hr = piRequest->BinaryRead(&varLen, &varData);
			if(FAILED(hr))return S_OK;
		}catch(...)
		{return S_OK;}

		hr = vPtr.Attach(varData);
		if(FAILED(hr))return S_OK;

		hr = mBuffer.Write(vPtr.m_pData, vPtr.m_nSize);
		if(FAILED(hr))return S_OK;

		varData.Clear();
	}

	nSize = (long)mBuffer.GetLength();
	if(!nSize)return S_OK;

	CXAutoPtr<char> pBuffer;
	CComVariant varIn(_T("CONTENT_TYPE"));
	CComVariant varType;
	CComPtr<IRequestDictionary> piReadDict;
	CXString strType;

	pBuffer.Allocate(nSize);
	mBuffer.SeekToBegin();
	mBuffer.Read(pBuffer, nSize);
	mBuffer.SetSize(0);

	hr = piRequest->get_ServerVariables(&piReadDict);
	if(FAILED(hr))return hr;

	hr = piReadDict->get_Item(varIn, &varType);
	if(FAILED(hr))return hr;

	hr = varGetString(varType, strType);
	if(FAILED(hr))return hr;

	if(!strType.CompareNoCase(L"application/x-www-form-urlencoded"))
		return ParseUrlEncodeString(pBuffer, nSize);
	else return ParseUploadString(pBuffer, nSize);

	return S_OK;
}

STDMETHODIMP CXForm::OnEndPage ()  
{
	return S_OK;
}

static CXUploadList s_EmptyList;

STDMETHODIMP CXForm::get_Item(BSTR key, IXUploadList** pVariantReturn)
{
	CXComPtr<CXUploadList> pList;

	pList = &s_EmptyList;

	m_mapForm.Lookup(key, pList);

	*pVariantReturn = pList.Detach();

	return S_OK;
}

STDMETHODIMP CXForm::get__NewEnum(IUnknown** ppEnumReturn)
{
	return getNewEnum(this, ppEnumReturn);
}

STDMETHODIMP CXForm::get_Count(long* cStrRet)
{
	*cStrRet = (long)m_mapForm.GetCount();
	return S_OK;
}

STDMETHODIMP CXForm::Exists(BSTR key, VARIANT_BOOL* pExists)
{
	*pExists = m_mapForm.Lookup(key) ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}

void CXForm::FillEnum(CAtlArray<VARIANT>& arrayVariant)
{
	POSITION pos;
	CRBMap<CXKeyString, CXComPtr<CXUploadData> >::CPair* pPair;

	pos = m_mapForm.GetHeadPosition();

	while(pos)
	{
		VARIANT var = {VT_BSTR};

		pPair = (CRBMap<CXKeyString, CXComPtr<CXUploadData> >::CPair*)m_mapForm.GetNext(pos);
		var.bstrVal = pPair->m_key.m_str.AllocSysString();
		arrayVariant.Add(var);
	}
}

HRESULT CXForm::ParseUrlEncodeString(LPCSTR pstr, UINT nSize)
{
	LPCSTR pstrTemp;
	CXString strKey, strValue;

	while(nSize)
	{
		pstrTemp = pstr;
		while(nSize && *pstr != '=')
		{
			pstr ++;
			nSize --;
		}

		if(pstr > pstrTemp)
			strKey = CXEncoding::UrlDecode(pstrTemp, (UINT)(pstr - pstrTemp));
		else strKey.Empty();

		if(nSize)
		{
			nSize --;
			pstr ++;
		}

		pstrTemp = pstr;
		while(nSize && *pstr != '&')
		{
			pstr ++;
			nSize --;
		}

		if(!strKey.IsEmpty())
			if(pstr > pstrTemp)
				strValue = CXEncoding::UrlDecode(pstrTemp, (UINT)(pstr - pstrTemp));
			else strValue.Empty();

		if(nSize)
		{
			nSize --;
			pstr ++;
		}

		if(!strKey.IsEmpty())
		{
			CXComPtr<CXUploadData> pData;

			pData.CreateObject();
			pData->m_varData = strValue;
			AddValue(strKey, pData);
		}
	}

	return S_OK;
}

HRESULT CXForm::ParseUploadString(LPCSTR pstr, UINT nSize)
{
	CXStringA strName;
	CXStringA strFileName;
	CXStringA strContentType;
	const BYTE *p, *p1, *p2, *szQueryString;
	const BYTE *pstrSplit;
	UINT uiSplitSize, uiSize;
	BYTE ch;

	szQueryString = (const BYTE *)pstr;

	if(nSize < 2 || szQueryString[0] != '-' || szQueryString[1] != '-')
		return S_OK;

	p = szQueryString;
	while(nSize > 0 && *p >= ' ')
	{
		nSize --;
		p ++;
	}

	pstrSplit = szQueryString;
	uiSplitSize = (UINT)(p - szQueryString);
	szQueryString = p;

	while(nSize)
	{
		strFileName.Empty();
		strContentType.Empty();

		while(nSize > 0)
		{
			ch = *szQueryString ++;
			nSize --;
			if(ch != '\r' && ch != '\n')return S_OK;
			if(nSize > 0 && *szQueryString + ch == '\r' + '\n')
			{
				nSize --;
				szQueryString ++;
			}

			p = szQueryString;
			while(nSize > 0 && *p >= ' ')
			{
				nSize --;
				p ++;
			}

			p1 = szQueryString;
			szQueryString = p;

			if(p != p1)
			{
				//Content-Disposition: form-data; name="textfi&lg;<<sseld"
				if(p1 + 20 < p && !_strnicmp((char*)p1, "Content-Disposition:", 20))
				{
					p1 += 20;
					while(p1 < p && *p1 == ' ')p1 ++;
					if(p1 + 10 >= p || _strnicmp((char*)p1, "form-data;", 10))
						return S_OK;

					p1 += 10;
					while(p1 < p && *p1 == ' ')p1 ++;
					if(p1 + 4 >= p || _strnicmp((char*)p1, "name", 4))
						return S_OK;

					p1 += 4;
					while(p1 < p && *p1 == ' ')p1 ++;
					if(p1 + 1 >= p || *p1 != '=')
						return S_OK;

					p1 ++;
					while(p1 < p && *p1 == ' ')p1 ++;

					ch = ';';
					if(*p1 == '\"')
					{
						p1 ++;
						ch = '\"';
					}

					p2 = p1;
					while(p1 < p && *p1 != ch)p1 ++;

					strName.SetString((char*)p2, (int)(p1 - p2));
					if(p1 < p && *p1 == '\"')p1 ++;
					if(p1 < p && *p1 == ';')p1 ++;

					while(p1 < p && *p1 == ' ')p1 ++;
					if(p1 + 8 < p && !_strnicmp((char*)p1, "filename", 8))
					{
						p1 += 8;
						while(p1 < p && *p1 == ' ')p1 ++;
						if(p1 + 1 >= p || *p1 != '=')
							return S_OK;

						p1 ++;
						while(p1 < p && *p1 == ' ')p1 ++;

						ch = ';';
						if(*p1 == '\"')
						{
							p1 ++;
							ch = '\"';
						}

						p2 = p1;
						while(p1 < p && *p1 != ch)
						{
							if(*p1 == '/' || *p1 == '\\')
								p2 = p1 + 1;
							p1 ++;
						}

						strFileName.SetString((char*)p2, (int)(p1 - p2));
					}
				}else if(p1 + 13 < p && !_strnicmp((char*)p1, "Content-Type:", 13))
				{
					p1 += 13;
					while(p1 < p && *p1 == ' ')p1 ++;
					strContentType.SetString((char*)p1, (int)(p - p1));
				}
			}else
			{
				ch = *szQueryString ++;
				nSize --;
				if(ch != '\r' && ch != '\n')return S_OK;
				if(nSize > 0 && *szQueryString + ch == '\r' + '\n')
				{
					nSize --;
					szQueryString ++;
				}
				break;
			}
		}

		p = szQueryString;
		p1 = p + nSize;
		while(p1 > p && (p = (BYTE*)memchr(p, '-', p1 - p)) &&
			p1 > p + uiSplitSize &&
			memcmp(p, pstrSplit, uiSplitSize))
			p ++;

		if(!p || p1 <= p + uiSplitSize)break;

		nSize = (int)(p1 - p);
		p1 = szQueryString;
		szQueryString = p + uiSplitSize;

		p --;
		ch = *p;
		if(ch != '\r' && ch != '\n')return S_OK;
		if(*(p - 1) + ch == '\r' + '\n')p --;

		if(!strName.IsEmpty())
		{
			CXComPtr<CXUploadData> pUploadData;

			uiSize = (UINT)(p - p1);
			if(strFileName.IsEmpty())
			{
				pUploadData.CreateObject();
				pUploadData->m_varData = CXString((char*)p1, uiSize);
			}else
			{
				CXVarPtr varPtr;
				HRESULT hr;

				hr = varPtr.Create(uiSize);
				if(FAILED(hr))return hr;

				CopyMemory(varPtr.m_pData, p1, uiSize);

				pUploadData.CreateObject();

				hr = varPtr.GetVariant(pUploadData->m_varData);
				if(FAILED(hr))return hr;

				pUploadData->m_strFileName = strFileName;
				pUploadData->m_nSize = uiSize;
				pUploadData->m_strContentType = strContentType;
			}

			AddValue(CXString(strName), pUploadData);
		}
	}

	return S_OK;
}
