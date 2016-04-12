// XEncoding.cpp : Implementation of CXEncoding

#include "stdafx.h"
#include "XEncoding.h"
#include "XDate.h"


// CXEncoding


static HRESULT baseDecode(const char *pdecodeTable, UINT dwBits, BSTR baseString, VARIANT *retVal)
{
	UINT nWritten = 0, len = SysStringLen(baseString);
	CXVarPtr varPtr;

	varPtr.Create(len * dwBits / 8);

	if(len)
	{
		BSTR szEnd = baseString + len;
		DWORD dwCurr = 0;
		int nBits = 0;
		WCHAR ch;

		while(baseString < szEnd)
		{
			ch = *baseString++;
			int nCh = (ch > 0x20 && ch < 0x80) ? pdecodeTable[ch - 0x20] : -1;

			if (nCh != -1)
			{
				dwCurr <<= dwBits;
				dwCurr |= nCh;
				nBits += dwBits;

				while(nBits >= 8)
				{
					varPtr.m_pData[nWritten ++] = (BYTE) (dwCurr >> (nBits - 8));
					nBits -= 8;
				}
			}
		}
	}

	return varPtr.GetVariant(retVal, nWritten);
}

STDMETHODIMP CXEncoding::Base32Decode(BSTR base32String, VARIANT *retVal)
{
	static const char decodeTable[] =
	{
		-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1, /* 2x  !"#$%&'()*+,-./   */
		14,11,26,27,28,29,30,31,-1, 6,-1,-1,-1,-1,-1,-1, /* 3x 0123456789:;<=>?   */
		-1, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9,10,11,12,13,14, /* 4x @ABCDEFGHIJKLMNO   */
		15,16,17,18,19,20,21,22,23,24,25,-1,-1,-1,-1,-1, /* 5X PQRSTUVWXYZ[\]^_   */
		-1, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9,10,11,12,13,14, /* 6x `abcdefghijklmno   */
		15,16,17,18,19,20,21,22,23,24,25,-1,-1,-1,-1,-1  /* 7X pqrstuvwxyz{\}~DEL */
	};

	return baseDecode(decodeTable, 5, base32String, retVal);
}

static HRESULT baseEncode(const char *pEncodingTable, UINT dwBits, VARIANT varData, BSTR *retVal)
{
	CXVarPtr varPtr;
	UINT i, len = 0, len1, bits = 0;
	DWORD dwData = 0;
	DWORD dwSize;
	BYTE bMask = 0xff >> (8 - dwBits);

	HRESULT hr = varPtr.Attach(varData);
	if(FAILED(hr))return hr;

	if(dwBits == 6)dwSize = (varPtr.m_nSize + 2) / 3 * 4;
	else if(dwBits == 5)dwSize = (varPtr.m_nSize + 4) / 5 * 8;

	*retVal = ::SysAllocStringLen(NULL, dwSize);

	len1 = 0;
	for(i = 0; i < (UINT)varPtr.m_nSize; i ++)
	{
		dwData <<= 8;
		dwData |= varPtr.m_pData[i];
		bits += 8;

		while(bits >= dwBits)
		{
			(*retVal)[len ++] = pEncodingTable[(dwData >> (bits - dwBits)) & bMask];
			bits -= dwBits;
		}
	}

	if(bits)
	{
		(*retVal)[len ++] = pEncodingTable[(dwData << (dwBits - bits)) & bMask];
	}

	while(len < dwSize)
	{
		(*retVal)[len ++] = '=';
	}

	return S_OK;
}

STDMETHODIMP CXEncoding::Base32Encode(VARIANT varData, BSTR *retVal)
{
	return baseEncode("ABCDEFGHIJKLMNOPQRSTUVWXYZ234567", 5, varData, retVal);
}

STDMETHODIMP CXEncoding::Base64Decode(BSTR base64String, VARIANT *retVal)
{
	static const char decodeTable[] =
	{
		-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,62,-1,62,-1,63, /* 2x  !"#$%&'()*+,-./   */
		52,53,54,55,56,57,58,59,60,61,-1,-1,-1,-1,-1,-1, /* 3x 0123456789:;<=>?   */
		-1, 0, 1, 2, 3, 4, 5, 6, 7, 8, 9,10,11,12,13,14, /* 4x @ABCDEFGHIJKLMNO   */
		15,16,17,18,19,20,21,22,23,24,25,-1,-1,-1,-1,63, /* 5X PQRSTUVWXYZ[\]^_   */
		-1,26,27,28,29,30,31,32,33,34,35,36,37,38,39,40, /* 6x `abcdefghijklmno   */
		41,42,43,44,45,46,47,48,49,50,51,-1,-1,-1,-1,-1  /* 7X pqrstuvwxyz{\}~DEL */
	};

	return baseDecode(decodeTable, 6, base64String, retVal);
}

STDMETHODIMP CXEncoding::Base64Encode(VARIANT varData, BSTR *retVal)
{
	return baseEncode("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", 6, varData, retVal);
}

STDMETHODIMP CXEncoding::JSEncode(BSTR TextString, BSTR *retVal)
{
	CXAutoPtr<WCHAR> pszEncoded;
	UINT i, len = SysStringLen(TextString);
	UINT nPos = 0;
	WCHAR ch;

	pszEncoded.Allocate(len * 2);

	for(i = 0; i < len; i ++)
		switch(ch = TextString[i])
		{
		case '\\':
			pszEncoded[nPos ++] = '\\';
			pszEncoded[nPos ++] = '\\';
			break;
		case '/':
			pszEncoded[nPos ++] = '\\';
			pszEncoded[nPos ++] = '/';
			break;
		case '\r':
			pszEncoded[nPos ++] = '\\';
			pszEncoded[nPos ++] = 'r';
			break;
		case '\n':
			pszEncoded[nPos ++] = '\\';
			pszEncoded[nPos ++] = 'n';
			break;
		case '\t':
			pszEncoded[nPos ++] = '\\';
			pszEncoded[nPos ++] = 't';
			break;
		case '\'':
			pszEncoded[nPos ++] = '\\';
			pszEncoded[nPos ++] = '\'';
			break;
		case '\"':
			pszEncoded[nPos ++] = '\\';
			pszEncoded[nPos ++] = '\"';
			break;
		default:
			pszEncoded[nPos ++] = ch;
			break;
		}

	*retVal = ::SysAllocStringLen(pszEncoded, nPos);
	return S_OK;
}

#define HEX_ESCAPE '%'
static unsigned char isAcceptable[] =
{
	0,0,0,0,0,0,0,0,0,0,1,0,0,1,1,1, /* 2x  !"#$%&'()*+,-./   */
	1,1,1,1,1,1,1,1,1,1,1,0,0,0,0,0, /* 3x 0123456789:;<=>?   */
	1,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1, /* 4x @ABCDEFGHIJKLMNO   */
	1,1,1,1,1,1,1,1,1,1,1,0,0,0,0,1, /* 5X PQRSTUVWXYZ[\]^_   */
	0,1,1,1,1,1,1,1,1,1,1,1,1,1,1,1, /* 6x `abcdefghijklmno   */
	1,1,1,1,1,1,1,1,1,1,1,0,0,0,0,0  /* 7X pqrstuvwxyz{\}~DEL */
};
#define ACCEPTABLE(a)	( a>=32 && a<128 && (isAcceptable[a-32]))
static const char *hex = "0123456789ABCDEF";
#define HEXDATA(ch)		((ch) >= '0' && (ch) <= '9' ? (ch) - '0' : \
						(ch) >= 'a' && (ch) <= 'f' ? (ch) - 'a' + 10 : \
						(ch) >= 'A' && (ch) <= 'F' ? (ch) - 'A' + 10 : 0)
#define ISHEXCHA(ch)	(((ch) >= '0' && (ch) <= '9') || \
						((ch) >= 'a' && (ch) <= 'f') || \
						((ch) >= 'A' && (ch) <= 'F'))

CXString CXEncoding::UrlDecode(LPCSTR urlString, UINT len)
{
	CXString str;
	CXAutoPtr<char> pszDecoded;
	char ch;
	UINT pos;

	pszDecoded.Allocate(len + 1);
	pos = 0;
	while(len--)
	{
		switch(ch = *urlString ++)
		{
		case '+':
			pszDecoded[pos ++] = ' ';
			break;
		case '%':
			if(len >= 2 && ISHEXCHA(urlString[0]) && ISHEXCHA(urlString[1]))
			{
				pszDecoded[pos ++] = (char)((HEXDATA(urlString[0]) << 4) + HEXDATA(urlString[1]));
				urlString += 2;
				len -=2;
			}else if(len >= 5 && (urlString[0] == 'u' || urlString[0] == 'U') &&
					ISHEXCHA(urlString[1]) &&
					ISHEXCHA(urlString[2]) &&
					ISHEXCHA(urlString[3]) &&
					ISHEXCHA(urlString[4]))
			{
				urlString ++;
				len --;

				if(pos)
				{
					pszDecoded[pos] = 0;
					str += pszDecoded;
					pos = 0;
				}

				str.AppendChar((WCHAR)( (HEXDATA(urlString[0]) << 12) + (HEXDATA(urlString[1]) << 8) +
					(HEXDATA(urlString[2]) << 4) + HEXDATA(urlString[3])));

				urlString += 4;
				len -=4;
			}
			else pszDecoded[pos ++] = '%';
			break;
		default:
			pszDecoded[pos ++] = ch;
		}
	}

	if(pos)
	{
		pszDecoded[pos] = 0;
		str += pszDecoded;
		pos = 0;
	}

	return str;
}

STDMETHODIMP CXEncoding::UrlDecode(BSTR urlString, BSTR *retVal)
{
	*retVal = UrlDecode(CStringA(urlString, SysStringLen(urlString))).AllocSysString();
	return S_OK;
}

CXStringA CXEncoding::UrlEncode(LPCSTR urlString, UINT len, BOOL bEncodeDBMS)
{
	CXStringA str;
	LPSTR pszEncoded;
	UINT i, nLength = 0;
	BYTE ch;

	pszEncoded = str.GetBuffer(len * 3 + 1);

	for(i = 0; i < len; i ++)
	{
		ch = (BYTE)urlString[i];

		if(!bEncodeDBMS)
		{
			if(IsDBCSLeadByte(ch))
			{
				pszEncoded[nLength ++] = ch;
				if(i < len)
				{
					i ++;
					pszEncoded[nLength ++] = (BYTE)urlString[i];
				}

				continue;
			}
		}

		if (!ACCEPTABLE(ch))
		{
			pszEncoded[nLength ++] = HEX_ESCAPE;	/* Means hex commming */
			pszEncoded[nLength ++] = hex[(ch >> 4) & 15];
			pszEncoded[nLength ++] = hex[ch & 15];
		}
		else
			pszEncoded[nLength ++] = ch;
	}

	str.ReleaseBuffer(nLength);

	return str;
}

STDMETHODIMP CXEncoding::UrlEncode(BSTR urlString, VARIANT bDBCS, BSTR *retVal)
{
	*retVal = UrlEncode(CStringA(urlString, SysStringLen(urlString)), varGetNumbar(bDBCS)).AllocSysString();
	return S_OK;
}

STDMETHODIMP CXEncoding::QuotedDecode(BSTR QuotedString, VARIANT nCodePage, BSTR* retVal)
{
	return S_OK;
}

STDMETHODIMP CXEncoding::QuotedEncode(BSTR txtString, VARIANT nCodePage, BSTR* retVal)
{
	return S_OK;
}

STDMETHODIMP CXEncoding::BinToStr(VARIANT varData, VARIANT varCodePage, BSTR *retVal)
{
	CXVarPtr varPtr;
	int nCodePage = varGetNumbar(varCodePage);

	HRESULT hr = varPtr.Attach(varData);
	if(FAILED(hr))return hr;

	if(nCodePage == -1)
	{
		*retVal = ::SysAllocStringByteLen(NULL, varPtr.m_nSize);
		CopyMemory(*retVal, varPtr.m_pData, varPtr.m_nSize);
		return S_OK;
	}

	if(nCodePage == 0)nCodePage = _AtlGetConversionACP();

	int _nTempCount = ::MultiByteToWideChar(nCodePage, 0, (LPCSTR)varPtr.m_pData, varPtr.m_nSize, NULL, NULL);
	*retVal = ::SysAllocStringLen(NULL, _nTempCount);
	if(*retVal != NULL)
		::MultiByteToWideChar(nCodePage, 0, (LPCSTR)varPtr.m_pData, varPtr.m_nSize, *retVal, _nTempCount);

	return S_OK;
}

STDMETHODIMP CXEncoding::StrToBin(BSTR strData, VARIANT varCodePage, VARIANT *retVal)
{
	CXVarPtr varPtr;
	UINT strLen = SysStringByteLen(strData);
	int nCodePage = varGetNumbar(varCodePage);

	if(nCodePage == -1)
	{
		varPtr.Create(strLen);
		CopyMemory(varPtr.m_pData, strData, strLen);
		return varPtr.GetVariant(retVal);
	}

	if(nCodePage == 0)nCodePage = _AtlGetConversionACP();

	int _nTempCount = ::WideCharToMultiByte(nCodePage, 0, strData, strLen / 2, NULL, 0, NULL, NULL);
	varPtr.Create(_nTempCount);

	::WideCharToMultiByte(nCodePage, 0, strData, strLen / 2, (LPSTR)varPtr.m_pData, _nTempCount, NULL, NULL);

	return varPtr.GetVariant(retVal);
}

STDMETHODIMP CXEncoding::GMTDecode(BSTR GMTString, DATE* retVal)
{
	*(CXDate*)retVal = GMTString;
	return S_OK;
}

STDMETHODIMP CXEncoding::GMTEncode(DATE varDate, BSTR* retVal)
{
	CXString str;

	str = *(CXDate*)&varDate;
	*retVal = str.AllocSysString();
	return S_OK;
}

STDMETHODIMP CXEncoding::HexDecode(BSTR HexString, VARIANT *retVal)
{
	UINT i, pos, len = SysStringLen(HexString);
	CXVarPtr varPtr;
	WCHAR ch1, ch2;

	varPtr.Create(len / 2);

	pos = 0;
	for(i = 0; i < len; i ++)
	{
		ch1 = HexString[i];

		if((ch1 >= 'a' && ch1 <= 'f') || (ch1 >= 'A' && ch1 <= 'F'))
			ch1 = (ch1 & 0xf) + 9;
		else if(ch1 >= '0' && ch1 <= '9')
			ch1 &= 0xf;
		else continue;

		i ++;
		if(i < len)
		{
			ch2 = HexString[i];
			if((ch2 >= 'a' && ch2 <= 'f') || (ch2 >= 'A' && ch2 <= 'F'))
				ch2 = (ch2 & 0xf) + 9;
			else if(ch2 >= '0' && ch2 <= '9')
				ch2 &= 0xf;
			else
			{
				ch2 = ch1;
				ch1 = 0;
			}
		}else
		{
			ch2 = ch1;
			ch1 = 0;
		}

		varPtr.m_pData[pos ++] = (ch1 << 4) + ch2;
	}

	return varPtr.GetVariant(retVal, pos);
}

STDMETHODIMP CXEncoding::HexEncode(VARIANT varData, BSTR *retVal)
{
	CXVarPtr varPtr;
	LPWSTR pstr;
	static char HexChar[] = "0123456789ABCDEF";
	UINT i, pos, len1;

	HRESULT hr = varPtr.Attach(varData);
	if(FAILED(hr))return hr;

	i = varPtr.m_nSize * 2;
	pstr = *retVal = ::SysAllocStringLen(NULL, i);

	len1 = 0;
	pos = 0;
	if(*retVal != NULL)
		for(i = 0; i < varPtr.m_nSize; i ++)
		{
			pstr[pos * 2] = HexChar[varPtr.m_pData[i] >> 4];
			pstr[pos * 2 + 1] = HexChar[varPtr.m_pData[i] & 0xf];
			pos ++;
			len1 += 2;
		}

	return S_OK;
}

STDMETHODIMP CXEncoding::XmlEncode(BSTR txtString, BSTR* retVal)
{
	CXAutoPtr<WCHAR> pszEncoded;
	UINT i, len = SysStringLen(txtString);
	int nLength = 0;
	WCHAR ch;

	pszEncoded.Allocate(len * 6 + 1);

	for(i = 0; i < len; i ++)
	{
		switch(ch = txtString[i])
		{
		case '\'':
			wcsncpy_s(pszEncoded + nLength, 7, L"&apos;", 6);
			nLength += 6;
			break;
		case '\"':
			wcsncpy_s(pszEncoded + nLength, 7, L"&quot;", 6);
			nLength += 6;
			break;
		case '<':
			wcsncpy_s(pszEncoded + nLength, 5, L"&lt;", 4);
			nLength += 4;
			break;
		case '>':
			wcsncpy_s(pszEncoded + nLength, 5, L"&gt;", 4);
			nLength += 4;
			break;
		case '&':
			wcsncpy_s(pszEncoded + nLength, 6, L"&amp;", 5);
			nLength += 5;
			break;
		default:
			if(ch == 9 || ch == 0x0A || ch == 0x0D ||
				(ch >= 0x020 && ch <= 0x07E) ||
				(ch >= 0x085 && ch <= 0x0D7FF) ||
				(ch >= 0x0E000 && ch <= 0x0FFFD))
			pszEncoded[nLength ++] = ch;
		}
	}

	*retVal = ::SysAllocStringLen(pszEncoded, nLength);
	return S_OK;
}

STDMETHODIMP CXEncoding::XmlFilter(BSTR txtString, BSTR* retVal)
{
	CXAutoPtr<WCHAR> pszEncoded;
	UINT i, len = SysStringLen(txtString);
	int nLength = 0;
	WCHAR ch;

	pszEncoded.Allocate(len + 1);

	for(i = 0; i < len; i ++)
	{
		ch = txtString[i];
		if(ch == 9 || ch == 0x0A || ch == 0x0D ||
			(ch >= 0x020 && ch <= 0x07E) ||
			(ch >= 0x085 && ch <= 0x0D7FF) ||
			(ch >= 0x0E000 && ch <= 0x0FFFD))
		pszEncoded[nLength ++] = ch;
	}

	*retVal = ::SysAllocStringLen(pszEncoded, nLength);
	return S_OK;
}

typedef struct{char *sName;WORD code;}sHEC;

extern BYTE s_CharWidth[];
extern sHEC s_HtmlEnChar[];

static int compareHtmlEnChar(const void *p1, const void *p2)
{
	return strcmp(*((char**)p1), *((char**)p2));
}

WCHAR getHtmlChar(BSTR& ptr)
{
	BSTR ptr1 = ptr;
	WCHAR ch, ch1;
	char enStr[16], *pKey;
	int n;

	ch = *ptr1++;
	if(ch == '&')
	{
		if(*ptr1 == '#')
		{
			n = 0;
			ptr1 ++;

			while(n < 7)
			{
				ch1 = *ptr1 ++;
				if(ch1 >= '0' && ch1 <= '9')
					enStr[n++] = (char)ch1;
				else if(ch1 == ';')
				{
					enStr[n++] = 0;
					break;
				}else break;
			}

			if(n > 0 && enStr[n - 1] == 0)
			{
				n = atol(enStr);
				if(n > 0 && n < 65536)
				{
					ptr = ptr1;
					ch = n;
				}else ptr ++;
			}else ptr ++;
		}else
		{
			n = 0;

			while(n < 14)
			{
				ch1 = *ptr1 ++;
				if((ch1 >= 'a' && ch1 <= 'z') || (ch1 >= 'A' && ch1 <= 'Z') || (ch1 >= '0' && ch1 <= '9'))
					enStr[n++] = (char)ch1;
				else if(ch1 == ';')
				{
					enStr[n++] = 0;
					break;
				}else break;
			}

			if(n > 0 && enStr[n - 1] == 0)
			{
				pKey = enStr;
				sHEC *p = (sHEC*)bsearch(&pKey, s_HtmlEnChar, 252, sizeof(s_HtmlEnChar[0]), compareHtmlEnChar);

				if(p)
				{
					ptr = ptr1;
					ch = p->code;
				}else ptr ++;
			}else ptr ++;
		}
	}else ptr ++;

	return ch;
}

STDMETHODIMP CXEncoding::FixString(BSTR txtString, short nLen, BSTR* pVal)
{
	BSTR pBase, ptr;
	int l, n;

	nLen *= 16;

	l = 0;
	n = 0;
	ptr = txtString;
	pBase = txtString;

	while(*txtString && l < nLen)
	{
		if(nLen - l > 16)
		{
			ptr = txtString;
			n = l;
		}else if(nLen - l > 32)
		{
			ptr = txtString;
			n = l;
		}

		l += s_CharWidth[getHtmlChar(txtString)];
	}

	if(*txtString || l > nLen)
	{
		n = (nLen - n) / 16;
		*pVal = ::SysAllocStringLen(pBase, ptr - pBase + n);
		ptr = *pVal + (ptr - pBase);
		while(n --)*ptr++ = '.';
	}else *pVal = ::SysAllocStringLen(pBase, txtString - pBase);

	return S_OK;
}

STDMETHODIMP CXEncoding::FixStringLen(BSTR txtString, LONG* pVal)
{
	int l = 0;

	while(*txtString)
		l += s_CharWidth[getHtmlChar(txtString)];

	*pVal = l / 16;

	return S_OK;
}

STDMETHODIMP CXEncoding::HtmlDecode(BSTR HtmlString, BSTR* retVal)
{
	BSTR ptr = HtmlString;
	int n = 0;

	while(*HtmlString)
		ptr[n++] = getHtmlChar(HtmlString);

	*retVal = ::SysAllocStringLen(ptr, n);

	return S_OK;
}

STDMETHODIMP CXEncoding::FindString(VARIANT varStart, BSTR txtString, VARIANT varFind, LONG* pVal)
{
	CComVariant var;
	VARIANT* pvar;
	BSTR sString, sFind;
	int nStart = 0, len;
	HRESULT hr;

	if(varFind.vt == VT_ERROR)
	{
		pvar = &varStart;
		while(pvar && pvar->vt == (VT_VARIANT | VT_BYREF))
			pvar = pvar->pvarVal;

		if(pvar->vt != VT_BSTR)
		{
			hr = VariantChangeType(&var, pvar, 0, VT_BSTR);
			if(FAILED(hr))return hr;

			pvar = &var;
		}

		sString = pvar->bstrVal;
		sFind = txtString;
	}else
	{
		nStart = varGetNumbar(varStart, 1) - 1;
		sString = txtString;
		len = ::SysStringLen(sString);

		if(nStart < 0)nStart = 0;
		if(nStart > len)
			nStart = len;

		pvar = &varFind;
		while(pvar && pvar->vt == (VT_VARIANT | VT_BYREF))
			pvar = pvar->pvarVal;

		if(pvar->vt != VT_BSTR)
		{
			hr = VariantChangeType(&var, pvar, 0, VT_BSTR);
			if(FAILED(hr))return hr;

			pvar = &var;
		}

		sFind = pvar->bstrVal;
	}

	LPCWSTR ptr = wcsstr(sString + nStart, sFind);
	long n = 0;

	if(ptr)n = ptr - sString + 1;

	*pVal = n;
	return S_OK;
}

STDMETHODIMP CXEncoding::FormatMessage(BSTR bstrFmtString, 
										VARIANT bstr1, VARIANT bstr2, VARIANT bstr3, VARIANT bstr4, VARIANT bstr5, VARIANT bstr6, VARIANT bstr7, VARIANT bstr8, VARIANT bstr9,
										VARIANT bstra, VARIANT bstrb, VARIANT bstrc, VARIANT bstrd, VARIANT bstre, VARIANT bstrf, VARIANT bstrg, VARIANT bstrh,
										VARIANT bstri, VARIANT bstrj, VARIANT bstrk, VARIANT bstrl, VARIANT bstrm, VARIANT bstrn, VARIANT bstro, VARIANT bstrp,
										VARIANT bstrq, VARIANT bstrr, VARIANT bstrs, VARIANT bstrt, VARIANT bstru, VARIANT bstrv, VARIANT bstrw, VARIANT bstrx,
										VARIANT bstry, VARIANT bstrz,
										BSTR *bstrMessage)
{
	CXString strResult;
	CXString strOpts[35];
	int i;
	VARIANT *pOpts = &bstr1;
	HRESULT hr;
	LPCWSTR ptr, ptr1, ptrEnd;

	for(i = 0; i < 35; i ++)
	{
		hr = varGetString(pOpts ++, strOpts[i]);
		if(FAILED(hr))return hr;
	}

	ptr = bstrFmtString;
	ptrEnd = ptr + ::SysStringLen(bstrFmtString);

	while(ptr < ptrEnd)
	{
		ptr1 = ptr;
		while(ptr1 < ptrEnd && *ptr1 != '%')
			ptr1 ++;

		strResult.Append(ptr, ptr1 - ptr);
		ptr = ptr1 + 1;

		if(ptr < ptrEnd)
		{
			int ch = *ptr ++;

			if(ch >= '1' && ch <= '9')
				strResult.Append(strOpts[ch - '1']);
			else if(ch >= 'a' && ch <= 'z')
				strResult.Append(strOpts[ch - 'a' + 9]);
			else if(ch == '%')
				strResult.AppendChar('%');
		}
	}

	*bstrMessage = strResult.AllocSysString();

	return S_OK;
}
