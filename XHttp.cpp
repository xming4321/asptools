// XHttp.cpp : Implementation of CXHttp

#include "stdafx.h"
#include "XHttp.h"
#include "atlib.h"
#include <atlsync.h>

// CXHttp

void CXHttp::putConnection(LPSTR strHost, SOCKET s)
{
	POSITION pos;
	int nCount = 0;

	m_cs.Enter();

	pos = m_mapConnPool.FindFirstWithKey(strHost);

	while(pos)
	{
		nCount ++;
		m_mapConnPool.GetNextWithKey(pos, strHost);
	}

	if(nCount < 8)
	{
		m_mapConnPool.Insert(strHost, s);
		s = 0;
	}

	m_cs.Leave();

	if(s)
		closesocket(s);
}

SOCKET CXHttp::getConnection(LPSTR strHost)
{
	SOCKET s = 0;
	POSITION pos = 0;
	CRBMultiMap<CStringA, SOCKET>::CPair* pPair = NULL;
	SOCKADDR_IN addr;

	m_cs.Enter();

	pos = m_mapConnPool.FindFirstWithKey(strHost);
	if(pos)
	{
		pPair = (CRBMultiMap<CStringA, SOCKET>::CPair*)pos;
		s = pPair->m_value;
		m_mapConnPool.RemoveAt(pos);
	}

	m_cs.Leave();

	if(s)return s;

	ZeroMemory(&addr, sizeof(SOCKADDR_IN));

	addr.sin_family = AF_INET;
	addr.sin_port = htons(80);

	addr.sin_addr.s_addr = inet_addr(strHost);

	if (addr.sin_addr.s_addr == INADDR_NONE)
	{
		LPHOSTENT lphost;

		lphost = gethostbyname(strHost);
		if (lphost == NULL)
			return 0;
		addr.sin_addr.s_addr = ((LPIN_ADDR)lphost->h_addr)->s_addr;
	}

	s = socket(AF_INET, SOCK_STREAM, 0);
	if(s == SOCKET_ERROR)
		return 0;

	if(connect(s,(struct sockaddr*)&addr, sizeof(addr)) == SOCKET_ERROR)
	{
		closesocket(s);
		return 0;
	}

	return s;
}

STDMETHODIMP CXHttp::Get(BSTR url, VARIANT varPost, BSTR* retVal)
{
	if(_wcsnicmp(url, L"http://", 7))
		return CXObject::SetErrorInfo(L"Not HTTP protocol.");

	LPSTR pstr, pstr1, strHost;
	CStringA strUrl(url), strPost, strData;
	SOCKET s;
	int n, n1;
	bool bKeepAlive = true, bLoop = true;
	int nReConnect = 2;
	int nStep = 0, nPos = 0, nLinePos = 0, nContentLength = 0x7fffffff;
	char buf[4096];

	strHost = pstr1 = strUrl.GetBuffer() + 7;

	while(*pstr1 && *pstr1 != '/')
		pstr1 ++;

	if(!*pstr1 || strHost == pstr1)
		return CXObject::SetErrorInfo(L"No HTTP host.");

	*pstr1 = 0;

	varGetString(varPost, strData);
	if(strData.IsEmpty())
		strPost = "GET /";
	else strPost = "POST /";
	strPost += pstr1 + 1;
	strPost += " HTTP/1.0\r\nHost: ";
	strPost += strHost;
	strPost += "\r\nConnection: Keep-Alive\r\n";
	if(strData.IsEmpty())
		strPost += "\r\n";
	else
	{
		strPost.AppendFormat("Content-Length: %d\r\n\r\n", strData.GetLength());
		strPost += strData;
	}

	pstr = strPost.GetBuffer();
	n = strPost.GetLength();

	s = getConnection(strHost);
	if(!s)
		return CXObject::GetLastError();

	while(n > 0)
	{
		n1 = send(s, pstr, n, 0);
		if(n1 <= 0)
		{
			closesocket(s);

			if(nReConnect >= 0 && n == strPost.GetLength())
			{
				nReConnect --;

				s = getConnection(strHost);
				if(!s)
					return CXObject::GetLastError();
			}else
				return CXObject::GetLastError();
		}else
		{
			n -= n1;
			pstr += n1;
		}
	}

	strPost.Empty();
	strData.Empty();

	while(nContentLength)
	{
		n = nContentLength;
		if(n > sizeof(buf))
			n = sizeof(buf);

		n1 = recv(s, buf, n, 0);
		if(n1 <= 0)
			break;

		strData.Append(buf, n1);

		bLoop = true;
		while(bLoop && nPos < strData.GetLength())
		{
			switch(nStep)
			{
			case 0:	//  Check "HTTP/1.1 2"
				if(strData.GetLength() - nPos < 10)
				{
					bLoop = false;
					break;
				}

				if(memcmp((LPCSTR)strData + nPos, "HTTP/1.", 7))
				{
					closesocket(s);
					return S_OK;
				}

				if(strData[nPos + 7] < '0' ||  strData[nPos + 7] > '9')
				{
					closesocket(s);
					return S_OK;
				}

				if(strData[nPos + 7] == '0')
					bKeepAlive = false;
				else bKeepAlive = true;

				if(strData[nPos + 8] != ' ')
				{
					closesocket(s);
					return S_OK;
				}

				if(strData[nPos + 9] < '0' || strData[nPos + 9] > '9')
				{
					closesocket(s);
					return S_OK;
				}

				nPos += 10;

				nStep = 1;
			case 1: // new line
				while(nPos < strData.GetLength() && strData[nPos] != '\n')
					nPos ++;

				if(nPos == strData.GetLength())
					break;

				if(nPos == nLinePos || strData[nLinePos] == '\r')
				{
					nPos ++;
					nLinePos = nPos;

					nStep = 3;
					break;
				}

				nStep = 2;
			case 2: // check header: Connection, Content-Length
				if(!_strnicmp((LPCSTR)strData + nLinePos, "Connection:", 11))
				{
					nLinePos += 11;
					while(strData[nLinePos] == ' ')
						nLinePos ++;

					if(!_strnicmp((LPCSTR)strData + nLinePos, "Close", 5))
						bKeepAlive = false;
					else if(!_strnicmp((LPCSTR)strData + nLinePos, "Keep-Alive", 10))
						bKeepAlive = true;
				}else if(!_strnicmp((LPCSTR)strData + nLinePos, "Content-Length:", 15))
				{
					nLinePos += 15;
					while(strData[nLinePos] == ' ')
						nLinePos ++;

					nContentLength = atoi((LPCSTR)strData + nLinePos);
				}

				nPos ++;
				nLinePos = nPos;
				nStep = 1;
				break;

			case 3: // get data
				nContentLength -= strData.GetLength() - nPos;
				nPos = strData.GetLength();
				break;
			}
		}
	}

	if(bKeepAlive)
		putConnection(strHost, s);
	else
		closesocket(s);

	*retVal = strData.Mid(nLinePos).AllocSysString();

	return S_OK;
}
