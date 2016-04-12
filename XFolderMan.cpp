// XFolderMan.cpp : Implementation of CXFolderMan

#include "stdafx.h"
#include "XFolderMan.h"
#include <httpfilt.h>
#include <iphlpapi.h>
#include "XEncoding.h"

CCriticalSection s_csFolders;
static CXFolderMan::CXFolderMaps s_mapDocFolders, s_mapUserFolders, s_mapBoardFolders, s_mapDataFolders;
static CAtlMap< CXString, CXFolderMan::CXFolderMaps*> s_mapFolders;
static CRBMap<long, CXComPtr<CXFolderMan::CXChannel> > s_mapID2Channel;
static CRBMap<CXStringA, CXComPtr<CXFolderMan::CXChannel> > s_mapPath2Channel;
static CRBMap<CXStringA, CXComPtr<CXFolderMan::CXChannel> > s_mapDomain2Channel;
char strExpires[] = "\r\nExpires: Thu, 01 Dec 2053 16:00:00 GMT\r\n";
static BOOL s_bISAPI;

// CXFolderMan


STDMETHODIMP CXFolderMan::CXChannel::get_Item(VARIANT key, VARIANT* pVal)
{
	HRESULT hr;
	CXComPtr<CXChannel> pChannel;

	pChannel = this;

	while(pChannel && ((hr = pChannel->m_pContents->get_Item(key, pVal)) == S_FALSE))
	{
		CXStringA strTemp;
		int nPos;

		if((nPos = pChannel->m_strPath.ReverseFind('/')) != -1)
			strTemp = pChannel->m_strPath.Left(nPos);
		else break;

		pChannel.Release();

		CXAutoLock l(s_csFolders, TRUE);
		s_mapPath2Channel.Lookup(strTemp, pChannel);
	}

	return hr;
}

STDMETHODIMP CXFolderMan::ClearFolder(void)
{
	CXAutoLock l(s_csFolders, TRUE);
	s_mapDocFolders.Clear();
	s_mapUserFolders.Clear();
	s_mapBoardFolders.Clear();

	return S_OK;
}

STDMETHODIMP CXFolderMan::newBand(BSTR sType, LONG nBegin, LONG nEnd, IXPathBand** pVal)
{

	return S_OK;
}

STDMETHODIMP CXFolderMan::GetBand(BSTR sType, LONG nBegin, LONG nEnd, IXPathBand** pVal)
{

	return S_OK;
}

STDMETHODIMP CXFolderMan::GetPath(BSTR sType, LONG i, BSTR* pVal)
{

	return S_OK;
}

STDMETHODIMP CXFolderMan::newDocBand(LONG nBegin, LONG nEnd, IXPathBand** pVal)
{
	return s_mapDocFolders.newDocBand(nBegin, nEnd, pVal);
}

STDMETHODIMP CXFolderMan::newUserBand(LONG nBegin, LONG nEnd, IXPathBand** pVal)
{
	return s_mapUserFolders.newDocBand(nBegin, nEnd, pVal);
}

STDMETHODIMP CXFolderMan::newBoardBand(LONG nBegin, LONG nEnd, IXPathBand** pVal)
{
	return s_mapBoardFolders.newDocBand(nBegin, nEnd, pVal);
}

STDMETHODIMP CXFolderMan::newDataBand(LONG nBegin, LONG nEnd, IXPathBand** pVal)
{
	return s_mapDataFolders.newDocBand(nBegin, nEnd, pVal);
}

STDMETHODIMP CXFolderMan::GetDocBand(LONG nBegin, LONG nEnd, IXPathBand** pVal)
{
	return s_mapDocFolders.GetDocBand(nBegin, nEnd, pVal);
}

STDMETHODIMP CXFolderMan::GetUserBand(LONG nBegin, LONG nEnd, IXPathBand** pVal)
{
	return s_mapUserFolders.GetDocBand(nBegin, nEnd, pVal);
}

STDMETHODIMP CXFolderMan::GetBoardBand(LONG nBegin, LONG nEnd, IXPathBand** pVal)
{
	return s_mapBoardFolders.GetDocBand(nBegin, nEnd, pVal);
}

STDMETHODIMP CXFolderMan::GetDataBand(LONG nBegin, LONG nEnd, IXPathBand** pVal)
{
	return s_mapDataFolders.GetDocBand(nBegin, nEnd, pVal);
}

STDMETHODIMP CXFolderMan::GetDocPath(LONG i, BSTR* pVal)
{
	return s_mapDocFolders.GetFullPath(i, pVal);
}

STDMETHODIMP CXFolderMan::GetUserPath(LONG i, BSTR* pVal)
{
	return s_mapUserFolders.GetFullPath(i, pVal);
}

STDMETHODIMP CXFolderMan::GetBoardPath(LONG i, BSTR* pVal)
{
	return s_mapBoardFolders.GetFullPath(i, pVal);
}

STDMETHODIMP CXFolderMan::GetDataPath(LONG i, BSTR* pVal)
{
	return s_mapDataFolders.GetFullPath(i, pVal);
}

STDMETHODIMP CXFolderMan::ClearChannel(void)
{
	CXAutoLock l(s_csFolders, TRUE);
	s_mapID2Channel.RemoveAll();
	s_mapPath2Channel.RemoveAll();
	s_mapDomain2Channel.RemoveAll();

	return S_OK;
}

void CXFolderMan::AddOneChannel(LONG id, LPCWSTR path, LPCWSTR domain, LPCWSTR homepage, LPCWSTR mark, LPCWSTR folder, CXChannel** pVal)
{
	CXComPtr<CXChannel> pChannel;
	CXComPtr<CXChannel> pParentChannel;
	CXStringA parentPath;
	int p;

	pChannel.CreateObject();

	pChannel->m_strPath = path;
	pChannel->m_strPath.MakeLower();
	pChannel->m_strPath.Replace('\\', '/');
	while(!pChannel->m_strPath.IsEmpty() && pChannel->m_strPath[pChannel->m_strPath.GetLength() - 1] == '/')
		pChannel->m_strPath = pChannel->m_strPath.Left(pChannel->m_strPath.GetLength() - 1);

	parentPath = pChannel->m_strPath;
	p = parentPath.ReverseFind('/');
	if(p != -1)
		parentPath = parentPath.Left(p);
	else parentPath.Empty();

	pChannel->m_strDomain = domain;
	pChannel->m_strDomain.MakeLower();

	pChannel->m_strHomePage = homepage;
	pChannel->m_strHomePage.Replace('\\', '/');

	pChannel->m_strMark = mark;
	if(pChannel->m_strMark.IsEmpty())pChannel->m_strMark = "a";

	pChannel->m_strFolder = folder;
	pChannel->m_strFolder.MakeLower();
	pChannel->m_strFolder.Replace('\\', '/');
	while(!pChannel->m_strFolder.IsEmpty() && pChannel->m_strFolder[pChannel->m_strFolder.GetLength() - 1] == '/')
		pChannel->m_strFolder = pChannel->m_strFolder.Left(pChannel->m_strFolder.GetLength() - 1);

	pChannel->m_nID = id;

	s_mapID2Channel.SetAt(id, pChannel);
	s_mapPath2Channel.SetAt(pChannel->m_strPath, pChannel);
	s_mapDomain2Channel.SetAt(pChannel->m_strDomain, pChannel);

	if(s_mapPath2Channel.Lookup(parentPath, pParentChannel) && (pParentChannel != pChannel))
	{
		pParentChannel->m_pSubs->AddValue(pChannel);
		pChannel->m_pParentChannel = pParentChannel;
		if(pChannel->m_strFolder.IsEmpty())
			pChannel->m_strFolder = pParentChannel->m_strFolder;
	}

	if(pChannel->m_strFolder.IsEmpty())
		pChannel->m_strFolder = "/club";

	if(pVal)
		*pVal = pChannel.Detach();
}

STDMETHODIMP CXFolderMan::AddChannel(LONG id, BSTR path, BSTR domain, BSTR homepage, BSTR mark, BSTR folder, IXChannel** pVal)
{
	s_csFolders.Enter();
	AddOneChannel(id, path, domain, homepage, mark, folder, (CXChannel**)pVal);
	s_csFolders.Leave();

	return S_OK;
}

void inline getSplitString(BSTR& batchText, CXString& str)
{
	BSTR strTemp = batchText;
	WCHAR ch = 0;

	while((ch = *batchText) && (ch != '\t') && (ch != '\n') && (ch != '\r'))
		batchText ++;

	str.SetString(strTemp, batchText - strTemp);

	if(ch == '\t')batchText ++;
}

STDMETHODIMP CXFolderMan::BatchAddChannel(BSTR batchText)
{
	HRESULT hr = S_OK;
	CXString path, domain, homepage, mark, folder;
	CXString str;
	int id;
	CXComPtr<CXChannel> pChannel;
	CComVariant k, v;

	s_csFolders.Enter();
	while(*batchText)
	{
		while((*batchText == ' ') || (*batchText == '\t') || (*batchText == '\n') || (*batchText == '\r'))
			batchText ++;

		if(!*batchText)break;

		getSplitString(batchText, str);
		id = _wtol(str);

		getSplitString(batchText, path);
		getSplitString(batchText, domain);
		getSplitString(batchText, homepage);
		getSplitString(batchText, mark);
		getSplitString(batchText, folder);

		AddOneChannel(id, path, domain, homepage, mark, folder, &pChannel);

		while(1)
		{
			getSplitString(batchText, str);
			if(str.IsEmpty())break;
			k = str;
			getSplitString(batchText, str);
			v = str;

			pChannel->put_Item(k, v);
		}

		while(*batchText && (*batchText != '\n') && (*batchText != '\r'))
			batchText ++;
	}
	s_csFolders.Leave();

	return S_OK;
}

STDMETHODIMP CXFolderMan::GetChannel(LONG id, IXChannel** pVal)
{
	CXComPtr<CXChannel> pChannel;

	CXAutoLock l(s_csFolders, TRUE);
	s_mapID2Channel.Lookup(id, pChannel);

	return pChannel.QueryInterface(pVal);
}


STDMETHODIMP CXFolderMan::GetChannelFromDomain(BSTR strDomain, IXChannel** pVal)
{
	CXComPtr<CXChannel> pChannel;
	CXStringA strTemp(strDomain, ::SysStringLen(strDomain));

	CXAutoLock l(s_csFolders, TRUE);
	s_mapDomain2Channel.Lookup(strTemp, pChannel);

	return pChannel.QueryInterface(pVal);
}

STDMETHODIMP CXFolderMan::GetFile(BSTR FilePath, IXFile** ppfile)
{
	HRESULT hr;
	CXComPtr<CXFileInfo> pFile;

	pFile.CreateObject();

	hr = pFile->GetFileInfo(FilePath);
	if(FAILED(hr))return hr;

	*ppfile = pFile.Detach();

	return S_OK;
}

STDMETHODIMP CXFolderMan::GetFiles(BSTR FilePath, IXList** ppfiles)
{
	CXComPtr< CXList< CXComPtr<CXFileInfo> > > pFiles;
	CRBMap<CXKeyString, int> mapItems;
	CXPath path(FilePath);

	pFiles.CreateObject();

	WIN32_FIND_DATAW fd;
	HANDLE hd;
	HRESULT hr;

	hd = ::FindFirstFileW(FilePath, &fd);
	if(hd == INVALID_HANDLE_VALUE)
	{
		hr = GetLastError();
		if(SUCCEEDED(hr))
			*ppfiles = pFiles.Detach();
		return hr;
	}

	BOOL bFindNext = TRUE;

	path.StripPath();

	while(bFindNext)
	{
		CXString strName(fd.cFileName);

		if(!mapItems.Lookup(strName))
		{
			mapItems.SetAt(strName, 0);
			if(!(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
			{
				CXComPtr<CXFileInfo> pFolder;

				pFolder.CreateObject();
				pFolder->FillData(&fd);
				pFolder->m_strPath = path + strName;

				pFiles->AddValue(pFolder);
			}
		}

		bFindNext = FindNextFileW(hd, &fd);
	}

	::FindClose(hd);

	*ppfiles = pFiles.Detach();

	return S_OK;
}

STDMETHODIMP CXFolderMan::FileExists(BSTR fn, VARIANT_BOOL* pVal)
{
	*pVal = PathFileExistsW(fn) ? VARIANT_TRUE : VARIANT_FALSE;

	return S_OK;
}

inline HRESULT delFile(LPCWSTR FileSpec, BOOL bForce)
{
	HRESULT hr = S_OK;

	if(!::DeleteFileW(FileSpec))
	{
		hr = CXObject::GetLastError();

		if(hr == E_ACCESSDENIED && bForce)
		{
			if(!SetFileAttributesW(FileSpec, 0))
				return CXObject::GetLastError();

			if(!::DeleteFileW(FileSpec))
				return CXObject::GetLastError();

			hr = S_OK;
		}
	}

	return hr;
}

STDMETHODIMP CXFolderMan::DeleteFile(BSTR FileSpec, VARIANT bForce, LONG* pVal)
{
	WIN32_FIND_DATAW FindFileData;
	HANDLE hFind = INVALID_HANDLE_VALUE;
	HRESULT hr = S_OK;
	BOOL bForce1 = varGetNumbar(bForce);
	CXString strPath(FileSpec, ::SysStringLen(FileSpec));
	long n;

	if(pVal == NULL)pVal = &n;
	*pVal = 0;

	strPath.Replace('/', '\\');

	hFind = FindFirstFileW(strPath, &FindFileData);

	if(hFind != INVALID_HANDLE_VALUE)
	{
		strPath = strPath.Left(strPath.ReverseFind('\\') + 1);
		do
		{
			if(!(FindFileData.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
			{
				hr = delFile(strPath + FindFileData.cFileName, bForce1);
				if(FAILED(hr))
				{
					FindClose(hFind);
					return hr;
				}
				*pVal ++;
			}
		}while (FindNextFileW(hFind, &FindFileData) != 0);

		FindClose(hFind);
	}

	return hr;
}

STDMETHODIMP CXFolderMan::MoveFile(BSTR Source, BSTR Destination)
{
	if(!::MoveFileW(Source, Destination))
		return CXObject::GetLastError();

	return S_OK;
}

STDMETHODIMP CXFolderMan::CopyFile(BSTR Source, BSTR Destination, VARIANT bOverwrite)
{
	if(!::CopyFileW(Source, Destination, varGetNumbar(bOverwrite)))
		return CXObject::GetLastError();

	return S_OK;
}

STDMETHODIMP CXFolderMan::CreateFolder(BSTR FileSpec)
{
	if(!::CreateDirectoryW(FileSpec, NULL))
		return CXObject::GetLastError();

	return S_OK;
}

STDMETHODIMP CXFolderMan::DeleteFolder(BSTR FileSpec)
{
	if(!::RemoveDirectoryW(FileSpec))
		return CXObject::GetLastError();

	return S_OK;
}

STDMETHODIMP CXFolderMan::CountConnection(BSTR Addr, LONG port, LONG* pVal)
{
	DWORD addr;
	PMIB_TCPTABLE pTcpTable = NULL;
	DWORD dwSize = 0;
	LONG nCount = 0;
	BOOL bClient = FALSE;

	if(port < 0)
	{
		bClient = TRUE;
		port = -port;
	}

	addr = inet_addr(CXStringA(Addr));
	port = htons((u_short)port);

	if (GetTcpTable(pTcpTable, &dwSize, TRUE) == ERROR_INSUFFICIENT_BUFFER)
		pTcpTable = (MIB_TCPTABLE*) malloc ((UINT) dwSize);

	if (GetTcpTable(pTcpTable, &dwSize, TRUE) == NO_ERROR)
		for (int i = 0; i < (int) pTcpTable->dwNumEntries; i++)
		{
			if(pTcpTable->table[i].dwLocalAddr == 0 || pTcpTable->table[i].dwLocalAddr == 0x0100007f)
				continue;

			if(addr != 0xffffffff)
			{
				if( !bClient && pTcpTable->table[i].dwLocalAddr != addr)
					continue;

				if( bClient && pTcpTable->table[i].dwRemoteAddr != addr)
					continue;
			}

			if(port && pTcpTable->table[i].dwLocalPort != port)
				continue;

			nCount ++;
		}

	free(pTcpTable);

	*pVal = nCount;

	return S_OK;
}

void CXFolderMan::ParseUrl(LPCSTR path, CXUrl* pVal)
{
	int i, nLen, nLen1;
	CXStringA strUrl, strClubUrl;
	CXComPtr<CXFolderMan::CXChannel> pChannel;
	char ch;

	if(!_strnicmp(path, "http://", 7))
		path = strchr(path + 7, '/');

	if(!path || *path != '/')path = "/";

	nLen = 0;
	while((ch = path[nLen]) && ((ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9') || ch == '/' || ch == '-' || ch == '_'))
		nLen ++;

	nLen1 = nLen;

	CXAutoLock l(s_csFolders, TRUE);

	do
	{
		strClubUrl.SetString(path, nLen);
        if(s_mapPath2Channel.Lookup(strClubUrl, pChannel))
		{
			char ch = pChannel->m_strMark[0];

			if(ch == 'z')
				pVal->m_nCityID = pChannel->m_nID;
			else pVal->m_nClubID = pChannel->m_nID;

			pVal->m_strFolder = pChannel->m_strFolder;
			path += nLen;
			nLen1 -= nLen;
			nLen = nLen1;

			if(path[0] == '/' && path[1] == 0)
			{
				if(!pChannel->m_strHomePage.IsEmpty())
					pVal->m_strScript = pChannel->m_strHomePage;
				else if(ch == 'a')
					pVal->m_strScript = pChannel->m_strFolder + "/home.asp";
				else
				{
					pVal->m_strScript = pChannel->m_strFolder + "/home_";
					pVal->m_strScript += ch;
					pVal->m_strScript += ".asp";
				}
				pVal->m_strPath = "/";
				return;
			}

			if(pVal->m_nClubID)
				break;

			pChannel.Release();
		}
		else
		{
			nLen --;
			while(nLen && path[nLen] != '/')nLen--;
		}
	}while(nLen);
	l.Leave();

	nLen = nLen1;
	while(path[nLen])
		nLen ++;

	if(pVal->m_nCityID == 0 && pVal->m_nClubID == 0)
	{
		if(nLen > 2 && (path[1] == 'b' || path[1] == 'u' || path[1] == 's'))
		{
			for(i = 2; i < nLen && isNumber(path[i]); i ++);

			if(path[i] == '/' || path[i] == 0)
			{
				switch(path[1])
				{
				case 'b':
					pVal->m_nBoardID = atoi(path + 2);
					pVal->m_nUserID = 0;
					pVal->m_nSubjectID = 0;
					break;
				case 'u':
					pVal->m_nUserID = atoi(path + 2);
					pVal->m_nBoardID = 0;
					pVal->m_nSubjectID = 0;
					break;
				case 's':
					pVal->m_nSubjectID = atoi(path + 2);
					pVal->m_nBoardID = 0;
					pVal->m_nUserID = 0;
					break;
				}

				path += i;
				nLen -= i;

				if(nLen == 1)
				{
					if(pVal->m_nBoardID)
						pVal->m_strScript = "/board/home.asp";
					else if(pVal->m_nUserID)
						pVal->m_strScript = "/blog/home.asp";
					else
						pVal->m_strScript = "/subject/home.asp";

					pVal->m_strPath = "/";
					return;
				}
			}
		}
	}

	if(pVal->m_nBoardID > 0 || pVal->m_nUserID > 0 || pVal->m_nSubjectID > 0 || pVal->m_nCityID > 0 || pVal->m_nClubID > 0)
	{
		if(nLen > 3 && (path[1] >= 'a' && path[1] <= 'z') && isNumber(path[2]))
		{
			ch = path[1];

			for(i = 2; i < nLen && isNumber(path[i]); i ++);

			if(path[i] == '.')
			{
				pVal->m_nDocID = atoi((LPCSTR)path + 2);

				path += i + 1;
				nLen -= i + 1;

				if(strcmp(path, "htm"))
				{
					for(i = 0; i < nLen && path[i] != '.'; i ++);

					if(path[i] == '.')
					{
						pVal->m_strPageNo.SetString(path, i);

						path += i + 1;
						nLen -= i + 1;
					}
				}

				if(!strcmp(path, "htm"))
				{
					if(ch == 'd')
					{
						if(pVal->m_nBoardID)
							pVal->m_strScript = "/board/doc.asp";
						else if(pVal->m_nUserID)
							pVal->m_strScript = "/blog/doc.asp";
						else if(pVal->m_nSubjectID)
							pVal->m_strScript = "/subject/doc.asp";
						else if((pVal->m_nCityID > 0 || pVal->m_nClubID > 0) && pChannel)
							pVal->m_strScript = pChannel->m_strFolder + "/doc.asp";
					}else
					{
						if(pVal->m_nBoardID)
							pVal->m_strScript = "/board/doc_";
						else if(pVal->m_nUserID)
							pVal->m_strScript = "/blog/doc_";
						else if(pVal->m_nSubjectID)
							pVal->m_strScript = "/subject/doc_";
						else if((pVal->m_nCityID > 0 || pVal->m_nClubID > 0) && pChannel)
							pVal->m_strScript = pChannel->m_strFolder + "/doc_";

						pVal->m_strScript.AppendChar(ch);
						pVal->m_strScript.Append(".asp");
					}

					return;
				}
			}
		}
	}

	if(nLen > 3 && (path[1] >= 'a' && path[1] <= 'z') && path[2] == '_')
	{
		ch = path[1];


		for(i = 2; i < nLen && path[i] != '.'; i ++);

		if(path[i] == '.')
		{
			pVal->m_strTag.SetString(path + 3, i - 3);

			path += i + 1;
			nLen -= i + 1;

			if(strcmp(path, "htm"))
			{
				for(i = 0; i < nLen && path[i] != '.'; i ++);

				if(path[i] == '.')
				{
					pVal->m_strPageNo.SetString(path, i);

					path += i + 1;
					nLen -= i + 1;
				}
			}

			if(!strcmp(path, "htm"))
			{
				if(pVal->m_nBoardID)
					pVal->m_strScript = "/board/tag_";
				else if(pVal->m_nUserID)
					pVal->m_strScript = "/blog/tag_";
				else if(pVal->m_nSubjectID)
					pVal->m_strScript = "/subject/tag_";
				else if((pVal->m_nCityID > 0 || pVal->m_nClubID > 0) && pChannel)
					pVal->m_strScript = pChannel->m_strFolder + "/tag_";
				else pVal->m_strScript = "/tag_";

				pVal->m_strScript.AppendChar(ch);
				pVal->m_strScript.Append(".asp");
			}
		}

	}

	pVal->m_strPath = path;

	return;
}

HRESULT CXFolderMan::isISAPIInstalled(VARIANT_BOOL* pVal)
{
	*pVal = s_bISAPI ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}

STDMETHODIMP CXFolderMan::ParseUrl(BSTR path, IXUrl** pVal)
{
	CXStringA strPath(path, ::SysStringLen(path));
	CXComPtr<CXUrl> pUrl;

	strPath.MakeLower();
	pUrl.CreateObject();
	ParseUrl(strPath, pUrl);

	*pVal = pUrl.Detach();

	return S_OK;
}

void inline putStr(char *ptrBuf, const char *ptr, int n)
{
	while(n --)*ptrBuf ++ = *ptr ++;
}

void inline putInt(char *ptrBuf, int v, int n)
{
	while(n --)
	{
		ptrBuf[n] = (v % 10) + '0';
		v /= 10;
	}
}

void genExpiresString()
{
	SYSTEMTIME st;
	FILETIME d;
	__int64 t;
	static __int64 ot;
	static char szMonth[][4] =
	{
		"Jan", "Feb", "Mar", "Apr", "May", "Jun",
		"Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
	};
	static char szDays[][4] =
	{
		"Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"
	};

	GetSystemTime(&st);
	SystemTimeToFileTime(&st, &d);

	t = *(__int64*)&d / 10000000 - 9435312000;

	if(t == ot)return;
	ot = t;

	t += 60 * 60;

	*(__int64*)&d = (t + 9435312000) * 10000000;
	FileTimeToSystemTime(&d, &st);

	putStr(strExpires + 11, szDays[st.wDayOfWeek], 3);
	putInt(strExpires + 16, st.wDay, 2);
	putStr(strExpires + 19, szMonth[st.wMonth-1], 3);
	putInt(strExpires + 23, st.wYear, 4);
	putInt(strExpires + 28, st.wHour, 2);
	putInt(strExpires + 31, st.wMinute, 2);
	putInt(strExpires + 34, st.wSecond, 2);
}

extern "C" BOOL WINAPI __stdcall GetFilterVersion(HTTP_FILTER_VERSION *pVer)
{
	s_bISAPI = TRUE;

	pVer->dwFlags = SF_NOTIFY_PREPROC_HEADERS | SF_NOTIFY_URL_MAP | SF_NOTIFY_SEND_RESPONSE;

	pVer->dwFilterVersion = HTTP_FILTER_REVISION;

	strcpy_s(pVer->lpszFilterDesc, sizeof(pVer->lpszFilterDesc), "XICI url hook, Version 1.0");

	return TRUE;
}

extern "C" DWORD WINAPI __stdcall HttpFilterProc(HTTP_FILTER_CONTEXT *pfc, DWORD NotificationType, VOID *pvData)
{
	s_bISAPI = TRUE;

	if(NotificationType == SF_NOTIFY_SEND_RESPONSE)
	{
		HTTP_FILTER_SEND_RESPONSE* pData = (HTTP_FILTER_SEND_RESPONSE*)pvData;

		if(pData->HttpStatus == 404)
		{
			DWORD dwSize;
			CXStringA strUrl;

			dwSize = 1024;
			if(pData->GetHeader(pfc, "Location:", strUrl.GetBuffer(dwSize), &dwSize))
			{
				strUrl.ReleaseBuffer(dwSize - 1);

				pfc->ServerSupportFunction(pfc, SF_REQ_SEND_RESPONSE_HEADER, (PVOID) "302 Redirect", (DWORD)(LPCSTR)("Location: " + strUrl + "\r\n\r\n"), 0); 
				return SF_STATUS_REQ_FINISHED;
			}
		}else if(pData->HttpStatus >= 500)
		{
			DWORD dwSize;
			CXStringA strUrl;

			dwSize = 1024;
			if(pData->GetHeader(pfc, "Path:", strUrl.GetBuffer(dwSize), &dwSize))
			{
				strUrl.ReleaseBuffer(dwSize - 1);

				pfc->ServerSupportFunction(pfc, SF_REQ_SEND_RESPONSE_HEADER, (PVOID) "302 Redirect", (DWORD)(LPCSTR)("Location: /service/500.htm?" + strUrl + "\r\n\r\n"), 0); 
				return SF_STATUS_REQ_FINISHED;
			}
		}
	}else if(NotificationType == SF_NOTIFY_URL_MAP)
	{
		HTTP_FILTER_URL_MAP* pData = (HTTP_FILTER_URL_MAP*)pvData;

		if(pData->pszURL[0] == '/' && pData->pszURL[1] == '*')
		{
			const char *ptr = pData->pszURL + 2;
			char *ptr1 = pData->pszPhysicalPath;

			if(pData->pszURL[2] == '/' || pData->pszURL[2] == '\\')
				*ptr1 ++ = '\\';

			char ch;

			while(ch = *ptr ++)
				if(ch == '/')*ptr1 ++ = '\\';
				else if(ch == '*')*ptr1 ++ = '%';
				else *ptr1 ++ = ch;

			*ptr1 ++ = 0;

			strcpy_s(pData->pszPhysicalPath, ptr1 - pData->pszPhysicalPath, CXStringA(CXEncoding::UrlDecode(pData->pszPhysicalPath, ptr1 - pData->pszPhysicalPath)));
		}
	}else if(NotificationType == SF_NOTIFY_PREPROC_HEADERS)
	{
		HTTP_FILTER_PREPROC_HEADERS* pData = (HTTP_FILTER_PREPROC_HEADERS*)pvData;
		CXStringA strHost, strUrl, strFile, strQueryString;
		DWORD dwSize;
		int nPos, i;
		CXFolderMan::CXUrl url;
		char urlBuffer[1024], ch;

		dwSize = sizeof(urlBuffer);
		if(!pData->GetHeader(pfc, "url", urlBuffer, &dwSize))
			return SF_STATUS_REQ_NEXT_NOTIFICATION;

		if(!pData->AddHeader(pfc, "path:", urlBuffer))
			return SF_STATUS_REQ_NEXT_NOTIFICATION;

		nPos = 0;
		while(ch = urlBuffer[nPos])
		{
			if((ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9') || ch == '/' || ch == '-' || ch == '_')ch;
			else if(ch >= 'A' && ch <= 'Z')
				urlBuffer[nPos] = ch + ('a' - 'A');
			else if(ch == '\\')urlBuffer[nPos]='/';
			else break;
			nPos ++;
		}

		while(urlBuffer[nPos] && urlBuffer[nPos] != '?')nPos ++;

		if(urlBuffer[nPos])
		{
			strQueryString = urlBuffer + nPos;
			strUrl.SetString(urlBuffer, nPos);
		}else
		{
			strUrl.SetString(urlBuffer, nPos);
		}

		dwSize = 128;
		if(!pData->GetHeader(pfc, "host:", strHost.GetBuffer(dwSize), &dwSize))
			return SF_STATUS_REQ_NEXT_NOTIFICATION;
		strHost.ReleaseBuffer(dwSize - 1);
		strHost.MakeLower();

		if(strHost.Right(10).Compare(".elong.com"))
		{
			if(strHost.Right(9).Compare(".xici.net"))
			{
				pfc->ServerSupportFunction(pfc, SF_REQ_SEND_RESPONSE_HEADER, (PVOID) "302 Redirect", (DWORD)(LPCSTR)("Location: http://www.xici.net\r\n\r\n"), 0); 
				return SF_STATUS_REQ_FINISHED;
			}

			if(strHost != "www.xici.net" && 
				strHost != "user.xici.net" && 
				strHost != "imgs.xici.net" && 
				strHost != "files.xici.net" && 
				strHost != "mfs.xici.net" && 
				strHost != "icons.xici.net" && 
				strHost != "pics.xici.net" && 
				strHost != "jubo.xici.net" && 
				strHost != "download.xici.net" && 
				strHost != "preview.xici.net")
			{
				pData->SetHeader(pfc, "url", (LPSTR)(LPCSTR)("/domain.asp?" + strHost.Left(strHost.GetLength() - 9)));
				return SF_STATUS_REQ_NEXT_NOTIFICATION;
			}
		}

		CXFolderMan::ParseUrl(strUrl, &url);

		if(url.m_strPath.GetLength() > 3 && url.m_strPath[1] == 'd')
		{
			for(i = 2; i < url.m_strPath.GetLength() && isNumber(url.m_strPath[i]); i ++);

			if(url.m_strPath[i] == '.')
				for(i ++; i < url.m_strPath.GetLength() && isNumber(url.m_strPath[i]); i ++);
			else if(url.m_strPath[i] == '/')
				pfc->AddResponseHeaders(pfc, (LPSTR)(LPCSTR)("Location: /service/" + url.m_strPath.Mid(i + 1) + "\r\n"), 0);

			if(url.m_strPath[i] == '/')
			{
				if(url.m_strPath.Find('/', i + 1) != -1)
					return SF_STATUS_REQ_NEXT_NOTIFICATION;

				strFile.Empty();

				if((nPos = url.m_strPath.ReverseFind('.')) != -1)
					strFile = url.m_strPath.Mid(nPos + 1);

				if(!strFile.Compare("swf"))
					pfc->AddResponseHeaders(pfc, 
						(LPSTR)(LPCSTR)("Content-Type: application/x-shockwave-flash\r\nExpires: Thu, 01 Dec 2050 16:00:00 GMT\r\n"), 0);
				else
				{
					strFile = url.m_strPath.Mid(i + 1);
					pfc->AddResponseHeaders(pfc, 
						(LPSTR)(LPCSTR)("Content-Type: application/download\r\nContent-disposition: attachment;filename=" + CXStringA(CXEncoding::UrlDecode(strFile)) + "\r\nExpires: Thu, 01 Dec 2050 16:00:00 GMT\r\n"), 0);
				}

				strFile = url.m_strPath.Mid(1);

				strFile.Replace('%', '*');
				strFile.Replace('/', '.');
				strUrl = s_mapDocFolders.GetFullPath(atoi((LPCSTR)url.m_strPath + 2));
				strUrl += strFile;
				strUrl.Append(".bin", 4);
				strUrl = "/*" + strUrl;

				pData->SetHeader(pfc, "url", (LPSTR)(LPCSTR)strUrl);
				return SF_STATUS_REQ_NEXT_NOTIFICATION;
			}
		}

		if(url.m_nCityID || url.m_nClubID || url.m_nBoardID || url.m_nUserID || url.m_nSubjectID || url.m_nDocID)
		{
			if(url.m_strScript.IsEmpty())
			{
				if(url.m_strPath.IsEmpty())
				{
					strFile.Format("Location: %s/\r\n\r\n", strUrl);
					if(!strQueryString.IsEmpty())strFile = strFile + strQueryString;
					pfc->ServerSupportFunction (pfc, SF_REQ_SEND_RESPONSE_HEADER, (PVOID) "302 Redirect", (DWORD)(LPCSTR)strFile, 0); 
					return SF_STATUS_REQ_FINISHED;
				}

				if(url.m_nBoardID || url.m_nUserID || url.m_nSubjectID)
				{
					if(!strncmp(url.m_strPath, "/upload.temp/", 13))
					{
						if(url.m_strPath.Find('/', 13) != -1)
							return SF_STATUS_REQ_NEXT_NOTIFICATION;

						strUrl = "/temp_upload.asp?file=";
						strUrl.Append((LPCSTR)url.m_strPath + 13);

						pData->SetHeader(pfc, "url", (LPSTR)(LPCSTR)strUrl);
						return SF_STATUS_REQ_NEXT_NOTIFICATION;
					}else if(!strncmp(url.m_strPath, "/files/", 7))
					{
						if(url.m_strPath.Find('/', 7) != -1)
							return SF_STATUS_REQ_NEXT_NOTIFICATION;

						strFile = url.m_strPath.Mid(7);
						pfc->AddResponseHeaders(pfc, (LPSTR)(LPCSTR)((url.m_nBoardID ? "Location: /board/files/" : (url.m_nUserID ? "Location: /blog/files/" : "Location: /subject/files/")) + strFile + strExpires), 0);

						if(url.m_nBoardID)
							strUrl = s_mapBoardFolders.GetFullPath(url.m_nBoardID);
						else if(url.m_nUserID)
							strUrl = s_mapUserFolders.GetFullPath(url.m_nUserID);
						else strUrl = s_mapBoardFolders.GetFullPath(url.m_nSubjectID);
						strUrl.AppendChar(url.m_nBoardID ? 'b' : (url.m_nUserID ? 'u' : 's'));
						strUrl.AppendFormat("%d.", url.m_nBoardID ? url.m_nBoardID : (url.m_nUserID ? url.m_nUserID : url.m_nSubjectID));
						strUrl.Append(strFile);
						strUrl.Append(".bin", 4);
						strUrl = "/*" + strUrl;

						pData->SetHeader(pfc, "url", (LPSTR)(LPCSTR)strUrl);
						return SF_STATUS_REQ_NEXT_NOTIFICATION;
					}
				}else if((url.m_nCityID || url.m_nClubID) && !strncmp(url.m_strPath, "/special/", 9))
				{
					if(url.m_strPath.Find('/', 9) != -1 || url.m_strPath.Right(4).CompareNoCase(".htm"))
						return SF_STATUS_REQ_NEXT_NOTIFICATION;

					if(strQueryString.IsEmpty())strQueryString.Append("?name=");
					else strQueryString.Append("&name=");
					strQueryString.Append(url.m_strPath.Mid(9, url.m_strPath.GetLength() - 13));
					url.m_strPath = url.m_strFolder + "/special.asp";
				}

				if(url.m_strPath.Find('/', 1) == -1)
					strUrl = (url.m_nBoardID ? "/board" : (url.m_nUserID ? "/blog" : (url.m_nSubjectID ? "/subject" : url.m_strFolder))) + url.m_strPath;
				else strUrl = url.m_strPath;
			}else strUrl = url.m_strScript;

			if(!strUrl.Right(4).CompareNoCase(".asp"))
			{
				if(url.m_nCityID)
					if(strQueryString.IsEmpty())strQueryString.AppendFormat("?city_id=%d", url.m_nCityID);
					else strQueryString.AppendFormat("&city_id=%d", url.m_nCityID);
				if(url.m_nClubID)
					if(strQueryString.IsEmpty())strQueryString.AppendFormat("?club_id=%d", url.m_nClubID);
					else strQueryString.AppendFormat("&club_id=%d", url.m_nClubID);
				if(url.m_nBoardID)
					if(strQueryString.IsEmpty())strQueryString.AppendFormat("?bd_id=%d", url.m_nBoardID);
					else strQueryString.AppendFormat("&bd_id=%d", url.m_nBoardID);
				if(url.m_nUserID)
					if(strQueryString.IsEmpty())strQueryString.AppendFormat("?user_id=%d", url.m_nUserID);
					else strQueryString.AppendFormat("&user_id=%d", url.m_nUserID);
				if(url.m_nSubjectID)
					if(strQueryString.IsEmpty())strQueryString.AppendFormat("?subject_id=%d", url.m_nSubjectID);
					else strQueryString.AppendFormat("&subject_id=%d", url.m_nSubjectID);

				if(url.m_nDocID)
				{
					if(strQueryString.IsEmpty())strQueryString.AppendFormat("?doc_id=%d", url.m_nDocID);
					else strQueryString.AppendFormat("&doc_id=%d", url.m_nDocID);

					if(!url.m_strPageNo.IsEmpty())
					{
						strQueryString.Append("&pn=", 4);
						strQueryString.Append(url.m_strPageNo);
					}
				}else if(!url.m_strTag.IsEmpty())
				{
					if(strQueryString.IsEmpty())strQueryString.Append("?tag=");
					else strQueryString.Append("&tag=");

					strQueryString.Append(url.m_strTag);

					if(!url.m_strPageNo.IsEmpty())
					{
						strQueryString.Append("&pn=", 4);
						strQueryString.Append(url.m_strPageNo);
					}
				}

				strUrl += strQueryString;
			}else pfc->AddResponseHeaders(pfc, strExpires + 2, 0);

			pData->SetHeader(pfc, "url", (LPSTR)(LPCSTR)strUrl);
		}else if(strUrl.GetLength() > 5 && strUrl.Right(4).CompareNoCase(".asp"))
			pfc->AddResponseHeaders(pfc, strExpires + 2, 0);
	}

	return SF_STATUS_REQ_NEXT_NOTIFICATION;
}

CXFolderMan::~CXFolderMan()
{
	s_mapDocFolders.Clear();
	s_mapUserFolders.Clear();
	s_mapBoardFolders.Clear();

	s_mapID2Channel.RemoveAll();
	s_mapPath2Channel.RemoveAll();
	s_mapDomain2Channel.RemoveAll();
}

