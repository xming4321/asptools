// XFolderMan.h : Declaration of the CXFolderMan

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"
#include "atlib.h"
#include "XDictionary.h"
#include "XRecords.h"
#include "XDate.h"
#include "XPath.h"
#include "XList.h"


// CXFolderMan

extern CCriticalSection s_csFolders;

class ATL_NO_VTABLE CXFolderMan : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXFolderMan, &CLSID_XFolderMan>,
	public IDispatchImpl<IXFolderMan, &IID_IXFolderMan, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	class CXChannel : public CXDispatch<IXChannel>
	{
	public:
		CXChannel()
		{
			m_pContents.CoCreateObject();
			m_pSubs.CreateObject();
			m_nRank = 0;
			m_pParentChannel = 0;
		}

	public:
		STDMETHOD(get_Item)(VARIANT key, VARIANT* pVal);

		STDMETHOD(put_Item)(VARIANT key, VARIANT newVal)
		{
			return m_pContents->put_Item(key, newVal);
		}

		STDMETHOD(putref_Item)(VARIANT key, VARIANT newVal)
		{
			return m_pContents->putref_Item(key, newVal);
		}

		STDMETHOD(get_Id)(long *pVal)
		{
			*pVal = m_nID;

			return S_OK;
		}

		STDMETHOD(get_Path)(BSTR *pVal)
		{
			*pVal = m_strPath.AllocSysString();

			return S_OK;
		}

		STDMETHOD(get_Domain)(BSTR *pVal)
		{
			*pVal = m_strDomain.AllocSysString();

			return S_OK;
		}

		STDMETHOD(get_HomePage)(BSTR *pVal)
		{
			*pVal = m_strHomePage.AllocSysString();

			return S_OK;
		}

		STDMETHOD(get_Mark)(BSTR *pVal)
		{
			*pVal = m_strMark.AllocSysString();

			return S_OK;
		}

		STDMETHOD(get_Contents)(IXDictionary** pVal)
		{
			return m_pContents.QueryInterface(pVal);
		}

		STDMETHOD(get_Subs)(IXList** pVal)
		{
			return m_pSubs.QueryInterface(pVal);
		}

		STDMETHOD(get_Parent)(IXChannel** pVal)
		{
			if(m_pParentChannel)
				return m_pParentChannel->QueryInterface(IID_IXChannel, (void**)pVal);
			return S_OK;
		}

	public:
		long m_nID;
		double m_nRank;
		CXStringA m_strPath;
		CXStringA m_strDomain;
		CXStringA m_strHomePage;
		CXStringA m_strFolder;
		CXStringA m_strMark;
		CXComPtr<CXDictionary> m_pContents;
		CXComPtr<CXList< CXComPtr<CXChannel> > > m_pSubs;
		CXChannel* m_pParentChannel;
	};

	class CXClass : public CXDispatch<IXClass>
	{
		STDMETHOD(AddField)(BSTR key, VARIANT type);

	public:
		int m_nIndex;
	};

	class CXUrl : public CXDispatch<IXUrl>
	{
	public:
		CXUrl() :
			m_nCityID(0),
			m_nClubID(0),
			m_nBoardID(0),
			m_nUserID(0),
			m_nSubjectID(0),
            m_nDocID(0)
		{}

	public:
		STDMETHOD(get_CityID)(long* pVal)
		{
			*pVal = m_nCityID;
			return S_OK;
		}

		STDMETHOD(get_ClubID)(long* pVal)
		{
			*pVal = m_nClubID;
			return S_OK;
		}

		STDMETHOD(get_BoardID)(long* pVal)
		{
			*pVal = m_nBoardID;
			return S_OK;
		}

		STDMETHOD(get_UserID)(long* pVal)
		{
			*pVal = m_nUserID;
			return S_OK;
		}

		STDMETHOD(get_SubjectID)(long* pVal)
		{
			*pVal = m_nSubjectID;
			return S_OK;
		}

		STDMETHOD(get_DocID)(long* pVal)
		{
			*pVal = m_nDocID;
			return S_OK;
		}

		STDMETHOD(get_PageNo)(BSTR* pVal)
		{
			*pVal = m_strPageNo.AllocSysString();
			return S_OK;
		}

		STDMETHOD(get_path)(BSTR* pVal)
		{
			*pVal = m_strPath.AllocSysString();
			return S_OK;
		}

		STDMETHOD(get_tag)(BSTR* pVal)
		{
			*pVal = m_strTag.AllocSysString();
			return S_OK;
		}

		STDMETHOD(get_script)(BSTR* pVal)
		{
			*pVal = m_strScript.AllocSysString();
			return S_OK;
		}

	public:
		int m_nCityID;
		int m_nClubID;
		int m_nBoardID;
		int m_nUserID;
		int m_nSubjectID;
		int m_nDocID;
		CXStringA m_strPageNo;
		CXStringA m_strPath;
		CXStringA m_strFolder;
		CXStringA m_strTag;
		CXStringA m_strScript;
	};

	class CXFileInfo : public CXDispatch<IXFile>
	{
	public:
		CXFileInfo(void) : m_nAttr(0), m_fSize(0)
		{}

	public:
		//	IXFile
		STDMETHOD(get_Path)(BSTR *pbstrPath)
		{
			*pbstrPath = m_strPath.AllocSysString();
			return S_OK;
		}

		STDMETHOD(get_Name)(BSTR *pbstrName)
		{
			*pbstrName = m_strName.AllocSysString();
			return S_OK;
		}

		STDMETHOD(get_Attributes)(short *pfa)
		{
			*pfa = m_nAttr;
			return S_OK;
		}

		STDMETHOD(get_DateCreated)(DATE *pdate)
		{
			*pdate = m_dateCreated;
			return S_OK;
		}

		STDMETHOD(get_DateLastModified)(DATE *pdate)
		{
			*pdate = m_dateLastModified;
			return S_OK;
		}

		STDMETHOD(get_DateLastAccessed)(DATE *pdate)
		{
			*pdate = m_dateLastAccessed;
			return S_OK;
		}

		STDMETHOD(get_Size)(DOUBLE *pvarSize)
		{
			*pvarSize = m_fSize;
			return S_OK;
		}

		STDMETHOD(get_Type)(BSTR *pbstrType)
		{
			SHFILEINFOW fi;

			ZeroMemory(&fi,sizeof(fi));
			SHGetFileInfoW(m_strPath, 0, &fi, sizeof(fi), SHGFI_TYPENAME);

			*pbstrType = ::SysAllocString(fi.szTypeName);

			return S_OK;
		}

		STDMETHOD(get_Files)(IXList **ppfiles)
		{
			if(!m_pFiles)
			{
				HRESULT hr;

				hr = GetFolderInfo();
				if(FAILED(hr))return hr;
			}

			return m_pFiles.QueryInterface(ppfiles);
		}

		STDMETHOD(get_SubFolders)(IXList **ppfolders)
		{
			if(!m_pFolders)
			{
				HRESULT hr;

				hr = GetFolderInfo();
				if(FAILED(hr))return hr;
			}

			return m_pFolders.QueryInterface(ppfolders);
		}

	public:
		void FillData(WIN32_FIND_DATAW *pfd)
		{
			m_strName = pfd->cFileName;
			m_nAttr = (BYTE)pfd->dwFileAttributes;

			m_dateCreated = pfd->ftCreationTime;
			m_dateLastModified = pfd->ftLastWriteTime;
			m_dateLastAccessed = pfd->ftLastAccessTime;

			m_fSize = (double)((__int64)pfd->nFileSizeHigh << 32) + pfd->nFileSizeLow;
		}

		HRESULT GetFileInfo(BSTR FilePath)
		{
			m_strPath = FilePath;

			WIN32_FIND_DATAW fd;
			HANDLE hd;

			hd = ::FindFirstFileW(m_strPath, &fd);
			if(hd == INVALID_HANDLE_VALUE)
				return GetLastError();
			::FindClose(hd);

			FillData(&fd);

			return S_OK;
		}

		HRESULT GetFolderInfo()
		{
			if(!(FILE_ATTRIBUTE_DIRECTORY & m_nAttr))
				return E_NOTIMPL;

			CXComPtr< CXList< CXComPtr<CXFileInfo> > > pFiles;
			CXComPtr< CXList< CXComPtr<CXFileInfo> > > pFolders;
			CXPath path(m_strPath);
			CRBMap<CXKeyString, int> mapItems;

			path.AddBackslash();

			pFiles.CreateObject();
			pFolders.CreateObject();

			WIN32_FIND_DATAW fd;
			HANDLE hd;

			hd = ::FindFirstFileW(path.m_strPath + L"*", &fd);
			if(hd == INVALID_HANDLE_VALUE)
				return GetLastError();

			BOOL bFindNext = TRUE;

			while(bFindNext)
			{
				CXString strName(fd.cFileName);

				if(!mapItems.Lookup(strName))
				{
					mapItems.SetAt(strName, 0);
					if(fd.cFileName[0] != '.' || !(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY))
					{
						CXComPtr<CXFileInfo> pFolder;

						pFolder.CreateObject();
						pFolder->FillData(&fd);
						pFolder->m_strPath = path + strName;

						if(fd.dwFileAttributes & FILE_ATTRIBUTE_DIRECTORY)
							pFolders->AddValue(pFolder);
						else
							pFiles->AddValue(pFolder);
					}
				}

				bFindNext = FindNextFileW(hd, &fd);
			}

			::FindClose(hd);

			m_pFiles = pFiles;
			m_pFolders = pFolders;

			return S_OK;
		}

	public:
		CXString m_strPath;
		CXString m_strName;
		short m_nAttr;
		CXDate m_dateCreated;
		CXDate m_dateLastModified;
		CXDate m_dateLastAccessed;
		DOUBLE m_fSize;
		CXComPtr< CXList< CXComPtr<CXFileInfo> > > m_pFiles;
		CXComPtr< CXList< CXComPtr<CXFileInfo> > > m_pFolders;
		HANDLE m_hdNotify;
	};

	class CXPathBand : public CXDispatch<IXPathBand>
	{
	public:
		CXPathBand()
		{
			m_nBegin = 0;
			m_nEnd = 0;
			m_mapFolders.SetCount(1000);
		}

		STDMETHOD(get_Item)(LONG i, BSTR *pVal)
		{
			if(i < 0)i = 0;

			CXAutoLock l(s_csFolders, TRUE);
			if(!m_mapFolders.IsEmpty())
				*pVal = m_mapFolders[i % 1000].AllocSysString();
			else return E_NOTIMPL;

			return S_OK;
		}

		STDMETHOD(put_Item)(LONG i, BSTR newVal)
		{
			if(i < 0)i = 0;

			CXStringA str(newVal, ::SysStringLen(newVal));

			str.Replace('/', '\\');
			if(!str.IsEmpty() && str[str.GetLength() - 1]!= '\\')
				str.AppendChar('\\');

			CXAutoLock l(s_csFolders, TRUE);
			m_mapFolders[i % 1000] = str;

			return S_OK;
		}

		STDMETHOD(get_BeginIndex)(LONG *pVal)
		{
			*pVal = m_nBegin;
			return S_OK;
		}

		STDMETHOD(get_EndIndex)(LONG *pVal)
		{
			*pVal = m_nEnd;
			return S_OK;
		}

		STDMETHOD(isReady)(VARIANT_BOOL* pVal)
		{
			int i;

			*pVal = VARIANT_TRUE;

			CXAutoLock l(s_csFolders, TRUE);
			for(i = 0; i < 1000; i ++)
				if(m_mapFolders[i].IsEmpty())
				{
					*pVal = VARIANT_FALSE;
					break;
				}

			return S_OK;
		}

	public:
		long m_nBegin, m_nEnd;
		CAtlArray<CXStringA> m_mapFolders;
	};

	class CXFolderMaps
	{
	public:
		void Clear()
		{
			m_mapFolders.RemoveAll();
		}

		HRESULT GetDocBand(LONG nBegin, LONG nEnd, IXPathBand** pVal)
		{
			int i;
			CXPathBand* pPath = NULL;

			s_csFolders.Enter();

			for(i = 0; i < (int)m_mapFolders.GetCount(); i ++)
			{
				pPath = m_mapFolders[i];
				if(pPath->m_nBegin == nBegin && pPath->m_nEnd == nEnd)
					break;
				pPath = NULL;
			}

			s_csFolders.Leave();

			if(!pPath)return DISP_E_BADINDEX;

			return pPath->QueryInterface(IID_IDispatch, (void**)pVal);
		}

		HRESULT newDocBand(LONG nBegin, LONG nEnd, IXPathBand** pVal)
		{
			CXComPtr<CXPathBand> pPath;

			pPath.CreateObject();

			pPath->m_nBegin = nBegin;
			pPath->m_nEnd = nEnd;

			s_csFolders.Enter();
			m_mapFolders.InsertAt(0, pPath);
			s_csFolders.Leave();

			*pVal = pPath.Detach();

			return S_OK;
		}

		CXString GetFullPath(long n)
		{
			CXString str;
			WCHAR wstr[15];
			int len, len1, slen;
			WCHAR *ptr;
			int i;
			CXPathBand* pPath = NULL;

			if(n < 0)n = 0;

			CXAutoLock l(s_csFolders, TRUE);

			for(i = 0; i < (int)m_mapFolders.GetCount(); i ++)
			{
				pPath = m_mapFolders[i];
				if(n >= pPath->m_nBegin && n <= pPath->m_nEnd)
					break;
				else pPath = NULL;
			}

			if(!pPath)
				return str;

			str = pPath->m_mapFolders[n % 1000];

			l.Leave();

			_itow_s(n, wstr, 15, 10);
			len = len1 = wcslen(wstr);

			slen = str.GetLength();

			ptr = str.GetBuffer(slen + (len / 3) * 5 + len + 1) + slen;

			while(len > 3)
			{
				*ptr++ = 'f';
				*ptr++ = wstr[len - 3];
				*ptr++ = wstr[len - 2];
				*ptr++ = wstr[len - 1];
				*ptr++ = '\\';
				len -= 3;
				slen += 5;
			}

			str.ReleaseBuffer(slen);

			return str;
		}

		HRESULT GetFullPath(long n, BSTR* pVal)
		{
			CXString str;
			str = GetFullPath(n);
			if(str.IsEmpty())
				return DISP_E_BADCALLEE;
			*pVal = str.AllocSysString();

			return S_OK;
		}

	public:
		CAtlArray< CXComPtr<CXPathBand> > m_mapFolders;
	};

public:
	CXFolderMan()
	{
		m_pUnkMarshaler = NULL;
	}

	~CXFolderMan();

DECLARE_REGISTRY_RESOURCEID(IDR_XFOLDERMAN)


BEGIN_COM_MAP(CXFolderMan)
	COM_INTERFACE_ENTRY(IXFolderMan)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
END_COM_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()
	DECLARE_GET_CONTROLLING_UNKNOWN()

	HRESULT FinalConstruct()
	{
		return CoCreateFreeThreadedMarshaler(
			GetControllingUnknown(), &m_pUnkMarshaler.p);
	}

	void FinalRelease()
	{
		m_pUnkMarshaler.Release();
	}

	CComPtr<IUnknown> m_pUnkMarshaler;

public:
	STDMETHOD(ClearFolder)(void);
	STDMETHOD(newBand)(BSTR sType, LONG nBegin, LONG nEnd, IXPathBand** pVal);
	STDMETHOD(GetBand)(BSTR sType, LONG nBegin, LONG nEnd, IXPathBand** pVal);
	STDMETHOD(GetPath)(BSTR sType, LONG i, BSTR* pVal);
	STDMETHOD(newDocBand)(LONG nBegin, LONG nEnd, IXPathBand** pVal);
	STDMETHOD(newUserBand)(LONG nBegin, LONG nEnd, IXPathBand** pVal);
	STDMETHOD(newBoardBand)(LONG nBegin, LONG nEnd, IXPathBand** pVal);
	STDMETHOD(newDataBand)(LONG nBegin, LONG nEnd, IXPathBand** pVal);
	STDMETHOD(GetDocBand)(LONG nBegin, LONG nEnd, IXPathBand** pVal);
	STDMETHOD(GetUserBand)(LONG nBegin, LONG nEnd, IXPathBand** pVal);
	STDMETHOD(GetBoardBand)(LONG nBegin, LONG nEnd, IXPathBand** pVal);
	STDMETHOD(GetDataBand)(LONG nBegin, LONG nEnd, IXPathBand** pVal);
	STDMETHOD(GetDocPath)(LONG i, BSTR* pVal);
	STDMETHOD(GetUserPath)(LONG i, BSTR* pVal);
	STDMETHOD(GetBoardPath)(LONG i, BSTR* pVal);
	STDMETHOD(GetDataPath)(LONG i, BSTR* pVal);
	STDMETHOD(ClearChannel)(void);
	STDMETHOD(AddChannel)(LONG id, BSTR path, BSTR domain, BSTR homepage, BSTR mark, BSTR folder, IXChannel** pVal);
	STDMETHOD(BatchAddChannel)(BSTR batchText);
	STDMETHOD(GetChannel)(LONG id, IXChannel** pVal);
	STDMETHOD(GetChannelFromDomain)(BSTR strDomain, IXChannel** pVal);
	STDMETHOD(GetFile)(BSTR FilePath, IXFile** ppfile);
	STDMETHOD(GetFiles)(BSTR FilePath, IXList** ppfiles);
	STDMETHOD(FileExists)(BSTR fn, VARIANT_BOOL* pVal);
	STDMETHOD(DeleteFile)(BSTR FileSpec, VARIANT bForce, LONG* pVal);
	STDMETHOD(MoveFile)(BSTR Source, BSTR Destination);
	STDMETHOD(CopyFile)(BSTR Source, BSTR Destination, VARIANT bOverwrite);
	STDMETHOD(CreateFolder)(BSTR FileSpec);
	STDMETHOD(DeleteFolder)(BSTR FileSpec);
	STDMETHOD(ParseUrl)(BSTR path, IXUrl** pVal);
	STDMETHOD(isISAPIInstalled)(VARIANT_BOOL* pVal);
	STDMETHOD(CountConnection)(BSTR Addr, LONG port, LONG* pVal);


private:
	void AddOneChannel(LONG id, LPCWSTR path, LPCWSTR domain, LPCWSTR homepage, LPCWSTR mark, LPCWSTR folder, CXChannel** pVal);

public:
	static 	void ParseUrl(LPCSTR path, CXUrl* pVal);
};

OBJECT_ENTRY_AUTO(__uuidof(XFolderMan), CXFolderMan)
