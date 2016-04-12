// XSessionMan.cpp : Implementation of CXSessionMan

#include "stdafx.h"
#include <iphlpapi.h>
#include "XSessionMan.h"
#include "XSession.h"
#include <openssl\md5.h>
#include <initguid.h>
#include "XFileStream.h"
#include "XDoc.h"
#include <stdio.h> 


static CHAR s_strVersion[] = "1.1.1049 " __DATE__;

DEFINE_GUID(CLSID_VBScript, 0xb54f3741, 0x5b07, 0x11cf, 0xa4, 0xb0, 0x0, 0xaa, 0x0, 0x4a, 0x55, 0xe8);

static BOOL s_bStart;
static CCriticalSection s_csSession;
static CCriticalSection s_csSessionClear;

static CAtlArray<CXSessionMan::CXIndex> s_arrayIndex;				// s_nOnlineTimeout 内的 Session 索引
static CXSessionMan::CXSessionList s_mapPY[29];						// 拼音映射

static CXSessionMan::CXMapSession s_mapSession;						// 全部 Session

static CXSessionMan::CXSessionList s_mapLocalOnline;				// s_nSuspendTimeout 内的本地 Session
static CXSessionMan::CXSessionList s_mapSyceOnline;					// s_nOnlineTimeout 内的全部 Session
static CXSessionMan::CXSessionList s_mapReserveSession;				// s_nOnlineTimeout 外的全部 Session

static CXComPtr<IApplicationObject> s_pApplication;
static CXComPtr<CXFields> s_pIndexFields;
static CAtlArray<CXString> s_arrayStaticObjectName;
static CAtlArray<IDispatch*> s_arrayStaticObject;
static CXComPtr<CXFields> s_pFields;
static __int64 s_nFieldHash;
static long s_nOnlineTimeout = 60, s_nSuspendTimeout = 300, s_nSessionTimeout = 12000;

static long s_nTimer = 0;
static __int64 s_dateTimer;

static CXString s_strDomain, s_strSessionKey;
static CXString s_strGlobalSource;
static DWORD s_dwSourceBaseLine;
static CXComPtr<CXSession> s_pEventSession;

static BOOL m_bExit;
static __int64 s_dateNow;
static 	MD5_CTX ctxMD5;

static SOCKET s_sMultiCast;
static SOCKADDR_IN s_dest, s_local;
static int s_nDebug;

#pragma pack(1)
struct SyncPack
{
	__int64		fmtHash;
	BYTE		op;
	BYTE		count;
	BYTE		buf[512];
};
#pragma pack()

#define OP_EVENT		128

void genExpiresString();

inline BOOL VerifyID(__int64 id)
{
	int i;
	__int64 id1 = id;

	for(i = 0; i < 8; i ++)
	{
		if((id & 0xff00000000000000) == 0)
			return FALSE;
		id <<= 8;
	}

	if(s_mapSession.Lookup(id1))return FALSE;

	return TRUE;
}

// CXSessionMan::CXSessionInst
CXSessionMan::CXSessionInst::CXSessionInst() :
	m_posLocalOnline(NULL),
	m_posSyceOnline(NULL),
	m_posReserveSession(NULL),
	m_nSessionID(0),
	m_nAccessTimes(1),
	m_posMap(NULL),
	m_posPY1(NULL),
	m_posPY2(NULL),
	m_chPY1(0),
	m_chPY2(-1),
	m_date(0),
	m_bIsLocal(FALSE)
{
	int i, count;

	if(count = s_pFields->GetCount())
	{
		m_arrayVariant.SetCount(count);
		ZeroMemory(&m_arrayVariant[0], sizeof(VARIANT) * count);

		for(i = 0; i < count; i ++)
			m_arrayVariant[i].vt = s_pFields->GetValue(i)->m_nType;
	}

	if(count = s_pIndexFields->GetCount())
	{
		m_arrayIndexVariant.SetCount(count);
		ZeroMemory(&m_arrayIndexVariant[0], sizeof(VARIANT) * count);

		for(i = 0; i < count; i ++)
			m_arrayIndexVariant[i].vt = s_pIndexFields->GetValue(i)->m_nType;
	}
}

HRESULT CXSessionMan::CXSessionInst::NewSession(IRequest* spRequest, IResponse* spResponse)
{
	while(!VerifyID(m_nSessionID))
	{
		BYTE result[MD5_DIGEST_LENGTH];
		MD5_CTX ctxMD5Temp;

		MD5_Update(&ctxMD5, &m_nSessionID, sizeof(m_nSessionID));
		MD5_Update(&ctxMD5, &s_dateNow, sizeof(s_dateNow));

		ctxMD5Temp = ctxMD5;
		MD5_Final(result, &ctxMD5Temp);
		m_nSessionID = *(__int64*)result;
	}

	m_posMap = s_mapSession.SetAt(m_nSessionID, this);

	m_dateSync = s_dateNow;
	m_date = s_dateNow;
	m_dateModify = s_dateNow;

	m_posSyceOnline = s_mapSyceOnline.AddHead(this);

	int i, count;

	count = s_pIndexFields->GetCount();
	m_arrayPos.SetCount(count);

	for(i = 0; i < count; i ++)
		m_arrayPos[i] = s_arrayIndex[i].Insert(*(CXVariant*)&m_arrayIndexVariant[i], this);

	m_posPY1 = s_mapPY[m_chPY1].AddHead(this);

	m_pContents.CoCreateObject();
	m_bIsLocal = TRUE;

	s_csSession.Leave();
	HRESULT hr = DoEvent(L"OnSessionStart", spRequest, spResponse);
	s_csSession.Enter();

	if(FAILED(hr))Offline(FALSE);

	return hr;
}

HRESULT CXSessionMan::CXSessionInst::UpdateDate(IRequest* spRequest, IResponse* spResponse)
{
	m_nAccessTimes ++;
	if(m_date == s_dateNow || (!spRequest && m_dateSync == s_dateNow))
		return S_OK;

	m_dateSync = s_dateNow;

	if(m_posReserveSession)
	{
		s_mapReserveSession.RemoveAt(m_posReserveSession);
		m_posReserveSession = NULL;
	}

	if(m_posSyceOnline)
		s_mapSyceOnline.RemoveAt(m_posSyceOnline);
	m_posSyceOnline = s_mapSyceOnline.AddHead(this);

	if(m_arrayPos.IsEmpty())
	{
		int i, count;

		count = s_pIndexFields->GetCount();
		m_arrayPos.SetCount(count);

		for(i = 0; i < count; i ++)
			m_arrayPos[i] = s_arrayIndex[i].Insert(*(CXVariant*)&m_arrayIndexVariant[i], this);

		m_posPY1 = s_mapPY[m_chPY1].AddHead(this);
		if(m_chPY2 != -1)m_posPY2 = s_mapPY[m_chPY2].AddHead(this);
	}

	HRESULT hr = S_OK;

	if(spRequest && (!m_pContents || !m_bIsLocal))
	{
		m_bIsLocal = TRUE;

		if(!m_pContents)m_pContents.CoCreateObject();

		if(m_nAccessTimes == 2)
		{
			s_csSession.Leave();
			hr = DoEvent(L"OnSessionRecover", spRequest, spResponse);
			s_csSession.Enter();

			if(FAILED(hr))Offline(FALSE);
		}else
		{
			s_csSession.Leave();
			hr = DoEvent(L"OnSessionResume", spRequest, spResponse);
			s_csSession.Enter();

			if(FAILED(hr))Suspend(FALSE);
		}
	}

	return hr;
}

void CXSessionMan::CXSessionInst::Reserve()
{
	if(m_posSyceOnline)
	{
		s_mapSyceOnline.RemoveAt(m_posSyceOnline);
		m_posSyceOnline = NULL;
		m_posReserveSession = s_mapReserveSession.AddHead(this);

		int i, count;

		count = s_pIndexFields->GetCount();

		for(i = 0; i < count; i ++)
			s_arrayIndex[i].RemoveAt(m_arrayPos[i]);

		m_arrayPos.SetCount(0);

		s_mapPY[m_chPY1].RemoveAt(m_posPY1);
		if(m_posPY2)s_mapPY[m_chPY2].RemoveAt(m_posPY2);
		m_posPY1 = m_posPY2 = NULL;
	}
}

void CXSessionMan::CXSessionInst::Suspend(BOOL bEvent)
{
	Reserve();

	if(m_pContents)
	{
		if(bEvent)
		{
			s_csSession.Leave();
			DoEvent(L"OnSessionSuspend");
			s_csSession.Enter();
		}

		m_pContents.Release();
		UpdateLocal();
	}

	if(m_nAccessTimes == 1)
	{
		m_nAccessTimes ++;
		Offline(bEvent);
	}
}

void CXSessionMan::CXSessionInst::Offline(BOOL bEvent)
{
	Suspend(bEvent);

	if(m_posReserveSession)
	{
		s_mapReserveSession.RemoveAt(m_posReserveSession);
		m_posReserveSession = NULL;

		if(m_posMap)s_mapSession.RemoveAt(m_posMap);
		m_posMap = NULL;

		if(m_date == m_dateSync && bEvent)
		{
			s_csSession.Leave();
			DoEvent(L"OnSessionEnd");
			s_csSession.Enter();
		}
	}
}

void CXSessionMan::CXSessionInst::UpdateLocal()
{
	if(m_pContents)
	{
		if(m_nAccessTimes > 1)
		{
			if(m_posLocalOnline)
			{
				if(m_date == s_dateNow)
					return;

				s_mapLocalOnline.RemoveAt(m_posLocalOnline);
			}

			m_posLocalOnline = s_mapLocalOnline.AddHead(this);
		}
		m_date = s_dateNow;
	}else if(m_posLocalOnline)
	{
		s_mapLocalOnline.RemoveAt(m_posLocalOnline);
		m_posLocalOnline = NULL;
	}
}

STDMETHODIMP CXSession::GetItemInfo(LPCOLESTR pstrName, DWORD dwReturnMask, IUnknown **ppiunkItem, ITypeInfo **ppti)
{
   	if(ppiunkItem != NULL)*ppiunkItem = NULL;
	if(ppti != NULL)*ppti = NULL;

	if(dwReturnMask & SCRIPTINFO_IUNKNOWN)
	{
		if(!_wcsicmp(pstrName, L"Application"))
			return s_pApplication.QueryInterface(ppiunkItem);

		if(!_wcsicmp(pstrName, L"Request"))
			return m_spRequest.QueryInterface(ppiunkItem);

		if(!_wcsicmp(pstrName, L"Response"))
			return m_spResponse.QueryInterface(ppiunkItem);

		if(!_wcsicmp(pstrName, L"Session"))
			return QueryInterface(IID_IUnknown, (void**)ppiunkItem);

		if(s_pApplication)
		{
			int i, count = (int)s_arrayStaticObjectName.GetCount();

			for(i = 0; i < count; i ++)
				if(!s_arrayStaticObjectName[i].CompareNoCase(pstrName))
					return s_arrayStaticObject[i]->QueryInterface(IID_IUnknown, (void**)ppiunkItem);
		}
	}else return E_FAIL;

	return S_OK;
}

HRESULT CXSession::scriptError(HRESULT hr, LPCWSTR strEvent, EXCEPINFO* pei)
{
	if(FAILED(hr))
	{
		HANDLE hEventSource = ::RegisterEventSource(NULL, _T("asptools"));

		if (hEventSource != NULL)
		{
			LPCTSTR pstrMsg;
			CString str(strEvent);

			str.AppendFormat(":Global.asa[%d] ", s_dwSourceBaseLine + m_nErrorLine);

			if(pei->bstrDescription)
			{
				str += pei->bstrDescription;
				CXObject::SetErrorInfo(str);
			}

			pstrMsg = str;

			::ReportEvent(hEventSource, EVENTLOG_ERROR_TYPE, 0, 0, NULL, 1, 0, (LPCTSTR*) &pstrMsg, NULL);
			::DeregisterEventSource(hEventSource); 
		}

		if(pei->bstrSource)::SysFreeString(pei->bstrSource);
		if(pei->bstrDescription)::SysFreeString(pei->bstrDescription);
		if(pei->bstrHelpFile)::SysFreeString(pei->bstrHelpFile);
	}

	return hr;
}

HRESULT CXSessionMan::CXSessionInst::DoEvent(LPCWSTR strEvent, IRequest* spRequest, IResponse* spResponse)
{
	CXComPtr<CXSession> pSession;
	HRESULT hr;

	hr = pSession.CoCreateObject();
	if(FAILED(hr))return hr;

	pSession->Attach(this, spRequest, spResponse);
	pSession->LoadGlobal();
	hr = pSession->DoEvent(strEvent);
	pSession->ClearGlobal();
	pSession->Detach();

	return hr;
}

void CXSessionMan::CXSessionInst::DoEvent(LPCWSTR strEvent)
{
	if(!s_pEventSession)
		s_pEventSession.CoCreateObject();

	if(!s_pEventSession->m_pDispGlobal)
		s_pEventSession->LoadGlobal();

	s_pEventSession->Attach(this, NULL, NULL);
	s_pEventSession->DoEvent(strEvent);
	s_pEventSession->Detach();
}

HRESULT CXSession::DoEvent(LPCWSTR strEvent)
{
	if(m_pDispGlobal)
	{
		DISPPARAMS dispparams = { NULL, NULL, 0, 0};
		EXCEPINFO ei = {0};
		DISPID id;
		HRESULT hr;

		hr = m_pDispGlobal.GetIDOfName(strEvent, &id);
		if(FAILED(hr) && !_wcsicmp(strEvent, L"OnSessionRecover"))
		{
			strEvent = L"OnSessionStart";
			hr = m_pDispGlobal.GetIDOfName(strEvent, &id);
		}
		if(FAILED(hr))return S_OK;

		hr = m_pDispGlobal->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, NULL, &ei, NULL);
		return scriptError(hr, strEvent, &ei);
	}

	return S_OK;
}

HRESULT CXSessionMan::CXSessionInst::get_Item(VARIANT key, VARIANT* pVal)
{
	int i = s_pIndexFields->FindField(key);

	if(i < 0 || i >= (int)s_pIndexFields->GetCount())
	{
		i = s_pFields->FindField(key);
		if(i < 0 || i >= (int)s_pFields->GetCount())
		{
			CXComPtr<CXDictionary> pContents;

			s_csSession.Enter();
			pContents = m_pContents;
			s_csSession.Leave();

			if(pContents)return pContents->get_Item(key, pVal);
			VariantClear(pVal);
			return S_OK;
		}

		s_csSession.Enter();
		VariantCopy(pVal, &m_arrayVariant[i]);
		s_csSession.Leave();

		return S_OK;
	}

	s_csSession.Enter();
	VariantCopy(pVal, &m_arrayIndexVariant[i]);
	s_csSession.Leave();

	return S_OK;
}

extern char pys[];

void CXSessionMan::CXSessionInst::put_IndexItem(int i, VARIANT& var)
{
	if(((CXVariant*)&var)->Compare(&m_arrayIndexVariant[i]))
	{
		if(m_arrayPos.GetCount())
		{
			s_arrayIndex[i].RemoveAt(m_arrayPos[i]);
			m_arrayPos[i] = s_arrayIndex[i].Insert(*(CXVariant*)&var, this);

			s_mapPY[m_chPY1].RemoveAt(m_posPY1);
			if(m_posPY2)s_mapPY[m_chPY2].RemoveAt(m_posPY2);
			m_posPY1 = m_posPY2 = NULL;

			m_posPY1 = s_mapPY[m_chPY1].AddHead(this);
			if(m_chPY2 != -1)m_posPY2 = s_mapPY[m_chPY2].AddHead(this);
		}

		VariantClear(&m_arrayIndexVariant[i]);
		CopyMemory(&m_arrayIndexVariant[i], &var, sizeof(VARIANT));

		m_dateModify = s_dateNow;
	}else VariantClear(&var);
}

inline char getCharIndex(WCHAR ch)
{
	if(ch == 0 || ch == ' ' || ch == L'　')
		return 0;
	else if(ch >= 'a' && ch <= 'z')
		return ch - 'a' + 1;
	else if(ch >= 'A' && ch <= 'Z')
		return ch - 'A' + 1;
	else if(ch >= L'ａ' && ch <= L'ｚ')
		return ch - L'ａ' + 1;
	else if(ch >= L'Ａ' && ch <= L'Ｚ')
		return ch - L'Ａ' + 1;
	else if((ch >= '0' && ch <= '9') || (ch >= L'０' && ch <= L'９'))
		return 27;

	return 28;
}

void CXSessionMan::CXSessionInst::put_Item(int i, VARIANT& var)
{
	if(((CXVariant*)&var)->Compare(&m_arrayVariant[i]))
	{
		if(m_arrayPos.GetCount())
		{
			s_mapPY[m_chPY1].RemoveAt(m_posPY1);
			if(m_posPY2)s_mapPY[m_chPY2].RemoveAt(m_posPY2);
			m_posPY1 = m_posPY2 = NULL;
		}

		if(i == 0)
		{
			WCHAR ch;

			m_chPY1 = 0;
			m_chPY2 = -1;

			if(var.bstrVal && (ch = *var.bstrVal))
			{
				if(ch >= 19968 && ch < 40869)
				{
					char* p = pys + (ch - 19968) * 2;

					if(p[0] == ' ')m_chPY1 = 0;
					else m_chPY1 = p[0] - 'a' + 1;

					if(p[1] == ' ')m_chPY2 = -1;
					else m_chPY2 = p[1] - 'a' + 1;
				}else m_chPY1 = getCharIndex(ch);
			}
		}

		if(m_arrayPos.GetCount())
		{
			m_posPY1 = s_mapPY[m_chPY1].AddHead(this);
			if(m_chPY2 != -1)m_posPY2 = s_mapPY[m_chPY2].AddHead(this);
		}

		VariantClear(&m_arrayVariant[i]);
		CopyMemory(&m_arrayVariant[i], &var, sizeof(VARIANT));

		m_dateModify = s_dateNow;
	}else VariantClear(&var);
}

HRESULT CXSessionMan::CXSessionInst::put_Item(VARIANT key, VARIANT newVal)
{
	int i = s_pIndexFields->FindField(key);
	VARIANT var = {VT_EMPTY};
	HRESULT hr;

	if(i < 0 || i >= (int)s_pIndexFields->GetCount())
	{
		i = s_pFields->FindField(key);
		if(i < 0 || i >= (int)s_pFields->GetCount())
		{
			CXComPtr<CXDictionary> pContents;

			s_csSession.Enter();
			pContents = m_pContents;
			if(pContents)m_dateModify = s_dateNow;
			s_csSession.Leave();

			if(pContents)return pContents->put_Item(key, newVal);
			return DISP_E_BADINDEX;
		}

		hr = VariantChangeType(&var, &newVal, VARIANT_ALPHABOOL, s_pFields->GetValue(i)->m_nType);
		if(SUCCEEDED(hr))
		{
			s_csSession.Enter();
			put_Item(i, var);
			s_csSession.Leave();
		}

		return hr;
	}

	hr = VariantChangeType(&var, &newVal, VARIANT_ALPHABOOL, s_pIndexFields->GetValue(i)->m_nType);
	if(SUCCEEDED(hr))
	{
		s_csSession.Enter();
		put_IndexItem(i, var);
		s_csSession.Leave();
	}

	return hr;
}

HRESULT CXSessionMan::CXSessionInst::putref_Item(VARIANT key, VARIANT newVal)
{
	CXComPtr<CXDictionary> pContents;

	s_csSession.Enter();
	pContents = m_pContents;
	s_csSession.Leave();

	if(pContents)return pContents->putref_Item(key, newVal);
	return DISP_E_BADINDEX;
}

HRESULT CXSessionMan::CXSessionInst::get_LastAccessTime(DATE* pVal)
{
	*pVal = m_date / (24 * 60 * 60 * 1000.0);
	return S_OK;
}

HRESULT CXSessionMan::CXSessionInst::get_LastModifyTime(DATE* pVal)
{
	*pVal = m_dateModify / (24 * 60 * 60 * 1000.0);
	return S_OK;
}

HRESULT CXSessionMan::CXSessionInst::get_AccessTimes(long* pVal)
{
	*pVal = m_nAccessTimes;
	return S_OK;
}

HRESULT CXSessionMan::CXSessionInst::get_Contents(IXDictionary** pVal)
{
	CXComPtr<CXDictionary> pContents;

	s_csSession.Enter();
	pContents = m_pContents;
	s_csSession.Leave();

	return pContents.QueryInterface(pVal);
}

HRESULT CXSessionMan::CXSessionInst::get_isActived(VARIANT_BOOL* pVal)
{
	*pVal = m_pContents ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}

BOOL CXSessionMan::CXSessionInst::isLocal()
{
	return m_bIsLocal;
}

HRESULT CXSessionMan::CXSessionInst::get_isLocal(VARIANT_BOOL* pVal)
{
	*pVal = m_bIsLocal ? VARIANT_TRUE : VARIANT_FALSE;

	return S_OK;
}

HRESULT CXSessionMan::CXSessionInst::get_Count(LONG* pVal)
{
	CXComPtr<CXDictionary> pContents;

	s_csSession.Enter();
	pContents = m_pContents;
	s_csSession.Leave();

	if(pContents)pContents->Count(pVal);
	else *pVal = 0;

	*pVal += s_pIndexFields->GetCount();
	return S_OK;
}

HRESULT CXSessionMan::CXSessionInst::Remove(VARIANT key)
{
	CXComPtr<CXDictionary> pContents;

	s_csSession.Enter();
	pContents = m_pContents;
	s_csSession.Leave();

	if(s_pIndexFields->FindField(key) != -1 || !pContents)return E_NOTIMPL;
	return pContents->Remove(key);
}

HRESULT CXSessionMan::CXSessionInst::RemoveAll(void)
{
	VARIANT var = {VT_EMPTY};
	int i, count;
	CXComPtr<CXDictionary> pContents;

	count = s_pIndexFields->GetCount();

	s_csSession.Enter();

	for(i = 0; i < count; i ++)
	{
		var.vt = s_pIndexFields->GetValue(i)->m_nType;
		put_IndexItem(i, var);
	}

	for(i = 0; i < count; i ++)
	{
		var.vt = s_pFields->GetValue(i)->m_nType;
		put_Item(i, var);
	}
	pContents = m_pContents;

	s_csSession.Leave();

	if(pContents)pContents->RemoveAll();

	return S_OK;
}

HRESULT CXSessionMan::CXSessionInst::get_SessionID(BSTR* pVal)
{
	WCHAR str[20];

	swprintf_s(str, 20, L"%016I64X", m_nSessionID);

	*pVal = ::SysAllocString(str);
	return S_OK;
}

// CXSession

int CXSession::LoadGlobal()
{
	if(s_strGlobalSource.IsEmpty())
		return 1;

	CXComPtr<IActiveScriptParse> pActiveScriptParse;
	CComDispatchDriver pDisp;

	if(FAILED(m_pActiveScript.CoCreateObject(CLSID_VBScript)))
		return 2;

	if(FAILED(m_pActiveScript->SetScriptSite(this)))
		return 3;

	if(FAILED(m_pActiveScript.QueryInterface(&pActiveScriptParse)))
		return 4;

	if(FAILED(pActiveScriptParse->InitNew()))
		return 5;

	if(FAILED(m_pActiveScript->AddNamedItem(L"Session", SCRIPTITEM_ISVISIBLE)))
		return 6;

	if(s_pApplication)
	{
		int i, count = (int)s_arrayStaticObjectName.GetCount();

		for(i = 0; i < count; i ++)
			if(FAILED(m_pActiveScript->AddNamedItem((BSTR)(LPCWSTR)s_arrayStaticObjectName[i], SCRIPTITEM_ISVISIBLE)))
				return 7;
	}

	if(m_spRequest)
		if(FAILED(m_pActiveScript->AddNamedItem(L"Request", SCRIPTITEM_ISVISIBLE)))
			return 8;

	if(m_spResponse)
		if(FAILED(m_pActiveScript->AddNamedItem(L"Response", SCRIPTITEM_ISVISIBLE)))
			return 9;

	if(FAILED(pActiveScriptParse->ParseScriptText(s_strGlobalSource, 0, 0, NULL, 0, 0, SCRIPTTEXT_ISPERSISTENT, 0, 0)))
		return 10;

	if(FAILED(m_pActiveScript->SetScriptState(SCRIPTSTATE_STARTED)))
		return 11;

	if(FAILED(m_pActiveScript->GetScriptDispatch(NULL, &m_pDispGlobal)))
		return 12;

	return 0;
}

void CXSession::ClearGlobal()
{
	m_pDispGlobal.Release();
	if(m_pActiveScript)
	{
		m_pActiveScript->SetScriptState(SCRIPTSTATE_CLOSED);
		m_pActiveScript.Release();
	}
}

static void loadGlobalSource(IScriptingContext* spContext)
{
	if(s_strGlobalSource.IsEmpty())
	{
		CComPtr<IServer> spServer;
		CComBSTR bstr1(L"/global.asa");
		CComBSTR bstr2;

		if(FAILED(spContext->get_Server(&spServer)))
			return;

		if(FAILED(spServer->MapPath(bstr1, &bstr2)))
			return;

		CXFileStream f;
		CXStringA strBuffer;
		const char *ptr, *ptr1, *pScript;
		int nSize;

		if(FAILED(f.Open(bstr2)))
			return;

		nSize = (int)f.GetLength();

		if(FAILED(f.Read(strBuffer.GetBuffer(nSize), nSize)))
			return;

		f.Close();

		strBuffer.ReleaseBuffer(nSize);

		ptr = strBuffer;
		while((ptr1 = strchr(ptr, '<')) && _strnicmp(ptr1, "<script", 7))
			ptr = ptr1 + 1;

		if(!ptr1)return;
		while(*ptr1 && *ptr1 != '>')ptr1++;
		if(!*ptr1++)return;
		pScript = ptr1;

		s_dwSourceBaseLine = 1;
		ptr = strBuffer;
		while(ptr < pScript)
			if(*ptr++ == '\n')s_dwSourceBaseLine ++;

		ptr = pScript;
		while((ptr1 = strchr(ptr, '<')) && _strnicmp(ptr1, "</script", 8))
			ptr = ptr1 + 1;

		if(!ptr1)return;

		CXString strScript(pScript, ptr1 - pScript);

		s_strGlobalSource = strScript;
	}
}

STDMETHODIMP CXSession::Rnd(long* pVal)
{
	BYTE result[MD5_DIGEST_LENGTH];
	MD5_CTX ctxMD5Temp;

	s_csSession.Enter();

	MD5_Update(&ctxMD5, &s_dateNow, sizeof(s_dateNow));

	ctxMD5Temp = ctxMD5;
	MD5_Final(result, &ctxMD5Temp);

	s_csSession.Leave();

	*(long*)result &= 0x7fffffff;
	if(*(long*)result == 0)*(long*)result = 1;

	*pVal = *(long*)result;

	return S_OK;
}

static HRESULT getApplication(IScriptingContext* spContext)
{
	CXComPtr<IVariantDictionary> pObjects;
	HRESULT hr;

	hr = spContext->get_Application(&s_pApplication);
	if(FAILED(hr))return hr;

	hr = s_pApplication->get_StaticObjects(&pObjects);
	if(FAILED(hr))return hr;

	int count = 0;
	VARIANT var = {VT_I4};

	pObjects->get_Count(&count);

	for(var.lVal = 1; var.lVal <= count; var.lVal ++)
	{
		CComVariant var1, var2;
		CXComPtr<IDispatch> pDispObj;

		hr = pObjects->get_Key(var, &var1);
		if(FAILED(hr))return hr;

		if(var1.vt != VT_BSTR || var1.bstrVal == NULL)
			return DISP_E_TYPEMISMATCH;

		hr = pObjects->get_Item(var1, &var2);
		if(FAILED(hr))return hr;

		if((var2.vt != VT_DISPATCH && var2.vt != VT_UNKNOWN) || (var2.pdispVal == NULL))
			return DISP_E_TYPEMISMATCH;

		hr = var2.pdispVal->QueryInterface(&pDispObj);
		if(FAILED(hr))return hr;

		s_arrayStaticObjectName.Add(var1.bstrVal);
		s_arrayStaticObject.Add(pDispObj);
	}

	return S_OK;
}

STDMETHODIMP CXSession::OnStartPage(IUnknown* pUnk)  
{
	if(!pUnk)
		return E_POINTER;

	if(!s_pIndexFields || m_pInstance)
		return E_NOTIMPL;

	if(s_strSessionKey.IsEmpty())
		s_strSessionKey = L"SessionID";

	CComPtr<IScriptingContext> spContext;
	CComPtr<IRequestDictionary> spCookies;
	CComPtr<IReadCookie> spCookie;
	CComPtr<IWriteCookie> spSendCookie;
	HRESULT hr;

	// Get the IScriptingContext Interface
	hr = pUnk->QueryInterface(__uuidof(IScriptingContext), (void **)&spContext);
	if(FAILED(hr))return hr;

	loadGlobalSource(spContext);

	if(!s_pApplication)
	{
		hr = getApplication(spContext);
		if(FAILED(hr))
		{
			s_arrayStaticObjectName.SetCount(0);
			s_arrayStaticObject.SetCount(0);
			s_pApplication.Release();
			return hr;
		}
	}

	hr = spContext->get_Request(&m_spRequest);
	if(FAILED(hr))return hr;

	hr = spContext->get_Response(&m_spResponse);
	if(FAILED(hr))return hr;

	hr = m_spRequest->get_Cookies(&spCookies);
	if(FAILED(hr))return hr;

	CComVariant varIn(s_strSessionKey);
	CComVariant varValue;

	hr = spCookies->get_Item(varIn, &varValue);
	if(FAILED(hr))return hr;

	hr = varValue.pdispVal->QueryInterface(&spCookie);
	if(FAILED(hr))return hr;

	VARIANT varOpt = {VT_ERROR};

	varOpt.scode = DISP_E_PARAMNOTFOUND;
	varValue.Clear();

	hr = spCookie->get_Item(varOpt, &varValue);
	if(FAILED(hr))return hr;

	hr = GetSession(varValue.bstrVal);
	if(FAILED(hr))return hr;

	if(m_bIsNewSession)
	{
		spCookies.Release();
		hr = m_spResponse->get_Cookies(&spCookies);
		if(FAILED(hr))return hr;

		varValue.Clear();
		hr = spCookies->get_Item(varIn, &varValue);
		if(FAILED(hr))return hr;

		hr = varValue.pdispVal->QueryInterface(&spSendCookie);
		if(FAILED(hr))return hr;

		CComBSTR bstr1;

		hr = get_SessionID(&bstr1);
		if(FAILED(hr))return hr;

		hr = spSendCookie->put_Item(varOpt, bstr1);
		if(FAILED(hr))return hr;

		if(!s_strDomain.IsEmpty())
		{
			bstr1 = s_strDomain;
			hr = spSendCookie->put_Domain(bstr1);
			if(FAILED(hr))return hr;
		}
	}

	return S_OK;
}

STDMETHODIMP CXSession::OnEndPage ()  
{
	if(m_pActiveScript)ClearGlobal();
	m_spRequest.Release();
	m_spResponse.Release();

	return S_OK;
}

STDMETHODIMP CXSession::Abandon(void)
{
	if(m_pDispGlobal || !m_pInstance)
		return E_NOTIMPL;

	if(m_spResponse)
	{
		CComPtr<IRequestDictionary> spCookies;
		CComPtr<IWriteCookie> spSendCookie;
		HRESULT hr;

		hr = m_spResponse->get_Cookies(&spCookies);
		if(FAILED(hr))return hr;

		CComVariant varIn(s_strSessionKey);
		CComVariant varValue;

		hr = spCookies->get_Item(varIn, &varValue);
		if(FAILED(hr))return hr;

		hr = varValue.pdispVal->QueryInterface(&spSendCookie);
		if(FAILED(hr))return hr;

		CComBSTR bstr1(L"");
		VARIANT varOpt = {VT_ERROR};
		varOpt.scode = DISP_E_PARAMNOTFOUND;

		hr = spSendCookie->put_Item(varOpt, bstr1);
		if(FAILED(hr))return hr;

		if(!s_strDomain.IsEmpty())
		{
			bstr1 = s_strDomain;
			hr = spSendCookie->put_Domain(bstr1);
			if(FAILED(hr))return hr;
		}
	}

	return S_OK;
}

STDMETHODIMP CXSession::Suspend(void)
{
	if(m_pDispGlobal)return E_NOTIMPL;
	s_csSession.Enter();
	m_pInstance->Suspend(TRUE);
	s_csSession.Leave();

	return S_OK;
}

STDMETHODIMP CXSession::Update(void)
{
	if(m_pDispGlobal)return E_NOTIMPL;

	s_csSession.Enter();
	m_pInstance->UpdateDate(m_spRequest, m_spResponse);
	s_csSession.Leave();

	return S_OK;
}

void CXSession::FinalRelease()
{
	if(m_pInstance)
	{
		s_csSession.Enter();
		if(s_pIndexFields)
			m_pInstance->UpdateLocal();
		s_csSession.Leave();
	}
}

STDMETHODIMP CXSession::GetSession(BSTR bstrID)
{
	if(!s_pIndexFields || m_pInstance)
		return E_NOTIMPL;

	s_bStart = TRUE;

	__int64 nID = 0;
	UINT pos = 0, len = SysStringLen(bstrID);
	WCHAR ch1, ch2;

	while(pos < len)
	{
		ch1 = bstrID[pos++];

		if((ch1 >= 'a' && ch1 <= 'f') || (ch1 >= 'A' && ch1 <= 'F'))
			ch1 = (ch1 & 0xf) + 9;
		else if(ch1 >= '0' && ch1 <= '9')
			ch1 &= 0xf;
		else
		{
			nID = 0;
			break;
		}

		if(pos < len)
		{
			ch2 = bstrID[pos++];

			if((ch2 >= 'a' && ch2 <= 'f') || (ch2 >= 'A' && ch2 <= 'F'))
				ch2 = (ch2 & 0xf) + 9;
			else if(ch2 >= '0' && ch2 <= '9')
				ch2 &= 0xf;
			else
			{
				nID = 0;
				break;
			}
		}else
		{
			nID = 0;
			break;
		}

		if(!ch1 && !ch2)
		{
			nID = 0;
			break;
		}

		nID = (nID << 8) + (ch1 << 4) + ch2;
	}

	HRESULT hr = S_OK;

	if(!nID)
	{
		m_pInstance.CreateObject();

		s_csSession.Enter();
		hr = m_pInstance->NewSession(m_spRequest, m_spResponse);
		s_csSession.Leave();

		m_bIsNewSession = TRUE;
	}else
	{
		s_csSession.Enter();
		if(!s_mapSession.Lookup(nID, m_pInstance))
		{
			m_pInstance.CreateObject();
			m_pInstance->SetSessionID(nID);
		}

		hr = m_pInstance->UpdateDate(m_spRequest, m_spResponse);
		s_csSession.Leave();
	}

	return hr;
}

STDMETHODIMP CXSession::get_Application(IDispatch** pVal)
{
	return s_pApplication.QueryInterface(pVal);
}

// CXSessionMan

STDMETHODIMP CXSessionMan::get_ComputerName(BSTR* pVal)
{
	WCHAR buf[MAX_COMPUTERNAME_LENGTH + 1];
	DWORD dwSize = MAX_COMPUTERNAME_LENGTH + 1;

	if(::GetComputerNameW(buf, &dwSize))
		*pVal = ::SysAllocStringLen(buf, dwSize);

	return S_OK;
}

STDMETHODIMP CXSessionMan::get_Version(BSTR* pVal)
{
	*pVal = CXString(s_strVersion).AllocSysString();

	return S_OK;
}

STDMETHODIMP CXSessionMan::ClearSession(void)
{
	CXMapSession::CPair * pPair;

	s_csSessionClear.Enter();

	if(s_pEventSession)
	{
		s_pEventSession->ClearGlobal();
		s_pEventSession.Release();
	}

	s_csSession.Enter();
	pPair = (CXMapSession::CPair*)s_mapSession.GetHeadPosition();
	while(pPair)
	{
		CXComPtr<CXSessionMan::CXSessionInst> pInst(pPair->m_value);

		pInst->Offline();
		pPair = (CXMapSession::CPair*)s_mapSession.GetHeadPosition();
	}
	s_csSession.Leave();

	if(s_pEventSession)
	{
		s_pEventSession->ClearGlobal();
		s_pEventSession.Release();
	}

	s_arrayStaticObjectName.SetCount(0);
	s_arrayStaticObject.SetCount(0);
	s_pApplication.Release();

	s_csSessionClear.Leave();

	return S_OK;
}

static long s_nSMCount;

HRESULT CXSessionMan::FinalConstruct()
{
	s_csSessionClear.Enter();

	s_nSMCount ++;

	if(s_nSMCount == 1)
	{
		s_dateNow = GetLocalFileTime();
		genExpiresString();

		MD5_Init(&ctxMD5);
		MD5_Update(&ctxMD5, &s_dateNow, sizeof(s_dateNow));

		s_pFields.CreateObject();
		s_pFields->AddField(L"UserName", VT_BSTR);

		s_pIndexFields.CreateObject();

		s_nFieldHash = s_pIndexFields->GetHash() | s_pFields->GetHash();

		m_hBackLoop = ::CreateThread(NULL, 0, staticBackLoopProc, 0, 0, 0);
		m_hBackClear = ::CreateThread(NULL, 0, staticBackClearProc, 0, 0, 0);
		m_hBackSync = ::CreateThread(NULL, 0, staticBackSyncProc, 0, 0, 0);
	}

	s_csSessionClear.Leave();

	return CoCreateFreeThreadedMarshaler(
		GetControllingUnknown(), &m_pUnkMarshaler.p);
}

void CXSessionMan::FinalRelease()
{
	s_csSessionClear.Enter();

	s_nSMCount --;
	if(s_nSMCount == 0)
	{
		m_bExit = TRUE;
		WaitForSingleObject(m_hBackLoop, INFINITE);
		WaitForSingleObject(m_hBackClear, INFINITE);
		WaitForSingleObject(m_hBackSync, INFINITE);

		if(s_pEventSession)
		{
			s_pEventSession->ClearGlobal();
			s_pEventSession.Release();
		}

		for(int i = 0; i < 29; i ++)
			s_mapPY[i].RemoveAll();

		s_mapSession.RemoveAll();
		s_mapLocalOnline.RemoveAll();
		s_mapSyceOnline.RemoveAll();
		s_mapReserveSession.RemoveAll();

		m_pUnkMarshaler.Release();
		s_pIndexFields.Release();
		s_pFields.Release();

		s_arrayStaticObjectName.SetCount(0);
		s_arrayStaticObject.SetCount(0);
		s_pApplication.Release();

		s_arrayIndex.SetCount(0);

		s_strGlobalSource.Empty();

		s_bStart = FALSE;
		m_bExit = FALSE;
	}

	s_csSessionClear.Leave();
}

STDMETHODIMP CXSessionMan::AddField(BSTR key, SHORT type, VARIANT varIndex)
{
	if(s_bStart)return E_NOTIMPL;

	BOOL bIndex = varGetNumbar(varIndex);

	MD5_Update(&ctxMD5, key, ::SysStringByteLen(key));
	MD5_Update(&ctxMD5, &type, 2);

	HRESULT hr;

	if(bIndex)
	{
		hr = s_pIndexFields->AddField(key, type);
		if(FAILED(hr))return hr;
		s_arrayIndex.Add();
	}else
	{
		hr = s_pFields->AddField(key, type);
		if(FAILED(hr))return hr;
	}

	s_nFieldHash = s_pIndexFields->GetHash() | s_pFields->GetHash();

	return S_OK;
}

class CXDoWhere
{
public:
	HRESULT Prepare(BSTR key, BSTR op, VARIANT v1, VARIANT v2)
	{
		HRESULT hr;

		m_index = s_pIndexFields->FindField(key);
		if(m_index < 0 || m_index >= (int)s_pIndexFields->GetCount())
			return DISP_E_BADINDEX;

		VARTYPE vt = s_pIndexFields->GetValue(m_index)->m_nType;

		if(!wcscmp(op, L"=") || !wcscmp(op, L"=="))
			m_opWhere = 1;
		else if(!wcscmp(op, L"<"))
			m_opWhere = 2;
		else if(!wcscmp(op, L"<=") || !wcscmp(op, L"=<"))
			m_opWhere = 3;
		else if(!wcscmp(op, L">"))
			m_opWhere = 4;
		else if(!wcscmp(op, L">=") || !wcscmp(op, L"=>"))
			m_opWhere = 5;
		else if(!wcscmp(op, L"!=") || !wcscmp(op, L"<>"))
			m_opWhere = 6;
		else if(!_wcsicmp(op, L"between"))
			m_opWhere = 7;
		else if(!_wcsicmp(op, L"in"))
			m_opWhere = 8;
		else return E_INVALIDARG;

		if(m_opWhere == 8)
		{
			hr = ::VariantCopyInd(&m_var1, &v1);
			if(FAILED(hr))return hr;

			if(m_var1.vt != (VT_ARRAY | VT_VARIANT))
				return E_INVALIDARG;

			ULONG i, nCount;
			LPSAFEARRAY pArray = m_var1.parray;

			nCount = 1;
			for(i = 0; i < pArray->cDims; i ++)
				nCount *= pArray->rgsabound[i].cElements;

			VARIANT* pData = (VARIANT*)pArray->pvData;
			for(i = 0; i < nCount; i ++)
			{
				hr = ::VariantChangeType(&pData[i], &pData[i], 0, vt);
				if(FAILED(hr))return hr;
			}
		}else
		{
			hr = m_var1.ChangeType(vt, &v1);
			if(FAILED(hr))return hr;
		}

		if(m_opWhere == 7)
		{
			hr = m_var2.ChangeType(vt, &v2);
			if(FAILED(hr))return hr;

			if(m_var1.Compare(&m_var2) > 0)
			{
				VARIANT varTemp;

				varTemp = *(VARIANT*)&m_var1;
				*(VARIANT*)&m_var1 = *(VARIANT*)&m_var2;
				*(VARIANT*)&m_var2 = varTemp;
			}
		}

		m_Action = 0;

		return S_OK;
	}

	HRESULT PrepareUpdate(BSTR key, VARIANT newVal, BSTR whereKey, BSTR op, VARIANT v1, VARIANT v2)
	{
		VARIANT *pvar = &newVal;
		CComVariant varTemp;

		while(pvar && pvar->vt == (VT_VARIANT | VT_BYREF))
			pvar = pvar->pvarVal;

		if(pvar && (pvar->vt == VT_UNKNOWN || pvar->vt == VT_DISPATCH))
		{
			HRESULT hr;
			CComDispatchDriver pDisp;

			if(pvar->punkVal == NULL)
				return DISP_E_UNKNOWNINTERFACE;

			hr = pvar->punkVal->QueryInterface(&pDisp);
			if(FAILED(hr))return hr;

			varTemp.Clear();
			hr = pDisp.GetProperty(0, &varTemp);
			if(FAILED(hr))return hr;

			pvar = &varTemp;
		}

		if(!pvar || !isVBType(pvar->vt))
			return E_INVALIDARG;

		HRESULT hr;

		hr = Prepare(whereKey, op, v1, v2);
		if(FAILED(hr))return hr;

		m_Action = 1;
		m_key = key;
		m_value = *pvar;
		m_bSync = FALSE;

		return S_OK;
	}

	HRESULT PreparePost(VARIANT msg, BSTR whereKey, BSTR op, VARIANT v1, VARIANT v2)
	{
		VARIANT *pvar = &msg;
		CComVariant varTemp;

		while(pvar && pvar->vt == (VT_VARIANT | VT_BYREF))
			pvar = pvar->pvarVal;

		if(pvar && (pvar->vt == VT_UNKNOWN || pvar->vt == VT_DISPATCH))
		{
			HRESULT hr;
			CComDispatchDriver pDisp;

			if(pvar->punkVal == NULL)
				return DISP_E_UNKNOWNINTERFACE;

			hr = pvar->punkVal->QueryInterface(&pDisp);
			if(FAILED(hr))return hr;

			varTemp.Clear();
			hr = pDisp.GetProperty(0, &varTemp);
			if(FAILED(hr))return hr;

			pvar = &varTemp;
		}

		if(!pvar || !isVBType(pvar->vt))
			return E_INVALIDARG;

		HRESULT hr;

		hr = Prepare(whereKey, op, v1, v2);
		if(FAILED(hr))return hr;

		m_Action = 2;
		m_key = L"";
		m_value = *pvar;
		m_bSync = FALSE;

		return S_OK;
	}

	long Invoke(CXList< CXComPtr<CXSessionMan::CXSessionInst> >* pList, BOOL bExists = FALSE)
	{
		CXComPtr< CXList< CXComPtr<CXSessionMan::CXSessionInst> > > pUpdateList;
		int count = 0;

		if(m_Action)
			if(!pList)
			{
				pUpdateList.CreateObject();
				pList = pUpdateList;
			}else pUpdateList = pList;

		s_csSession.Enter();
		if(s_arrayIndex[m_index].GetCount())
			switch(m_opWhere)
			{
			case 1: count = Invoke1(pList, bExists);break;
			case 2: count = Invoke2(pList, bExists);break;
			case 3: count = Invoke3(pList, bExists);break;
			case 4: count = Invoke4(pList, bExists);break;
			case 5: count = Invoke5(pList, bExists);break;
			case 6: count = Invoke6(pList, bExists);break;
			case 7: count = Invoke7(pList, bExists);break;
			case 8: count = Invoke8(pList, bExists);break;
			}
		s_csSession.Leave();

		if(m_Action)
		{
			BOOL bSync = FALSE;
			int i;

			for(i = 0; i < count; i ++)
			{
				CXComPtr<CXSessionMan::CXSessionInst> pInst;

				pInst = pUpdateList->GetValue(i);

				switch(m_Action)
				{
				case 1: pInst->put_Item(m_key, m_value);break;
				}

				if(!pInst->isLocal())
					bSync = TRUE;
			}

			if(bSync && s_sMultiCast && !m_bSync)SendCommand();
		}

		return count;
	}

	HRESULT SendCommand()
	{
		SyncPack bufUDP;
		int size = 0, size1;
		CXBlockStream mStream;
		CXStream StreamStub(&mStream);
		HRESULT hr;

		bufUDP.fmtHash = s_nFieldHash;
		bufUDP.op = m_opWhere + (m_Action << 4);
		bufUDP.count = 0;

		W(*(VARIANT*)&m_key);
		W(m_value.vt);
		W(*(VARIANT*)&m_value);
		W(m_index);
		if(mStream.m_dwCount > sizeof(bufUDP.buf))
			return S_FALSE;

		CopyMemory(bufUDP.buf, mStream.m_pBuffer, mStream.m_dwCount);
		size = size1 = mStream.m_dwCount;

		mStream.Clear();

		if(m_opWhere == 8)
		{
			ULONG i, nCount;
			LPSAFEARRAY pArray = m_var1.parray;

			nCount = 1;
			for(i = 0; i < pArray->cDims; i ++)
				nCount *= pArray->rgsabound[i].cElements;

			VARIANT* pData = (VARIANT*)pArray->pvData;

			for(i = 0; i < nCount; i ++)
			{
				W(pData[i]);
				if(mStream.m_dwCount <= sizeof(bufUDP.buf) - size1)
				{
					if(mStream.m_dwCount > sizeof(bufUDP.buf) - size)
					{
						sendto(s_sMultiCast, (const char *)&bufUDP, size + 10, 0, (SOCKADDR*)&s_dest, sizeof(s_dest));
						size = size1;
						bufUDP.count = 0;
					}

					CopyMemory(bufUDP.buf + size, mStream.m_pBuffer, mStream.m_dwCount);
					size += mStream.m_dwCount;
					bufUDP.count ++;

					mStream.Clear();
				}
			}
		}else
		{
			W(*(VARIANT*)&m_var1);
			if(m_opWhere == 7)W(*(VARIANT*)&m_var2);

			if(mStream.m_dwCount > sizeof(bufUDP.buf) - size)
				return S_FALSE;

			CopyMemory(bufUDP.buf + size, mStream.m_pBuffer, mStream.m_dwCount);
			size += mStream.m_dwCount;
		}

		if(size > size1)
			sendto(s_sMultiCast, (const char *)&bufUDP, size + 10, 0, (SOCKADDR*)&s_dest, sizeof(s_dest));

		return S_OK;
	}

	HRESULT ReadCommand(BYTE* ptr, int size)
	{
		m_bSync = TRUE;

		int i, count;
		CXBlockStream mStream;
		CXStream StreamStub(&mStream);
		HRESULT hr;
		VARIANT var = {VT_EMPTY};

		mStream.Attach(ptr + 2, size - 2);

		m_opWhere = ptr[0] & 0xf;
		m_Action = ptr[0] >> 4;
		count = ptr[1];

		m_key.vt = VT_BSTR;
		m_key.bstrVal = NULL;
		R(*(VARIANT*)&m_key);

		R(m_value.vt);
		m_value.bstrVal = NULL;
		R(*(VARIANT*)&m_value);
		R(m_index);

		VARTYPE vt = s_pIndexFields->GetValue(m_index)->m_nType;

		if(m_opWhere == 8)
		{
			CComSafeArray<VARIANT> Array;

			Array.Create(count);
			LPSAFEARRAY pArray = Array;
			VARIANT* pData = (VARIANT*)pArray->pvData;

			for(i = 0; i < count; i ++)
			{
				pData[i].vt = vt;
				pData[i].bstrVal = NULL;
				R(pData[i]);
			}

			m_var1.vt = VT_ARRAY | VT_VARIANT;
			m_var1.parray = Array.Detach();
		}else
		{
			m_var1.bstrVal = NULL;
			m_var1.vt = vt;
			R(*(VARIANT*)&m_var1);

			if(m_opWhere == 7)
			{
				m_var2.bstrVal = NULL;
				m_var2.vt = vt;
				R(*(VARIANT*)&m_var2);
			}
		}

		return S_OK;
	}

private:
	long Invoke1(CXList< CXComPtr<CXSessionMan::CXSessionInst> >* pList, BOOL bExists)
	{
		int count = 0;
		POSITION pos;
		CXSessionMan::CXIndex::CPair* pPair;

		pos = s_arrayIndex[m_index].FindFirstWithKey(m_var1);
		pPair = (CXSessionMan::CXIndex::CPair*)pos;
		while(pPair && !pPair->m_key.Compare(&m_var1))
		{
			if(pList)pList->AddValue(pPair->m_value);
			if(bExists)return 1;
			count ++;

			s_arrayIndex[m_index].GetNext(pos);
			pPair = (CXSessionMan::CXIndex::CPair*)pos;
		}

		return count;
	}

	long Invoke2(CXList< CXComPtr<CXSessionMan::CXSessionInst> >* pList, BOOL bExists)
	{
		int count = 0;
		POSITION pos, pos1;
		CXSessionMan::CXIndex::CPair* pPair;

		pos1 = s_arrayIndex[m_index].FindFirstKeyAfter(m_var1);
		pos = s_arrayIndex[m_index].GetHeadPosition();
		pPair = (CXSessionMan::CXIndex::CPair*)pos;
		while(pPair && pos != pos1)
		{
			if(pList)pList->AddValue(pPair->m_value);
			if(bExists)return 1;
			count ++;

			s_arrayIndex[m_index].GetNext(pos);
			pPair = (CXSessionMan::CXIndex::CPair*)pos;
		}

		return count;
	}

	long Invoke3(CXList< CXComPtr<CXSessionMan::CXSessionInst> >* pList, BOOL bExists)
	{
		int count = 0;
		POSITION pos, pos1;
		CXSessionMan::CXIndex::CPair* pPair;

		pos1 = s_arrayIndex[m_index].FindFirstKeyAfter(m_var1);
		pos = s_arrayIndex[m_index].GetHeadPosition();
		pPair = (CXSessionMan::CXIndex::CPair*)pos;
		while(pPair && pos != pos1)
		{
			if(pList)pList->AddValue(pPair->m_value);
			if(bExists)return 1;
			count ++;

			s_arrayIndex[m_index].GetNext(pos);
			pPair = (CXSessionMan::CXIndex::CPair*)pos;
		}

		pPair = (CXSessionMan::CXIndex::CPair*)pos1;
		while(pPair && !pPair->m_key.Compare(&m_var1))
		{
			if(pList)pList->AddValue(pPair->m_value);
			if(bExists)return 1;
			count ++;

			s_arrayIndex[m_index].GetNext(pos1);
			pPair = (CXSessionMan::CXIndex::CPair*)pos1;
		}

		return count;
	}

	long Invoke4(CXList< CXComPtr<CXSessionMan::CXSessionInst> >* pList, BOOL bExists)
	{
		int count = 0;
		POSITION pos;
		CXSessionMan::CXIndex::CPair* pPair;

		pos = s_arrayIndex[m_index].FindFirstKeyAfter(m_var1);
		pPair = (CXSessionMan::CXIndex::CPair*)pos;
		while(pPair && !pPair->m_key.Compare(&m_var1))
		{
			s_arrayIndex[m_index].GetNext(pos);
			pPair = (CXSessionMan::CXIndex::CPair*)pos;
		}

		while(pPair)
		{
			if(pList)pList->AddValue(pPair->m_value);
			if(bExists)return 1;
			count ++;

			s_arrayIndex[m_index].GetNext(pos);
			pPair = (CXSessionMan::CXIndex::CPair*)pos;
		}

		return count;
	}

	long Invoke5(CXList< CXComPtr<CXSessionMan::CXSessionInst> >* pList, BOOL bExists)
	{
		int count = 0;
		POSITION pos;
		CXSessionMan::CXIndex::CPair* pPair;

		pos = s_arrayIndex[m_index].FindFirstKeyAfter(m_var1);
		pPair = (CXSessionMan::CXIndex::CPair*)pos;
		while(pPair)
		{
			if(pList)pList->AddValue(pPair->m_value);
			if(bExists)return 1;
			count ++;

			s_arrayIndex[m_index].GetNext(pos);
			pPair = (CXSessionMan::CXIndex::CPair*)pos;
		}

		return count;
	}

	long Invoke6(CXList< CXComPtr<CXSessionMan::CXSessionInst> >* pList, BOOL bExists)
	{
		int count = 0;
		POSITION pos, pos1;
		CXSessionMan::CXIndex::CPair* pPair;

		pos1 = s_arrayIndex[m_index].FindFirstKeyAfter(m_var1);
		pos = s_arrayIndex[m_index].GetHeadPosition();
		pPair = (CXSessionMan::CXIndex::CPair*)pos;
		while(pPair && pos != pos1)
		{
			if(pList)pList->AddValue(pPair->m_value);
			if(bExists)return 1;
			count ++;

			s_arrayIndex[m_index].GetNext(pos);
			pPair = (CXSessionMan::CXIndex::CPair*)pos;
		}

		pos = pos1;
		pPair = (CXSessionMan::CXIndex::CPair*)pos;
		while(pPair && !pPair->m_key.Compare(&m_var1))
		{
			s_arrayIndex[m_index].GetNext(pos);
			pPair = (CXSessionMan::CXIndex::CPair*)pos;
		}

		while(pPair)
		{
			if(pList)pList->AddValue(pPair->m_value);
			if(bExists)return 1;
			count ++;

			s_arrayIndex[m_index].GetNext(pos);
			pPair = (CXSessionMan::CXIndex::CPair*)pos;
		}

		return count;
	}

	long Invoke7(CXList< CXComPtr<CXSessionMan::CXSessionInst> >* pList, BOOL bExists)
	{
		int count = 0;
		POSITION pos, pos1;
		CXSessionMan::CXIndex::CPair* pPair;

		pos = s_arrayIndex[m_index].FindFirstKeyAfter(m_var1);
		pos1 = s_arrayIndex[m_index].FindFirstKeyAfter(m_var2);

		pPair = (CXSessionMan::CXIndex::CPair*)pos;
		while(pPair && pos != pos1)
		{
			if(pList)pList->AddValue(pPair->m_value);
			if(bExists)return 1;
			count ++;

			s_arrayIndex[m_index].GetNext(pos);
			pPair = (CXSessionMan::CXIndex::CPair*)pos;
		}

		pPair = (CXSessionMan::CXIndex::CPair*)pos1;
		while(pPair && !pPair->m_key.Compare(&m_var2))
		{
			if(pList)pList->AddValue(pPair->m_value);
			if(bExists)return 1;
			count ++;

			s_arrayIndex[m_index].GetNext(pos1);
			pPair = (CXSessionMan::CXIndex::CPair*)pos1;
		}

		return count;
	}

	long Invoke8(CXList< CXComPtr<CXSessionMan::CXSessionInst> >* pList, BOOL bExists)
	{
		int count = 0;
		ULONG i, nCount;
		POSITION pos;
		CXSessionMan::CXIndex::CPair* pPair;
		LPSAFEARRAY pArray = m_var1.parray;

		nCount = 1;
		for(i = 0; i < pArray->cDims; i ++)
			nCount *= pArray->rgsabound[i].cElements;

		VARIANT* pData = (VARIANT*)pArray->pvData;
		for(i = 0; i < nCount; i ++)
		{
			pos = s_arrayIndex[m_index].FindFirstWithKey(*(CXVariant*)&pData[i]);
			pPair = (CXSessionMan::CXIndex::CPair*)pos;
			while(pPair && !pPair->m_key.Compare((CXVariant*)&pData[i]))
			{
				if(pList)pList->AddValue(pPair->m_value);
				if(bExists)return 1;
				count ++;

				s_arrayIndex[m_index].GetNext(pos);
				pPair = (CXSessionMan::CXIndex::CPair*)pos;
			}
		}

		return count;
	}

private:
	int m_index;
	BYTE m_opWhere;
	BYTE m_Action;
	BOOL m_bSync;
	CXVariant m_var1, m_var2;
	CComVariant m_key, m_value;
};

STDMETHODIMP CXSessionMan::ListUser(BSTR leadChar, IXList** pVal)
{
	int ch = getCharIndex(*leadChar);
	CXComPtr< CXList< CXComPtr<CXSessionInst> > > pList;
	pList.CreateObject();

	POSITION pos;

	s_csSession.Enter();

	pos = s_mapPY[ch].GetHeadPosition();

	while(pos)
		pList->AddValue(s_mapPY[ch].GetNext(pos));
	s_csSession.Leave();

	return pList.QueryInterface(pVal);
}

STDMETHODIMP CXSessionMan::CountUser(BSTR leadChar, long* pVal)
{
	int ch = getCharIndex(*leadChar);

	s_csSession.Enter();
	*pVal = s_mapPY[ch].GetCount();
	s_csSession.Leave();

	return S_OK;
}

STDMETHODIMP CXSessionMan::ListWhere(BSTR key, BSTR op, VARIANT v1, VARIANT v2, IXList** pVal)
{
	HRESULT hr;
	CXDoWhere dw;

	hr = dw.Prepare(key, op, v1, v2);
	if(FAILED(hr))return hr;

	CXComPtr< CXList< CXComPtr<CXSessionInst> > > pList;
	pList.CreateObject();

	dw.Invoke(pList);

	return pList.QueryInterface(pVal);
}

STDMETHODIMP CXSessionMan::CountWhere(BSTR key, BSTR op, VARIANT v1, VARIANT v2, long* pVal)
{
	HRESULT hr;
	CXDoWhere dw;

	hr = dw.Prepare(key, op, v1, v2);
	if(FAILED(hr))return hr;

	*pVal = dw.Invoke(NULL);

	return S_OK;
}

STDMETHODIMP CXSessionMan::ExistsWhere(BSTR key, BSTR op, VARIANT v1, VARIANT v2, VARIANT_BOOL* pVal)
{
	HRESULT hr;
	CXDoWhere dw;

	hr = dw.Prepare(key, op, v1, v2);
	if(FAILED(hr))return hr;

	*pVal = dw.Invoke(NULL, TRUE) ? VARIANT_TRUE : VARIANT_FALSE;

	return S_OK;
}

STDMETHODIMP CXSessionMan::UpdateWhere(BSTR key, VARIANT newVal, BSTR whereKey, BSTR op, VARIANT v1, VARIANT v2)
{
	HRESULT hr;
	CXDoWhere dw;

	hr = dw.PrepareUpdate(key, newVal, whereKey, op, v1, v2);
	if(FAILED(hr))return hr;

	dw.Invoke(NULL);

	return S_OK;
}

HRESULT CXSession::SyncEvent(BYTE * ptr, int size)
{
	CComVariant varTemp[8];
	CXBlockStream mStream;
	CXStream StreamStub(&mStream);
	CXString strEvent;
	HRESULT hr;
	int i, count;

	count = ptr[1];

	mStream.Attach(ptr + 2, size - 2);

	R(strEvent);

	for(i = 0; i < count; i ++)
	{
		R(varTemp[i].vt);
		varTemp[i].bstrVal = NULL;
		R(*(VARIANT*)&varTemp[i]);
	}
	if(m_pDispGlobal)
		hr = m_pDispGlobal.InvokeN(strEvent, (VARIANT*)&varTemp[0], i);

	return hr;
}

HRESULT CXSession::FireEvent(LPCWSTR strEvent, VARIANT *pvar, int size)
{
	if(m_pDispGlobal)
	{
		EXCEPINFO ei = {0};
		DISPPARAMS dispparams = {pvar, NULL, size, 0};
		HRESULT hr;
		DISPID id;

		hr = m_pDispGlobal.GetIDOfName(strEvent, &id);
		if(FAILED(hr))return hr;

		hr = m_pDispGlobal->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, NULL, &ei, NULL);
		if(FAILED(hr))return scriptError(hr, strEvent, &ei);

		if(s_sMultiCast)
		{
			SyncPack bufUDP;
			CXBlockStream mStream;
			CXStream StreamStub(&mStream);
			int i;

			bufUDP.fmtHash = s_nFieldHash;
			bufUDP.op = OP_EVENT;
			bufUDP.count = size;

			W(strEvent);

			for(i = 0; i < size; i ++)
			{
				W(pvar[i].vt);
				W(pvar[i]);
			}

			if(mStream.m_dwCount > sizeof(bufUDP.buf))
				return E_OUTOFMEMORY;

			CopyMemory(bufUDP.buf, mStream.m_pBuffer, mStream.m_dwCount);
			sendto(s_sMultiCast, (const char *)&bufUDP, mStream.m_dwCount + 10, 0, (SOCKADDR*)&s_dest, sizeof(s_dest));
		}
	}else return E_NOTIMPL;

	return S_OK;
}

STDMETHODIMP CXSessionMan::FireEvent(BSTR strEvent, VARIANT v1, VARIANT v2, VARIANT v3, VARIANT v4, VARIANT v5, VARIANT v6, VARIANT v7, VARIANT v8)
{
	VARIANT *pvar;
	CComVariant varTemp[8];
	int i;

	for(i = 0; i < 8 && (&v1)[i].vt != VT_ERROR; i ++)
	{
		pvar = &(&v1)[i];

		while(pvar && pvar->vt == (VT_VARIANT | VT_BYREF))
			pvar = pvar->pvarVal;

		if(pvar && (pvar->vt == VT_UNKNOWN || pvar->vt == VT_DISPATCH))
		{
			HRESULT hr;
			CComDispatchDriver pDisp;

			if(pvar->punkVal == NULL)
				return DISP_E_UNKNOWNINTERFACE;

			hr = pvar->punkVal->QueryInterface(&pDisp);
			if(FAILED(hr))return hr;

			hr = pDisp.GetProperty(0, &varTemp[7 - i]);
			if(FAILED(hr))return hr;
		}else varTemp[7 - i] = *pvar;

		if(!isVBType(varTemp[7 - i].vt))
			return E_INVALIDARG;
	}

	CXComPtr<CXSession> pSession;
	HRESULT hr;

	hr = pSession.CoCreateObject();
	if(FAILED(hr))return hr;

	pSession->LoadGlobal();
	hr = pSession->FireEvent(strEvent, (VARIANT*)&varTemp[8 - i], i);
	pSession->ClearGlobal();

	return hr;
}

STDMETHODIMP CXSessionMan::SetTimer(LONG Val)
{
	s_nTimer = Val;
	return S_OK;
}

STDMETHODIMP CXSessionMan::CountOnline(LONG* pVal)
{
	s_csSession.Enter();
	*pVal = s_mapSyceOnline.GetCount();
	s_csSession.Leave();

	return S_OK;
}

STDMETHODIMP CXSessionMan::CountLocal(LONG* pVal)
{
	s_csSession.Enter();
	*pVal = s_mapLocalOnline.GetCount();
	s_csSession.Leave();

	return S_OK;
}

STDMETHODIMP CXSessionMan::CountSession(LONG* pVal)
{
	s_csSession.Enter();
	*pVal = s_mapSyceOnline.GetCount() + s_mapReserveSession.GetCount();
	s_csSession.Leave();

	return S_OK;
}

STDMETHODIMP CXSessionMan::get_OnlineTimeout(LONG* pVal)
{
	*pVal = s_nOnlineTimeout;

	return S_OK;
}

STDMETHODIMP CXSessionMan::put_OnlineTimeout(LONG newVal)
{
	if(s_bStart)return E_NOTIMPL;
	s_nOnlineTimeout = newVal;

	return S_OK;
}

STDMETHODIMP CXSessionMan::get_SuspendTimeout(LONG* pVal)
{
	*pVal = s_nSuspendTimeout;

	return S_OK;
}

STDMETHODIMP CXSessionMan::put_SuspendTimeout(LONG newVal)
{
	if(s_bStart)return E_NOTIMPL;
	s_nSuspendTimeout = newVal;

	return S_OK;
}

STDMETHODIMP CXSessionMan::get_SessionTimeout(LONG* pVal)
{
	*pVal = s_nSessionTimeout;

	return S_OK;
}

STDMETHODIMP CXSessionMan::put_SessionTimeout(LONG newVal)
{
	if(s_bStart)return E_NOTIMPL;
	s_nSessionTimeout = newVal;

	return S_OK;
}

STDMETHODIMP CXSessionMan::get_SessionDomain(BSTR* pVal)
{
	*pVal = s_strDomain.AllocSysString();

	return S_OK;
}

STDMETHODIMP CXSessionMan::put_SessionDomain(BSTR newVal)
{
	if(s_bStart)return E_NOTIMPL;
	s_strDomain.SetString(newVal, ::SysStringLen(newVal));

	return S_OK;
}

STDMETHODIMP CXSessionMan::get_SessionKey(BSTR* pVal)
{
	*pVal = s_strSessionKey.AllocSysString();

	return S_OK;
}

STDMETHODIMP CXSessionMan::put_SessionKey(BSTR newVal)
{
	if(s_bStart)return E_NOTIMPL;
	s_strSessionKey.SetString(newVal, ::SysStringLen(newVal));

	return S_OK;
}

STDMETHODIMP CXSessionMan::get_Debug(int* pVal)
{
	*pVal = s_nDebug;
	return S_OK;
}

STDMETHODIMP CXSessionMan::Execute(BSTR strEvent, VARIANT v1, VARIANT v2, VARIANT v3, VARIANT v4, VARIANT v5, VARIANT v6, VARIANT v7, VARIANT v8, VARIANT* pVal)
{
	VARIANT *pvar;
	VARIANT varTemp[8];
	int i;
	CXComPtr<CXSession> pEventSession;

	for(i = 0; i < 8 && (&v1)[i].vt != VT_ERROR; i ++)
	{
		pvar = &(&v1)[i];

		while(pvar && pvar->vt == (VT_VARIANT | VT_BYREF))
			pvar = pvar->pvarVal;

		varTemp[7 - i] = *pvar;
	}


	if(!pEventSession)
		pEventSession.CoCreateObject();

	if(!pEventSession->m_pDispGlobal)
		pEventSession->LoadGlobal();

	if(pEventSession->m_pDispGlobal)
	{
		EXCEPINFO ei = {0};
		DISPPARAMS dispparams = { (VARIANT*)&varTemp[8 - i], NULL, i, 0};
		HRESULT hr;
		DISPID id;

		hr = pEventSession->m_pDispGlobal.GetIDOfName(strEvent, &id);
		if(FAILED(hr))return hr;

		return pEventSession->scriptError(pEventSession->m_pDispGlobal->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, pVal, &ei, NULL), strEvent, &ei);
	}
	else return E_NOTIMPL;
}

STDMETHODIMP CXSessionMan::StartBroadcast(BSTR strNetwork, BSTR strCastIP, int port)
{
	WSADATA wsd;
	SOCKET sock;
	MIB_IPADDRTABLE *pAddrs = NULL;
	ULONG uSizeAddrs = 0;

	WSAStartup(MAKEWORD(2,2),&wsd);

	s_local.sin_family = s_dest.sin_family = AF_INET;
	s_local.sin_port = s_dest.sin_port = htons(port);
	s_local.sin_addr.s_addr = 0;

	DWORD dwAddrNet = inet_addr(CXStringA(strNetwork));

	if(dwAddrNet)
	{
		GetIpAddrTable(pAddrs, &uSizeAddrs, FALSE);
		if(uSizeAddrs)
		{
			pAddrs = (MIB_IPADDRTABLE *)(new char[uSizeAddrs]);
			if(pAddrs)
			{
				if(!GetIpAddrTable(pAddrs, &uSizeAddrs, TRUE))
					for(UINT i = 0; i < pAddrs->dwNumEntries; i ++)
						if((pAddrs->table[i].dwAddr & pAddrs->table[i].dwMask) == (dwAddrNet & pAddrs->table[i].dwMask))
						{
							s_local.sin_addr.s_addr = pAddrs->table[i].dwAddr;
	//						break;
						}

				delete pAddrs;
			}
		}
	}

	s_dest.sin_addr.s_addr = inet_addr(CXStringA(strCastIP));

	sock = WSASocket(AF_INET, SOCK_DGRAM, 0, NULL, 0, WSA_FLAG_MULTIPOINT_C_LEAF | WSA_FLAG_MULTIPOINT_D_LEAF);
	if(INVALID_SOCKET == sock)
		return CXObject::WSAGetLastError();

	if(bind(sock, (struct sockaddr*)&s_local, sizeof(s_local)))
	{
		HRESULT hr = CXObject::WSAGetLastError();
		closesocket(sock);
		return hr;
	}

	BOOL bFlag = FALSE, bFlag1 = TRUE;
	char optval = 32;
	DWORD cbRet;

	if((s_dest.sin_addr.s_addr == INADDR_BROADCAST) ?
		setsockopt(sock, SOL_SOCKET, SO_BROADCAST, (char FAR *)&bFlag1, sizeof(bFlag1)) :
		(WSAIoctl(sock, SIO_MULTICAST_SCOPE, &optval, sizeof(optval), NULL, 0, &cbRet, NULL, NULL) ||
			WSAIoctl(sock, SIO_MULTIPOINT_LOOPBACK, &bFlag, sizeof(bFlag), NULL, 0, &cbRet, NULL, NULL) ||
			WSAJoinLeaf(sock, (SOCKADDR*)&s_dest, sizeof(s_dest), NULL, NULL, NULL, NULL, JL_RECEIVER_ONLY) == INVALID_SOCKET))
	{
		HRESULT hr = CXObject::WSAGetLastError();
		closesocket(sock);
		return hr;
	}

	s_sMultiCast = sock;

	return S_OK;
}

static CCriticalSection s_csSyncQueue;
static CAtlArray< CXAutoPtr<> > s_SyncQueue;
static CAtlArray< int > s_SyncQueueSize;

HRESULT CXSessionMan::CXSessionInst::readBuffer(CAtlArray<BYTE>& buf)
{
	CXBlockStream mStream;
	CXStream StreamStub(&mStream);
	HRESULT hr;
	int i, count;

	W(m_nSessionID);
	if(count = s_pFields->GetCount())
		for(i = 0; i < count; i ++)
			W(m_arrayVariant[i]);

	if(count = s_pIndexFields->GetCount())
		for(i = 0; i < count; i ++)
			W(m_arrayIndexVariant[i]);

	buf.SetCount(mStream.m_dwCount);
	CopyMemory(&buf[0], mStream.m_pBuffer, mStream.m_dwCount);

	return S_OK;
}

HRESULT CXSessionMan::CXSessionInst::syncBuffer(BYTE* buf, int size)
{
	CXBlockStream mStream;
	CXStream StreamStub(&mStream);
	HRESULT hr;
	int i, count;
	VARIANT var = {VT_EMPTY};

	mStream.Attach(buf, size);

	if(count = s_pFields->GetCount())
		for(i = 0; i < count; i ++)
		{
			var.bstrVal = NULL;
			var.vt = s_pFields->GetValue(i)->m_nType;
			R(var);
			put_Item(i, var);
		}

	if(count = s_pIndexFields->GetCount())
		for(i = 0; i < count; i ++)
		{
			var.bstrVal = NULL;
			var.vt = s_pIndexFields->GetValue(i)->m_nType;
			R(var);
			put_IndexItem(i, var);
		}

	UpdateDate(NULL, NULL);
	m_bIsLocal = FALSE;

	return S_OK;
}

void CXSessionMan::CXSessionInst::SetSessionID(__int64 id)
{
	m_nSessionID = id;

	m_posMap = s_mapSession.SetAt(m_nSessionID, this);
}

DWORD WINAPI CXSessionMan::staticBackLoopProc(LPVOID lpThreadParameter)
{
	::OleInitialize(NULL);

	SyncPack bufUDP;
	CXComPtr<CXSession> pEventSession;

	while(!m_bExit)
	{
		if(s_sMultiCast)
		{
			POSITION pos;
			CAtlArray< CAtlArray<BYTE> > SendBuffers;
			int count, i, p, size;
			CAtlArray< CXAutoPtr<> > queues;
			CAtlArray< int > queueSize;
			CXAutoPtr<> ptr;
			int bufSize;

			s_csSyncQueue.Enter();
			if(count = s_SyncQueue.GetCount())
			{
				queues.SetCount(count);
				for(i = 0; i < count; i ++)
				{
					queues[i] = s_SyncQueue[i];
					queueSize.Add(s_SyncQueueSize[i]);
				}
				s_SyncQueue.RemoveAll();
				s_SyncQueueSize.RemoveAll();
			}
			s_csSyncQueue.Leave();

			s_csSession.Enter();

			for(p = 0; p < (int)queues.GetCount(); p ++)
			{
				ptr = queues[p];
				bufSize = queueSize[p];

				count = ptr.m_p[1];

				if(ptr.m_p[0] == OP_EVENT)
				{
					s_csSession.Leave();

					if(!pEventSession)
						pEventSession.CoCreateObject();

					if(!pEventSession->m_pDispGlobal)
						pEventSession->LoadGlobal();

					pEventSession->SyncEvent(ptr, bufSize);

					s_csSession.Enter();
				}else if(ptr.m_p[0])
				{
					CXDoWhere dw;

					dw.ReadCommand(ptr, bufSize);
					dw.Invoke(NULL);
				}else
				{
					size = 2;

					for(i = 0; (i < count) && ((*(WORD*)(ptr.m_p + size) + 2) < bufSize); i ++)
					{
						CXMapSession::CPair* pPair;

						pPair = (CXMapSession::CPair*)s_mapSession.Lookup(*(__int64*)(ptr.m_p + size + 2));
						if(pPair)
							pPair->m_value->syncBuffer(ptr.m_p + size + 10, *(WORD*)(ptr.m_p + size) - 8);
						else
						{
							CXComPtr<CXSessionInst> pInstance;

							pInstance.CreateObject();
							pInstance->SetSessionID(*(__int64*)(ptr.m_p + size + 2));
							pInstance->syncBuffer(ptr.m_p + size + 10, *(WORD*)(ptr.m_p + size) - 8);
						}
						size += *(WORD*)(ptr.m_p + size) + 2;
					}
				}

				ptr.Free();
			}

			CXComPtr<CXSessionInst> pInstance;

			pos = s_mapLocalOnline.GetHeadPosition();
			while(pos)
			{
				pInstance = s_mapLocalOnline.GetNext(pos);
				if(pInstance->m_date < s_dateNow)break;

				SendBuffers.Add();
				pInstance->readBuffer(SendBuffers[SendBuffers.GetCount() - 1]);

				pInstance.Release();
			}
			
			s_dateNow = GetLocalFileTime();

			s_csSession.Leave();

			if(count = (int)SendBuffers.GetCount())
			{
				bufUDP.fmtHash = s_nFieldHash;
				bufUDP.op = 0;
				bufUDP.count = 0;
				size = 0;

				i = 0;
				while(i < count)
					if(sizeof(bufUDP.buf) - size >= SendBuffers[i].GetCount() + 2)
					{
						*(WORD*)(bufUDP.buf + size) = SendBuffers[i].GetCount();
						CopyMemory(bufUDP.buf + size + 2, &SendBuffers[i][0], SendBuffers[i].GetCount());
						bufUDP.count ++;
						size += SendBuffers[i].GetCount() + 2;
						i ++;
					}else if(size)
					{
						sendto(s_sMultiCast, (const char *)&bufUDP, size + 10, 0, (SOCKADDR*)&s_dest, sizeof(s_dest));
						bufUDP.count = 0;
						size = 0;
					}else i ++;

				if(size)
					sendto(s_sMultiCast, (const char *)&bufUDP, size + 10, 0, (SOCKADDR*)&s_dest, sizeof(s_dest));
			}

		}else s_dateNow = GetLocalFileTime();
		genExpiresString();

		long n;

		if((n = s_nTimer) > 0 && (s_dateTimer > s_dateNow || s_dateTimer + n * 1000 < s_dateNow))
		{
			s_dateTimer = s_dateNow;

			if(!pEventSession)
				pEventSession.CoCreateObject();

			if(!pEventSession->m_pDispGlobal)
				pEventSession->LoadGlobal();

			pEventSession->DoEvent(L"OnTimer");
		}

		Sleep(1);
	}

	if(pEventSession)
	{
		pEventSession->ClearGlobal();
		pEventSession.Release();
	}

	if(s_sMultiCast)
	{
		closesocket(s_sMultiCast);
		s_sMultiCast = NULL;
	}

	::OleUninitialize();

	return 0;
}

DWORD WINAPI CXSessionMan::staticBackSyncProc(LPVOID lpThreadParameter)
{
	BYTE bufUDP[sizeof(SyncPack)];
	SOCKADDR_IN from;
	int s, fl = sizeof(from);

	while(!m_bExit)
		if(s_sMultiCast)
		{
			fd_set rs = {1, s_sMultiCast};
			struct timeval tm = {0, 200};

			if(select(1, &rs, NULL, NULL, &tm) == 1)
			{
				s = recvfrom(s_sMultiCast, (char*)&bufUDP, sizeof(bufUDP), 0, (SOCKADDR*)&from, &fl);

				if(s > 12 && *(__int64*)bufUDP == s_nFieldHash && s_local.sin_addr.s_addr != from.sin_addr.s_addr)
				{
					CXAutoPtr<> ptr;

					ptr.Copy(bufUDP + 8, s - 8);
					s_csSyncQueue.Enter();
					s_SyncQueue.Add();
					s_SyncQueue[s_SyncQueue.GetCount() - 1] = ptr;
					s_SyncQueueSize.Add(s - 8);
					s_csSyncQueue.Leave();
				}
			}
		}else Sleep(1);

	if(s_sMultiCast)
	{
		closesocket(s_sMultiCast);
		s_sMultiCast = NULL;
	}

	return 0;
}

void WINAPI CXSessionMan::staticBackClearOnline()
{
	__int64 d = s_dateNow - s_nOnlineTimeout * 1000;
	CXComPtr<CXSessionInst> pInstance;

	while(s_mapSyceOnline.GetCount())
	{
		pInstance = s_mapSyceOnline.GetTail();
		if(pInstance->m_dateSync >= d)break;
		pInstance->Reserve();
	}
}

DWORD WINAPI CXSessionMan::staticBackClearProc(LPVOID lpThreadParameter)
{
	::OleInitialize(NULL);

	while(!m_bExit)
	{
		__int64 d;
		CXComPtr<CXSessionInst> pInstance;

		s_csSessionClear.Enter();

		s_csSession.Enter();

		d = s_dateNow - s_nSuspendTimeout * 1000;
		while(s_mapLocalOnline.GetCount())
		{
			staticBackClearOnline();
			pInstance = s_mapLocalOnline.GetTail();
			if(pInstance->m_date >= d)break;
			pInstance->Suspend(TRUE);
		}

		d = s_dateNow - s_nSessionTimeout * 1000;
		while(s_mapReserveSession.GetCount())
		{
			staticBackClearOnline();
			pInstance = s_mapReserveSession.GetTail();
			if(pInstance->m_dateSync >= d)break;
			pInstance->Offline();
		}

		staticBackClearOnline();

		s_csSession.Leave();

		s_csSessionClear.Leave();

		Sleep(1);
	}

	::OleUninitialize();

	return 0;
}

