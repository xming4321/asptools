// XSession.cpp : Implementation of CXSession

#include "stdafx.h"
#include "XSession.h"


STDMETHODIMP CXSession::GetLCID(LCID *plcid)
{
	return E_NOTIMPL;
}

STDMETHODIMP CXSession::GetDocVersionString(BSTR *pbstrVersion)
{
	return E_NOTIMPL;
}

STDMETHODIMP CXSession::OnScriptTerminate(const VARIANT *pvarResult, const EXCEPINFO *pexcepinfo)
{
	return E_NOTIMPL;
}

STDMETHODIMP CXSession::OnStateChange(SCRIPTSTATE ssScriptState)
{
	return E_NOTIMPL;
}

STDMETHODIMP CXSession::OnScriptError(IActiveScriptError *pscripterror)
{
	DWORD dwCookie;
	LONG nErrorChar;

	pscripterror->GetSourcePosition(&dwCookie, &m_nErrorLine, &nErrorChar);
	return E_NOTIMPL;
}

STDMETHODIMP CXSession::OnEnterScript(void)
{
	return E_NOTIMPL;
}

STDMETHODIMP CXSession::OnLeaveScript(void)
{
	return E_NOTIMPL;
}

// CXSession

STDMETHODIMP CXSession::get_Item(VARIANT key, VARIANT* pVal)
{
	if(!m_pInstance)return E_NOTIMPL;
	return m_pInstance->get_Item(key, pVal);
}

STDMETHODIMP CXSession::put_Item(VARIANT key, VARIANT newVal)
{
	if(!m_pInstance)return E_NOTIMPL;
	return m_pInstance->put_Item(key, newVal);
}

STDMETHODIMP CXSession::putref_Item(VARIANT key, VARIANT newVal)
{
	if(!m_pInstance)return E_NOTIMPL;
	return m_pInstance->putref_Item(key, newVal);
}

STDMETHODIMP CXSession::get_LastAccessTime(DATE* pVal)
{
	if(!m_pInstance)return E_NOTIMPL;
	return m_pInstance->get_LastAccessTime(pVal);
}

STDMETHODIMP CXSession::get_LastModifyTime(DATE* pVal)
{
	if(!m_pInstance)return E_NOTIMPL;
	return m_pInstance->get_LastModifyTime(pVal);
}

STDMETHODIMP CXSession::get_AccessTimes(long* pVal)
{
	if(!m_pInstance)return E_NOTIMPL;
	return m_pInstance->get_AccessTimes(pVal);
}

STDMETHODIMP CXSession::get_Contents(IXDictionary** pVal)
{
	if(!m_pInstance)return E_NOTIMPL;
	return m_pInstance->get_Contents(pVal);
}

STDMETHODIMP CXSession::get_Count(LONG* pVal)
{
	if(!m_pInstance)return E_NOTIMPL;
	return m_pInstance->get_Count(pVal);
}

STDMETHODIMP CXSession::get_SessionID(BSTR* pVal)
{
	if(!m_pInstance)return E_NOTIMPL;
	return m_pInstance->get_SessionID(pVal);
}

STDMETHODIMP CXSession::get_isActived(VARIANT_BOOL* pVal)
{
	if(!m_pInstance)return E_NOTIMPL;
	return m_pInstance->get_isActived(pVal);
}

STDMETHODIMP CXSession::get_isLocal(VARIANT_BOOL* pVal)
{
	if(!m_pInstance)return E_NOTIMPL;
	return m_pInstance->get_isLocal(pVal);
}

STDMETHODIMP CXSession::isNewSession(VARIANT_BOOL* pVal)
{
	*pVal = m_bIsNewSession ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}

STDMETHODIMP CXSession::Remove(VARIANT key)
{
	if(!m_pInstance)return E_NOTIMPL;
	return m_pInstance->Remove(key);
}

STDMETHODIMP CXSession::RemoveAll(void)
{
	if(!m_pInstance)return E_NOTIMPL;
	return m_pInstance->RemoveAll();
}

STDMETHODIMP CXSession::Execute(BSTR strEvent, VARIANT v1, VARIANT v2, VARIANT v3, VARIANT v4, VARIANT v5, VARIANT v6, VARIANT v7, VARIANT v8, VARIANT* pVal)
{
	VARIANT *pvar;
	VARIANT varTemp[8];
	int i;

	for(i = 0; i < 8 && (&v1)[i].vt != VT_ERROR; i ++)
	{
		pvar = &(&v1)[i];

		while(pvar && pvar->vt == (VT_VARIANT | VT_BYREF))
			pvar = pvar->pvarVal;

		varTemp[7 - i] = *pvar;
	}

	if(!m_pActiveScript)LoadGlobal();

	if(m_pDispGlobal)
	{
		EXCEPINFO ei = {0};
		DISPPARAMS dispparams = { (VARIANT*)&varTemp[8 - i], NULL, i, 0};
		HRESULT hr;
		DISPID id;

		hr = m_pDispGlobal.GetIDOfName(strEvent, &id);
		if(FAILED(hr))return hr;

		return scriptError(m_pDispGlobal->Invoke(id, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_METHOD, &dispparams, pVal, &ei, NULL), strEvent, &ei);
	}
	else return E_NOTIMPL;
}

