// XDoc.cpp : Implementation of CXDoc

#include "stdafx.h"
#include "XDoc.h"
#include <math.h>

#define XICIDOC		0x00434f4449434958
#define XICIZIP		0x0050495a49434958
#define XICILOG		0x00474f4c49434958

// CXDoc


HRESULT CXDoc::OpenFile(BSTR strFile, int nFlags)
{
	HRESULT hr;

	m_pBinFile.CreateObject();

	for(int i = 0; i < 64; i ++)
	{
		if(nFlags == 1)
			hr = m_pBinFile->Open(strFile, CXFileStream::modeReadWrite | CXFileStream::modeCreate, 0);
		else if(nFlags)
			hr = m_pBinFile->Create(strFile);
		else hr = m_pBinFile->Open(strFile, CXFileStream::modeRead, CXFileStream::shareRead);

		if(hr != HRESULT_FROM_WIN32(ERROR_SHARING_VIOLATION))
			return hr;

		Sleep(10);
	}

	return hr;
}

static __int64 s_mk = XICIDOC;
static __int64 s_mkCompress = XICIZIP;
static __int64 s_mkLog = XICILOG;
static __int64 s_mkPending = 0;

HRESULT CXDoc::ReadDoc(BOOL bClose)
{
	HRESULT hr;
	__int64 mk;
	int nLogPos = 0;

	hr = m_pBinFile->Read(&mk, sizeof(mk));
	if(hr == HRESULT_FROM_WIN32(ERROR_HANDLE_EOF))return S_OK;
	if(FAILED(hr))return hr;

	if(mk == XICIZIP)
	{
		hr = m_pBinFile->SetCompress(TRUE);
		if(FAILED(hr))return hr;
	}else if(mk != XICIDOC)
	{
		m_pBinFile->Seek(-(int)sizeof(int), CXFileStream::end);

		hr = m_pBinFile->Read(&nLogPos, sizeof(int));
		if(FAILED(hr))return hr;

		m_pBinFile->Seek(nLogPos);

		hr = m_pBinFile->Read(&mk, sizeof(mk));
		if(FAILED(hr))return hr;

		if(mk == XICILOG)
		{
			hr = m_pBinFile->SetCompress(TRUE);
			if(FAILED(hr))return hr;
		}else return CXObject::GetMessageFromError(ERROR_BAD_FORMAT);
	}

	CXComPtr<CXMemStream> pBufFile;
	pBufFile.CreateObject();

	hr = pBufFile->CopyFrom(m_pBinFile);
	if(bClose)m_pBinFile.Release();
	if(FAILED(hr))return hr;

	pBufFile->SeekToBegin();

	int count, i;
	CXStream StreamStub(pBufFile);

	R(m_nDocFormat);

	m_pBaseItems.Release();

	m_pBaseItems.CreateObject();

	hr = m_pBaseItems->Load(pBufFile, m_nDocFormat);
	if(FAILED(hr))return hr;

	m_mapItems.RemoveAll();
	m_bExpand = FALSE;
	R(count);

	if(count)
	{
		m_bExpand = TRUE;

		for(i = 0; i < count; i ++)
		{
			CXString key;
			CXComPtr<CXRecords> pRecords;
			VARIANT var = {VT_EMPTY};

			R(key);
			R(var.vt);

			if(var.vt == VT_DISPATCH)
			{
				pRecords.CreateObject();

				hr = pRecords->Load(pBufFile);
				if(FAILED(hr))return hr;

				var.pdispVal = pRecords.p;
			}else R(var);

			m_mapItems.SetAt(key, var);

			if(var.vt == VT_BSTR && var.bstrVal)::SysFreeString(var.bstrVal);
		}
	}

	m_listDocs->RemoveAll();
	R(count);

	for(i = 0; i < count; i ++)
	{
		CXComPtr<CXDocItem> pItem;

		pItem.CreateObject();
		hr = pItem->Load(pBufFile);
		if(FAILED(hr))
			break;

		m_listDocs->AddValue(pItem);
	}

	m_nEndOfFile = (UINT)pBufFile->GetPosition();
	pBufFile->SetSize(m_nEndOfFile);

	if(m_pBinFile)
	{
		if(mk == XICIZIP)
		{
			hr = m_pBinFile->SetCompress(FALSE);
			if(FAILED(hr))return hr;
		}else if(mk == XICILOG)
		{
			hr = m_pBinFile->SetCompress(FALSE);
			if(FAILED(hr))return hr;

			m_pBinFile->Seek(sizeof(s_mkPending));

			hr = m_pBinFile->SetCompress(TRUE);
			if(FAILED(hr))return hr;

			pBufFile->SeekToBegin();
			hr = pBufFile->CopyTo(m_pBinFile);
			if(FAILED(hr))return hr;

			hr = m_pBinFile->SetCompress(FALSE);
			if(FAILED(hr))return hr;

			hr = m_pBinFile->FlushBuffers();
			if(FAILED(hr))return hr;

			m_nEndOfFile = (UINT)m_pBinFile->GetPosition();

			m_pBinFile->SeekToBegin();
			hr = m_pBinFile->Write(&s_mkCompress, sizeof(s_mkCompress));
			if(FAILED(hr))return hr;

			hr = m_pBinFile->FlushBuffers();
			if(FAILED(hr))return hr;

			m_pBinFile->SetSize(m_nEndOfFile);
		}
	}

	if(m_pBinFile && m_bTransMode)
		m_pBufFile = pBufFile;

	return S_OK;
}

HRESULT CXDoc::WriteDoc()
{
	HRESULT hr;
	int i, count;
	CXMemStream mFile;
	CXStream StreamStub(&mFile);

	if(m_nDocFormat == -1 || !m_pBaseItems)
		return E_NOTIMPL;

	W(m_nDocFormat);

	hr = m_pBaseItems->Save(&mFile);
	if(FAILED(hr))return hr;

	count = m_mapItems.GetCount();

	W(count);

	if(count)
	{
		POSITION pos;
		CRBMap<CXString, CComVariant>::CPair* pPair;

		pos = m_mapItems.GetHeadPosition();

		while(pos)
		{
			pPair = (CRBMap<CXString, CComVariant>::CPair*)m_mapItems.GetNext(pos);
			W(pPair->m_key);
			W(pPair->m_value.vt);
			if(pPair->m_value.vt == VT_DISPATCH)
			{
				hr = ((CXRecords*)pPair->m_value.pdispVal)->Save(&mFile);
				if(FAILED(hr))return hr;
			}else W(pPair->m_value);
		}
	}

	count = m_listDocs->GetCount();

	W(count);

	for(i = 0; i < count; i ++)
	{
		hr = m_listDocs->GetValue(i)->Save(&mFile);
		if(FAILED(hr))return hr;
	}

	if(m_pBufFile)
	{
		if(mFile.GetSize() > m_nEndOfFile)
			m_nEndOfFile = mFile.GetSize();

		m_nEndOfFile = m_nEndOfFile * 2 + 1024;

		m_pBinFile->Seek(m_nEndOfFile);

		hr = m_pBinFile->Write(&s_mkLog, sizeof(s_mkLog));
		if(FAILED(hr))return hr;

		hr = m_pBinFile->SetCompress(TRUE);
		if(FAILED(hr))return hr;

		m_pBufFile->SeekToBegin();
		hr = m_pBufFile->CopyTo(m_pBinFile);
		if(FAILED(hr))return hr;

		hr = m_pBinFile->SetCompress(FALSE);
		if(FAILED(hr))return hr;

		hr = m_pBinFile->Write(&m_nEndOfFile, sizeof(m_nEndOfFile));
		if(FAILED(hr))return hr;

		hr = m_pBinFile->FlushBuffers();
		if(FAILED(hr))return hr;
	}

	m_pBinFile->SeekToBegin();

	if(m_bTransMode)
		hr = m_pBinFile->Write(&s_mkPending, sizeof(s_mkPending));
	else if(m_bCompressed)
		hr = m_pBinFile->Write(&s_mkCompress, sizeof(s_mkCompress));
	else hr = m_pBinFile->Write(&s_mk, sizeof(s_mk));
	if(FAILED(hr))return hr;

	if(m_bCompressed)
	{
		hr = m_pBinFile->SetCompress(TRUE);
		if(FAILED(hr))return hr;
	}

	mFile.SeekToBegin();
	hr = mFile.CopyTo(m_pBinFile);
	if(FAILED(hr))return hr;

	if(m_bCompressed)
	{
		hr = m_pBinFile->SetCompress(FALSE);
		if(FAILED(hr))return hr;
	}

	m_nEndOfFile = (UINT)m_pBinFile->GetPosition();

	if(m_bTransMode)
	{
		hr = m_pBinFile->FlushBuffers();
		if(FAILED(hr))return hr;

		m_pBinFile->SeekToBegin();
		if(m_bCompressed)
			hr = m_pBinFile->Write(&s_mkCompress, sizeof(s_mkCompress));
		else hr = m_pBinFile->Write(&s_mk, sizeof(s_mk));
		if(FAILED(hr))return hr;

		hr = m_pBinFile->FlushBuffers();
		if(FAILED(hr))return hr;
	}

	m_pBinFile->SetSize(m_nEndOfFile);

	return S_OK;
}

HRESULT CXDoc::SetFormat(__int64 nDocFormat)
{
	if(nDocFormat == -1)return E_NOTIMPL;

	m_bExpand = FALSE;

	m_nDocFormat = nDocFormat;
	m_pBaseItems.CreateObject();
	return m_pBaseItems->InitField(CXTypeManager::GetFormat(nDocFormat));
}

STDMETHODIMP CXDoc::Create(BSTR ver)
{
	return SetFormat(CXTypeManager::GetFormatHash(ver));
}

STDMETHODIMP CXDoc::Load(BSTR strPath)
{
	HRESULT hr;

	hr = OpenFile(strPath, 0);
	if(FAILED(hr))return hr;

	hr = ReadDoc(TRUE);
	m_pBinFile.Release();
	m_pBufFile.Release();
	if(HRESULT_FROM_WIN32(ERROR_HANDLE_EOF) == hr)
		return S_FALSE;

	return hr;
}

STDMETHODIMP CXDoc::Open(BSTR strPath, VARIANT varNew)
{
	HRESULT hr;
	int nNew = varGetNumbar(varNew);

	if(nNew)nNew = 1;

	hr = OpenFile(strPath, nNew + 1);
	if(FAILED(hr))return hr;

	hr = ReadDoc();
	if(HRESULT_FROM_WIN32(ERROR_HANDLE_EOF) == hr)
		return S_FALSE;

	if(FAILED(hr))
	{
		m_pBinFile.Release();
		m_pBufFile.Release();
	}

	return hr;
}

STDMETHODIMP CXDoc::Save(void)
{
	if(m_pBinFile == NULL)
		return CXObject::GetMessageFromError(ERROR_INVALID_HANDLE);

	HRESULT hr = WriteDoc();
	m_pBinFile.Release();

	return hr;
}

STDMETHODIMP CXDoc::Close(void)
{
	if(m_pBinFile != NULL)
		m_pBinFile.Release();

	return S_OK;
}

STDMETHODIMP CXDoc::ConvertFormat(BSTR newFmt)
{
	__int64 nNewDocFormat;
	CXComPtr<CXRecord> pNewItems;
	HRESULT hr;
	int count;

	nNewDocFormat = CXTypeManager::GetFormatHash(newFmt);
	if(nNewDocFormat == -1)return E_NOTIMPL;

	pNewItems.CreateObject();
	hr = pNewItems->InitField(CXTypeManager::GetFormat(nNewDocFormat));
	if(FAILED(hr))return hr;

	hr = pNewItems->Convert(m_pBaseItems);
	if(FAILED(hr))return hr;

	count = m_mapItems.GetCount();

	if(count)
	{
		POSITION pos;
		CRBMap<CXString, CComVariant>::CPair* pPair;

		pos = m_mapItems.GetHeadPosition();

		while(pos)
		{
			pPair = (CRBMap<CXString, CComVariant>::CPair*)m_mapItems.GetNext(pos);

			hr = pNewItems->SetValue(pPair->m_key, pPair->m_value);
			if(hr == DISP_E_BADINDEX);
			else if(FAILED(hr))return hr;
			else m_mapItems.RemoveAt(pPair);
		}
	}

	m_nDocFormat = nNewDocFormat;
	m_pBaseItems = pNewItems;

	return S_OK;
}

STDMETHODIMP CXDoc::get_FormatName(BSTR* pVal)
{
	*pVal = CXTypeManager::GetFormatName(m_nDocFormat);

	return S_OK;
}

STDMETHODIMP CXDoc::get_Compressed(VARIANT_BOOL* pVal)
{
	*pVal = m_bCompressed ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}

STDMETHODIMP CXDoc::put_Compressed(VARIANT_BOOL newVal)
{
	m_bCompressed = (newVal != VARIANT_FALSE);
	return S_OK;
}

STDMETHODIMP CXDoc::get_TransMode(VARIANT_BOOL* pVal)
{
	*pVal = m_bTransMode ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}

STDMETHODIMP CXDoc::put_TransMode(VARIANT_BOOL newVal)
{
	if(newVal == VARIANT_FALSE)
	{
		m_pBufFile.Release();
	}else if(m_pBinFile)
		return E_NOTIMPL;

	m_bTransMode = (newVal != VARIANT_FALSE);
	return S_OK;
}

STDMETHODIMP CXDoc::get_Item(BSTR key, VARIANT* pVal)
{
	HRESULT hr = DISP_E_BADINDEX;
	VARIANT varKey = {VT_BSTR};
	varKey.bstrVal = key;

	if(m_pBaseItems)
	{
		hr = m_pBaseItems->get_Item(varKey, pVal);
		if(hr != DISP_E_BADINDEX)return hr;
	}

	if(m_bExpand)
	{
		CXString strKey;

		strKey = key;
		strKey.MakeLower();

		m_mapItems.Lookup(strKey, *(CComVariant*)pVal);

		hr = S_OK;
	}

	return hr;
}

STDMETHODIMP CXDoc::put_Item(BSTR key, VARIANT newVal)
{
	HRESULT hr = DISP_E_BADINDEX;
	VARIANT varKey = {VT_BSTR};
	varKey.bstrVal = key;

	if(m_pBaseItems)
	{
		hr = m_pBaseItems->put_Item(varKey, newVal);
		if(hr != DISP_E_BADINDEX)return hr;
	}

	if(m_bExpand)
	{
		VARIANT *pvar = &newVal;
		VARIANT varTemp = {VT_EMPTY};

		while(pvar && pvar->vt == (VT_VARIANT | VT_BYREF))
			pvar = pvar->pvarVal;

		if(!pvar)return DISP_E_TYPEMISMATCH;

		if(pvar->vt == VT_UNKNOWN || pvar->vt == VT_DISPATCH)
		{
			HRESULT hr;
			CComDispatchDriver pDisp;

			if(pvar->punkVal == NULL)
				return DISP_E_UNKNOWNINTERFACE;

			hr = pvar->punkVal->QueryInterface(&pDisp);
			if(FAILED(hr))return hr;

			hr = pDisp.GetProperty(0, &varTemp);
			if(FAILED(hr))return hr;

			pvar = &varTemp;
		}

		if(!isVBType(pvar->vt))
		{
			if(varTemp.vt)VariantClear(&varTemp);
			return DISP_E_TYPEMISMATCH;
		}

		CXString strKey;

		strKey = key;
		strKey.MakeLower();

		m_mapItems.SetAt(strKey, *pvar);

		if(varTemp.vt)VariantClear(&varTemp);

		hr = S_OK;
	}

	return hr;
}

STDMETHODIMP CXDoc::Count(LONG* pVal)
{
	*pVal = (long)m_mapItems.GetCount();
	if(m_pBaseItems && m_pBaseItems->m_pFields)
		*pVal += m_pBaseItems->m_pFields->GetCount();

	return S_OK;
}

STDMETHODIMP CXDoc::get__NewEnum(IUnknown** pVal)
{
	return getNewEnum(this, pVal);
}

void CXDoc::FillEnum(CAtlArray<VARIANT>& arrayVariant)
{
	if(m_pBaseItems)
		m_pBaseItems->m_pFields->FillEnum(arrayVariant);

	POSITION pos;
	CRBMap<CXString, CComVariant>::CPair* pPair;

	pos = m_mapItems.GetHeadPosition();

	while(pos)
	{
		VARIANT var = {VT_EMPTY};

		pPair = (CRBMap<CXString, CComVariant>::CPair*)m_mapItems.GetNext(pos);
		*(CComVariant*)&var = pPair->m_key;
		arrayVariant.Add(var);
	}
}

STDMETHODIMP CXDoc::Remove(BSTR key)
{
	CXString strKey;

	strKey = key;
	strKey.MakeLower();

	m_mapItems.RemoveKey(strKey);
	return S_OK;
}

STDMETHODIMP CXDoc::Exists(BSTR key, VARIANT_BOOL* pVal)
{
	if(m_pBaseItems && m_pBaseItems->m_pFields->FindField(key) != -1)
	{
		*pVal = VARIANT_TRUE;
		return S_OK;
	}

	CXString strKey;

	strKey = key;
	strKey.MakeLower();

	*pVal = m_mapItems.Lookup(strKey) ? VARIANT_TRUE : VARIANT_FALSE;

	return S_OK;
}

STDMETHODIMP CXDoc::get_docs(IXList** pVal)
{
	return m_listDocs.QueryInterface(pVal);
}

STDMETHODIMP CXDoc::Append(BSTR user, LONG id, SHORT emote, SHORT rate, SHORT hot, BSTR host, BSTR ip, IXDocItem** pVal)
{
	CXComPtr<CXDocItem> pItem;

	pItem.CreateObject();
	pItem->SetValue(user, id, emote, rate, hot, host, ip);

	m_listDocs->AddValue(pItem);

	return pItem.QueryInterface(pVal);
}

STDMETHODIMP CXDoc::RemoveDoc(long i)
{
	if(i >= 0 && i < (int)m_listDocs->GetCount())
		m_listDocs->RemoveAt(i);
	return S_OK;
}

STDMETHODIMP CXDoc::AddRecordset(BSTR key, VARIANT varFmt, IXRecords** pVal)
{
	if(m_pBaseItems && m_pBaseItems->m_pFields->FindField(key) != -1)
		return DISP_E_TYPEMISMATCH;

	if(!m_bExpand)return DISP_E_BADINDEX;

	CXComPtr<CXRecords> pRecords;

	pRecords.CreateObject();

	CXString strFmt;

	varGetString(varFmt, strFmt);

	if(!strFmt.IsEmpty())
	{
		HRESULT hr = pRecords->Create((BSTR)(LPCWSTR)strFmt);
		if(FAILED(hr))return hr;
	}

	VARIANT var = {VT_DISPATCH};
	var.pdispVal = pRecords.p;

	CXString strKey;

	strKey = key;
	strKey.MakeLower();

	m_mapItems.SetAt(strKey, var);

	return pRecords.QueryInterface(pVal);
}

STDMETHODIMP CXDoc::CountHost(long* pval)
{
	int count = m_listDocs->GetCount();
	int i;
	CRBMap<CXString, int> mapCount;

	for(i = 0; i < count; i ++)
		mapCount.SetAt(m_listDocs->GetValue(i)->m_strHost, 1);

	*pval = (long)mapCount.GetCount();

	return S_OK;
}

STDMETHODIMP CXDoc::CountIP(long* pval)
{
	int count = m_listDocs->GetCount();
	int i;
	CRBMap<CXString, int> mapCount;

	for(i = 0; i < count; i ++)
		mapCount.SetAt(m_listDocs->GetValue(i)->m_strIP, 1);

	*pval = (long)mapCount.GetCount();

	return S_OK;
}

STDMETHODIMP CXDoc::CountUser(long* pval)
{
	int count = m_listDocs->GetCount();
	int i;
	CRBMap<CXString, int> mapCount;

	for(i = 0; i < count; i ++)
		mapCount.SetAt(m_listDocs->GetValue(i)->m_strUserName, 1);

	*pval = (long)mapCount.GetCount();

	return S_OK;
}

STDMETHODIMP CXDoc::CountRate(long* pval)
{
	int count = m_listDocs->GetCount();
	int i, nRateCount;
	CRBMap<CXString, int> mapCount;
	CXComPtr<CXDocItem> pItem;

	nRateCount = 0;
	for(i = 0; i < count; i ++)
	{
		pItem = m_listDocs->GetValue(i);

		if(pItem->m_nRate)
			if(!mapCount.Lookup(pItem->m_strHost))
			{
				mapCount.SetAt(pItem->m_strHost, 1);
				nRateCount += pItem->m_nRate;
			}
	}

	if(mapCount.GetCount() > 0)
		*pval = nRateCount * 10 / mapCount.GetCount();
	else *pval = 0;

	return S_OK;
}

STDMETHODIMP CXDoc::updateTime(long nStart, long nEnd, DATE* pval)
{
	int count = m_listDocs->GetCount();
	int i;
	DATE d, d1;
	CXComPtr<CXDocItem> pItem;

	if(nStart >= count)nStart = count;
	if(nEnd >= count)nEnd = count;
	if(nStart > nEnd)
		nStart = nEnd;

	d = d1 = 0;
	for(i = nStart; i < nEnd; i ++)
	{
		pItem = m_listDocs->GetValue(i);

		pItem->get_updateTime(&d1);
		if(d1 > d)
			d = d1;
	}

	*pval = d;

	return S_OK;
}

STDMETHODIMP CXDoc::HotRank(long timeline, long timeline1, double* pval)
{
	int count = m_listDocs->GetCount();
	int i;
	CXComPtr<CXDocItem> pItem;
	CRBMap<CXString, int> mapCountIP;
	CRBMap<CXString, int> mapCountHost;
	CXString strKey;
	double n, tl = 24.0 / timeline;
	DATE baseTime, time;

	n = 0;
	baseTime = GetVariantTime();

	for(i = count - 1; i >= 0; i --)
	{
		pItem = m_listDocs->GetValue(i);
		if(!pItem->isDel())
		{
			time = pItem->updateTime();

			strKey = pItem->m_strIP;
			if(mapCountIP.Lookup(strKey) == NULL)
			{
				mapCountIP.SetAt(strKey, 1);
				n += (pItem->m_nHot - 1) * 10 * pow(2, (time - baseTime) * tl) * pow(2, - (double)i / timeline1);
			}

			strKey = pItem->m_strHost;
			if(mapCountHost.Lookup(strKey) == NULL)
			{
				mapCountHost.SetAt(strKey, 1);
				n += (pItem->m_nHot - 1) * 10 * pow(2, (time - baseTime) * tl) * pow(2, - (double)i / timeline1);
			}
		}
	}

	*pval = n;

	return S_OK;
}

STDMETHODIMP CXDoc::Expand(void)
{
	m_bExpand = TRUE;
	return S_OK;
}

