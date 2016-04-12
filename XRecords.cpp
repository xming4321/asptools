#include "StdAfx.h"
#include "XRecords.h"
#include "XDoc.h"
#include <openssl\md5.h>

static char s_errInit[] = "Not Init.";

int CXFields::FindField(LPCWSTR key)
{
	int i;

	for(i = 0; i < (int)GetCount(); i ++)
		if(!GetValue(i)->m_strName.CompareNoCase(key))
			return i;

	return -1;
}

int CXFields::FindField(VARIANT key)
{
	VARIANT *pvar = &key;

	while(pvar && pvar->vt == (VT_VARIANT | VT_BYREF))
		pvar = pvar->pvarVal;

	if(!pvar)return -1;

	if(pvar->vt == VT_I2)
		return pvar->iVal;
	if(pvar->vt == VT_I4)
		return pvar->iVal;

	int i;
	CXString strKey;

	if(FAILED(varGetString(pvar, strKey)))return -1;

	for(i = 0; i < (int)GetCount(); i ++)
		if(!GetValue(i)->m_strName.CompareNoCase(strKey))
			return i;

	return -1;
}

HRESULT CXFields::AddField(LPCWSTR key, short type, __int64 fmt)
{
	int i;

	if(!isVBType(type) && type != VT_DISPATCH)
		return E_INVALIDARG;

	if(type == VT_DISPATCH)
		type = VT_USERDEFINED;

	for(i = 0; i < (int)GetCount(); i ++)
		if(!GetValue(i)->m_strName.CompareNoCase(key))
		{
			GetValue(i)->m_nType = type;
			GetValue(i)->m_nFormat = fmt;
			return S_OK;
		}

	CXComPtr<CXField> pField;

	pField.CreateObject();

	pField->m_strName = key;
	pField->m_nType = type;
	pField->m_nFormat = fmt;

	AddValue(pField);

	return S_OK;
}

HRESULT CXFields::AddRecordset(LPCWSTR key, __int64 fmt)
{
	int i;

	for(i = 0; i < (int)GetCount(); i ++)
		if(!GetValue(i)->m_strName.CompareNoCase(key))
		{
			GetValue(i)->m_nType = VT_DISPATCH;
			GetValue(i)->m_nFormat = fmt;
			return S_OK;
		}

	CXComPtr<CXField> pField;

	pField.CreateObject();

	pField->m_strName = key;
	pField->m_nType = VT_DISPATCH;
	pField->m_nFormat = fmt;

	AddValue(pField);

	return S_OK;
}

HRESULT CXFields::InsertField(int pos, BSTR key, short type)
{
	if(pos < 0 || pos >= (int)GetCount())
		return DISP_E_BADINDEX;

	if(!isVBType(type))
		return E_INVALIDARG;

	CXComPtr<CXField> pField;

	pField.CreateObject();

	pField->m_strName = key;
	pField->m_nType = type;

	InsertValue(pos, pField);

	return S_OK;
}

HRESULT CXFields::RemoveField(int pos)
{
	if(pos < 0 || pos >= (int)GetCount())
		return DISP_E_BADINDEX;

	RemoveAt(pos);
	return S_OK;
}

__int64 CXFields::GetHash()
{
	int i, count;
	CXString str;
	MD5_CTX ctxMD5;
	BYTE result[MD5_DIGEST_LENGTH];

	count = GetCount();
	MD5_Init(&ctxMD5);

	for(i =0; i < count; i ++)
	{
		str = m_arrayVariant[i]->m_strName;
		str.MakeLower();

		MD5_Update(&ctxMD5, (LPCWSTR)str, str.GetLength() * 2 + 2);
		MD5_Update(&ctxMD5, &m_arrayVariant[i]->m_nType, 2);
		if(m_arrayVariant[i]->m_nType == VT_DISPATCH || m_arrayVariant[i]->m_nType == VT_USERDEFINED)
			MD5_Update(&ctxMD5, &m_arrayVariant[i]->m_nFormat, 8);
	}

	MD5_Final(result, &ctxMD5);

	return *(__int64*)&result[0] ^ *(__int64*)&result[8];
}

HRESULT CXRecord::InitField(CXFields* pFields)
{
	m_pFields = pFields;

	HRESULT hr;
	int i, count = m_pFields->GetCount();

	if(count)
	{
		m_arrayVariant.SetCount(count);
		ZeroMemory(&m_arrayVariant[0], sizeof(VARIANT) * count);

		for(i = 0; i < count; i ++)
		{
			m_arrayVariant[i].vt = m_pFields->GetValue(i)->m_nType;
			if(m_arrayVariant[i].vt == VT_DISPATCH)
			{
				CXComPtr<CXRecords> pRecords;

				pRecords.CreateObject();

				if(m_pFields->GetValue(i)->m_nFormat != -1)
				{
					hr = pRecords->SetFormat(m_pFields->GetValue(i)->m_nFormat);
					if(FAILED(hr))return hr;
				}

				m_arrayVariant[i].pdispVal = pRecords.Detach();
			}else if(m_arrayVariant[i].vt == VT_USERDEFINED)
			{
				CXComPtr<CXRecord> pRecord;
				__int64 nDocFormat;

				pRecord.CreateObject();

				nDocFormat = m_pFields->GetValue(i)->m_nFormat;
				if(nDocFormat != -1)
				{
					CXFields* pFields;

					pFields = CXTypeManager::GetFormat(nDocFormat);
					if(!pFields)return CXObject::GetMessageFromError(ERROR_BAD_FORMAT);

					hr = pRecord->InitField(pFields);
					if(FAILED(hr))return hr;
				}
				m_arrayVariant[i].vt = VT_DISPATCH;
				m_arrayVariant[i].pdispVal = pRecord.Detach();
			}

		}
	}

	return S_OK;
}

STDMETHODIMP CXRecord::get_Item(VARIANT key, VARIANT* pVal)
{
	int i = m_pFields->FindField(key);

	if(i < 0 || i >= (int)m_pFields->GetCount())
		return DISP_E_BADINDEX;

	VariantCopy(pVal, &m_arrayVariant[i]);

	return S_OK;
}

STDMETHODIMP CXRecord::put_Item(VARIANT key, VARIANT newVal)
{
	int i = m_pFields->FindField(key);

	if(i < 0 || i >= (int)m_pFields->GetCount())
		return DISP_E_BADINDEX;

	return VariantChangeType(&m_arrayVariant[i], &newVal, VARIANT_ALPHABOOL, m_pFields->GetValue(i)->m_nType);
}

STDMETHODIMP CXRecord::SetFormat(BSTR newFmt)
{
	m_nDocFormat = CXTypeManager::GetFormatHash(newFmt);
	if(m_nDocFormat == -1)return E_NOTIMPL;

	return InitField(CXTypeManager::GetFormat(m_nDocFormat));
}

STDMETHODIMP CXRecord::get_FormatName(BSTR* pVal)
{
	*pVal = CXTypeManager::GetFormatName(m_nDocFormat);

	return S_OK;
}

HRESULT	CXRecord::SetValue(LPCWSTR key, VARIANT& v)
{
	int i = m_pFields->FindField(key);

	if(i < 0 || i >= (int)m_pFields->GetCount())
		return DISP_E_BADINDEX;

	CXField* pField = m_pFields->GetValue(i);

	if(pField->m_nType != VT_DISPATCH)
		return VariantChangeType(&m_arrayVariant[i], &v, VARIANT_ALPHABOOL, pField->m_nType);

	return DISP_E_TYPEMISMATCH;
}

HRESULT CXRecord::Convert(CXRecord* pFrom)
{
	if(pFrom)
	{
		int i;
		HRESULT hr;
		CXComPtr<CXFields> pFields = pFrom->m_pFields;

		for(i = 0; i < (int)pFields->GetCount(); i ++)
		{
			CXField* pField = pFields->GetValue(i);

			hr = SetValue(pField->m_strName, pFrom->m_arrayVariant[i]);
			if(hr == DISP_E_BADINDEX)hr = S_OK;
			if(FAILED(hr))return hr;
		}
	}

	return S_OK;
}

HRESULT CXRecord::Save(IStream* pFile)
{
	CXStream StreamStub(pFile);
	HRESULT hr;
	int i, count;
	VARTYPE vt;

	count = m_pFields->GetCount();

	for(i = 0; i < count; i ++)
	{
		vt = m_pFields->GetValue(i)->m_nType;
		if(vt == VT_DISPATCH)
		{
			if(m_arrayVariant[i].pdispVal == NULL)
				return E_NOINTERFACE;
			hr = ((CXRecords*)m_arrayVariant[i].pdispVal)->Save(pFile);
			if(FAILED(hr))return hr;
		}else if(vt == VT_USERDEFINED)
		{
			if(m_arrayVariant[i].pdispVal == NULL)
				return E_NOINTERFACE;

			__int64 nDocFormat = ((CXRecord*)m_arrayVariant[i].pdispVal)->m_nDocFormat;

			if(nDocFormat == -1)
				return CXObject::GetMessageFromError(ERROR_BAD_FORMAT);
			W(nDocFormat);

			hr = ((CXRecord*)m_arrayVariant[i].pdispVal)->Save(pFile);
			if(FAILED(hr))return hr;
		}else W(m_arrayVariant[i]);
	}

	return S_OK;
}

HRESULT CXRecord::Load(IStream* pFile, __int64 nDocFormat, CXFields* pFields)
{
	if(pFields)
		m_pFields = pFields;
	else
	{
		m_pFields = CXTypeManager::GetFormat(nDocFormat);

		if(!m_pFields)
			return CXObject::GetMessageFromError(ERROR_BAD_FORMAT);
	}

	m_nDocFormat = nDocFormat;

	CXStream StreamStub(pFile);
	HRESULT hr;
	int i, count;

	count = m_pFields->GetCount();

	if(count)
	{
		m_arrayVariant.SetCount(count);
		ZeroMemory(&m_arrayVariant[0], sizeof(VARIANT) * count);

		for(i = 0; i < count; i ++)
		{
			m_arrayVariant[i].vt = m_pFields->GetValue(i)->m_nType;
			if(m_arrayVariant[i].vt == VT_DISPATCH)
			{
				CXComPtr<CXRecords> pRecords;

				pRecords.CreateObject();

				hr = pRecords->Load(pFile);
				if(FAILED(hr))return hr;
				if(pRecords->m_nDocFormat != m_pFields->GetValue(i)->m_nFormat)
					return CXObject::GetMessageFromError(ERROR_BAD_FORMAT);

				m_arrayVariant[i].pdispVal = pRecords.Detach();
			}else if(m_arrayVariant[i].vt == VT_USERDEFINED)
			{
				CXComPtr<CXRecord> pRecord;

				R(nDocFormat);

				if(nDocFormat != m_pFields->GetValue(i)->m_nFormat && m_pFields->GetValue(i)->m_nFormat != -1)
					return CXObject::GetMessageFromError(ERROR_BAD_FORMAT);

				m_arrayVariant[i].vt = VT_DISPATCH;
				pRecord.CreateObject();

				hr = pRecord->Load(pFile, nDocFormat);
				if(FAILED(hr))return hr;

				pRecord->m_nDocFormat = nDocFormat;

				m_arrayVariant[i].pdispVal = pRecord.Detach();

			}else R(m_arrayVariant[i]);
		}
	}

	return S_OK;
}

HRESULT CXRecords::Save(IStream* pFile)
{
	if(!m_pFields)return SetErrorInfo(s_errInit);

	CXStream StreamStub(pFile);
	HRESULT hr;
	int i, count, count1;

	W(m_nDocFormat);

	if(m_nDocFormat == -1)
	{
		count = m_pFields->GetCount();
		W(count);

		for(i = 0; i < count; i ++)
		{
			W(m_pFields->GetValue(i)->m_strName);
			W(m_pFields->GetValue(i)->m_nType);
		}
	}

	count1 = m_listRecords->GetCount();
	W(count1);

	for(i = 0; i < count1; i ++)
	{
		hr = m_listRecords->GetValue(i)->Save(pFile);
		if(FAILED(hr))return hr;
	}

	return S_OK;
}

HRESULT CXRecords::Load(IStream* pFile)
{
	CXStream StreamStub(pFile);
	HRESULT hr;
	int i, count, count1;

	R(m_nDocFormat);

	if(m_nDocFormat == -1)
	{
		R(count);

		m_pFields.CreateObject();

		for(i = 0; i < count; i ++)
		{
			CXComPtr<CXField> pField;

			pField.CreateObject();

			R(pField->m_strName);
			R(pField->m_nType);

			m_pFields->AddValue(pField);
		}
	}else
	{
		m_pFields = CXTypeManager::GetFormat(m_nDocFormat);
		if(!m_pFields)return CXObject::GetMessageFromError(ERROR_BAD_FORMAT);
	}

	R(count1);

	for(i = 0; i < count1; i ++)
	{
		CXComPtr<CXRecord> pRecord;

		pRecord.CreateObject();

		hr = pRecord->Load(pFile, m_nDocFormat, m_pFields);
		if(FAILED(hr))return hr;

		m_listRecords->AddValue(pRecord);
	}

	return S_OK;
}

HRESULT CXRecords::Convert(CXRecords* pFrom)
{
	if(pFrom)
	{
		int i, count;
		HRESULT hr;

		count = pFrom->m_listRecords->GetCount();

		for(i = 0; i < count; i ++)
		{
			CXComPtr<CXRecord> pRecord;

			pRecord.CreateObject();

			hr = pRecord->InitField(m_pFields);
			if(FAILED(hr))return hr;

			hr = pRecord->Convert(pFrom->m_listRecords->GetValue(i));
			if(FAILED(hr))return hr;

			m_listRecords->AddValue(pRecord);
		}
	}

	return S_OK;
}

STDMETHODIMP CXRecords::get_Item(long i, IXRecord** pVal)
{
	if(!m_pFields)return SetErrorInfo(s_errInit);
	return m_listRecords->GetValue(i, (CXComPtr<CXRecord>*)pVal);
}

STDMETHODIMP CXRecords::RecordCount(LONG* pval)
{
	if(!m_pFields)return SetErrorInfo(s_errInit);

	return m_listRecords->Count(pval);
}

STDMETHODIMP CXRecords::Join(BSTR fmtString, BSTR* pval)
{
	if(!m_pFields)return SetErrorInfo(s_errInit);

	return m_listRecords->Join(fmtString, pval);
}

STDMETHODIMP CXRecords::AddNew(IXRecord **pVal)
{
	if(!m_pFields)return SetErrorInfo(s_errInit);

	HRESULT hr;
	CXComPtr<CXRecord> pRecord;

	if(m_pParent)
	{
		hr = m_pParent->AddNew((IXRecord**)&pRecord);
		if(FAILED(hr))return hr;
	}else
	{
		pRecord.CreateObject();
		HRESULT hr = pRecord->InitField(m_pFields);
		if(FAILED(hr))return hr;
	}

	m_listRecords->AddValue(pRecord);

	*pVal = pRecord.Detach();

	return S_OK;
}

STDMETHODIMP CXRecords::Remove(int pos)
{
	if(!m_pFields)return SetErrorInfo(s_errInit);

	m_listRecords->RemoveAt(pos);
	return S_OK;
}

STDMETHODIMP CXRecords::RemoveAll()
{
	m_listRecords->RemoveAll();
	return S_OK;
}

class CXRSWhere
{
public:
	HRESULT Prepare(CXFields* pFields, BSTR key, BSTR op, VARIANT v1, VARIANT v2)
	{
		if(!pFields)return CXObject::SetErrorInfo(s_errInit);

		HRESULT hr;

		m_index = pFields->FindField(key);
		if(m_index < 0 || m_index >= (int)pFields->GetCount())
			return DISP_E_BADINDEX;

		VARTYPE vt = pFields->GetValue(m_index)->m_nType;
		if(vt == VT_DISPATCH)return E_INVALIDARG;

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

		return S_OK;
	}

	BOOL Check(CXRecord* pRecord)
	{
		switch(m_opWhere)
		{
		case 1: return (m_var1.Compare(pRecord->m_arrayVariant[m_index]) == 0);
		case 2: return (m_var1.Compare(pRecord->m_arrayVariant[m_index]) > 0);
		case 3: return (m_var1.Compare(pRecord->m_arrayVariant[m_index]) >= 0);
		case 4: return (m_var1.Compare(pRecord->m_arrayVariant[m_index]) < 0);
		case 5: return (m_var1.Compare(pRecord->m_arrayVariant[m_index]) <= 0);
		case 6: return (m_var1.Compare(pRecord->m_arrayVariant[m_index]) != 0);
		case 7: return (m_var1.Compare(pRecord->m_arrayVariant[m_index]) <= 0 && m_var2.Compare(pRecord->m_arrayVariant[m_index]) >= 0);
		}

		ULONG i, nCount;
		LPSAFEARRAY pArray = m_var1.parray;

		nCount = 1;
		for(i = 0; i < pArray->cDims; i ++)
			nCount *= pArray->rgsabound[i].cElements;

		VARIANT* pData = (VARIANT*)pArray->pvData;
		for(i = 0; i < nCount; i ++)
			if(((CXVariant*)&pData[i])->Compare(pRecord->m_arrayVariant[m_index]) == 0)
				return TRUE;

		return FALSE;
	}

private:
	int m_index;
	BYTE m_opWhere;
	CXVariant m_var1, m_var2;
};

STDMETHODIMP CXRecords::SelectWhere(BSTR key, BSTR op, VARIANT v1, VARIANT v2, IXRecords** pVal)
{
	CXComPtr<CXRecords> pRecords;
	CXRSWhere where;
	HRESULT hr;

	hr = where.Prepare(m_pFields, key, op, v1, v2);
	if(FAILED(hr))return hr;

	int count = m_listRecords->GetCount();
	int i;

	pRecords.CreateObject();
	pRecords->m_nDocFormat = m_nDocFormat;
	pRecords->m_pFields = m_pFields;

	for(i = 0; i < count; i ++)
		if(where.Check(m_listRecords->GetValue(i)))
			pRecords->m_listRecords->AddValue(m_listRecords->GetValue(i));

	pRecords->m_pParent = this;
	*pVal = pRecords.Detach();

	return S_OK;
}

STDMETHODIMP CXRecords::CountWhere(BSTR key, BSTR op, VARIANT v1, VARIANT v2, long* pVal)
{
	CXRSWhere where;
	HRESULT hr;

	hr = where.Prepare(m_pFields, key, op, v1, v2);
	if(FAILED(hr))return hr;

	int count = m_listRecords->GetCount();
	int i, n = 0;

	for(i = 0; i < count; i ++)
		if(where.Check(m_listRecords->GetValue(i)))
			n ++;

	*pVal = n;

	return S_OK;
}

STDMETHODIMP CXRecords::ExistsWhere(BSTR key, BSTR op, VARIANT v1, VARIANT v2, VARIANT_BOOL* pVal)
{
	CXRSWhere where;
	HRESULT hr;

	hr = where.Prepare(m_pFields, key, op, v1, v2);
	if(FAILED(hr))return hr;

	int count = m_listRecords->GetCount();
	int i;

	for(i = 0; i < count; i ++)
		if(where.Check(m_listRecords->GetValue(i)))
		{
			*pVal = VARIANT_TRUE;
			return S_OK;;
		}

	*pVal = VARIANT_FALSE;

	return S_OK;
}

STDMETHODIMP CXRecords::Update(BSTR key, VARIANT newVal)
{
	HRESULT hr;

	int count = m_listRecords->GetCount();
	int i, n = 0;
	int index = m_pFields->FindField(key);

	if(index < 0 || index >= (int)m_pFields->GetCount())
		return DISP_E_BADINDEX;

	VARTYPE vt = m_pFields->GetValue(index)->m_nType;
	if(vt == VT_DISPATCH)return E_INVALIDARG;

	CComVariant var;

	hr = var.ChangeType(vt, &newVal);
	if(FAILED(hr))return hr;

	for(i = 0; i < count; i ++)
		*(CComVariant*)&(m_listRecords->GetValue(i)->m_arrayVariant[index]) = var;

	return S_OK;
}

STDMETHODIMP CXRecords::UpdateWhere(BSTR key, VARIANT newVal, BSTR whereKey, BSTR op, VARIANT v1, VARIANT v2, long* pVal)
{
	CXRSWhere where;
	HRESULT hr;

	hr = where.Prepare(m_pFields, whereKey, op, v1, v2);
	if(FAILED(hr))return hr;

	int count = m_listRecords->GetCount();
	int i, n = 0;
	int index = m_pFields->FindField(key);

	if(index < 0 || index >= (int)m_pFields->GetCount())
		return DISP_E_BADINDEX;

	VARTYPE vt = m_pFields->GetValue(index)->m_nType;
	if(vt == VT_DISPATCH)return E_INVALIDARG;

	CComVariant var;

	hr = var.ChangeType(vt, &newVal);
	if(FAILED(hr))return hr;

	for(i = 0; i < count; i ++)
		if(where.Check(m_listRecords->GetValue(i)))
		{
			*(CComVariant*)&(m_listRecords->GetValue(i)->m_arrayVariant[index]) = var;
			n ++;
		}

	*pVal = n;

	return S_OK;
}

void CXRecords::deleteItem(CXRecord* pRecord)
{
	int count = m_listRecords->GetCount();
	int i;

	for(i = 0; i < count; i ++)
	{
		if(m_listRecords->GetValue(i) == pRecord)
		{
			if(m_pParent)
				m_pParent->deleteItem(pRecord);

			m_listRecords->RemoveAt(i);
			i --;
			count --;
		}
	}
}

STDMETHODIMP CXRecords::DeleteWhere(BSTR key, BSTR op, VARIANT v1, VARIANT v2, long* pVal)
{
	CXRSWhere where;
	HRESULT hr;
	CXRecord* pRecord;

	hr = where.Prepare(m_pFields, key, op, v1, v2);
	if(FAILED(hr))return hr;

	int count = m_listRecords->GetCount();
	int i, n = 0;

	for(i = 0; i < count; i ++)
	{
		pRecord = m_listRecords->GetValue(i);
		if(where.Check(pRecord))
		{
			if(m_pParent)
				m_pParent->deleteItem(pRecord);

			m_listRecords->RemoveAt(i);
			n ++;
			i --;
			count --;
		}
	}

	*pVal = n;

	return S_OK;
}

STDMETHODIMP CXRecords::Max(BSTR key, VARIANT* pVal)
{
	int count = m_listRecords->GetCount();
	int i, n = 0;
	int index = m_pFields->FindField(key);

	if(index < 0 || index >= (int)m_pFields->GetCount())
		return DISP_E_BADINDEX;

	VARTYPE vt = m_pFields->GetValue(index)->m_nType;
	if(vt == VT_DISPATCH)return E_INVALIDARG;

	CXVariant& var = *(CXVariant*)pVal;

	for(i = 0; i < count; i ++)
	{
		VARIANT& varItem = m_listRecords->GetValue(i)->m_arrayVariant[index];
		if(var.vt == VT_EMPTY || var.Compare(varItem) < 0)
			var = varItem;
	}

	return S_OK;
}

STDMETHODIMP CXRecords::Min(BSTR key, VARIANT* pVal)
{
	int count = m_listRecords->GetCount();
	int i, n = 0;
	int index = m_pFields->FindField(key);

	if(index < 0 || index >= (int)m_pFields->GetCount())
		return DISP_E_BADINDEX;

	VARTYPE vt = m_pFields->GetValue(index)->m_nType;
	if(vt == VT_DISPATCH)return E_INVALIDARG;

	CXVariant& var = *(CXVariant*)pVal;

	for(i = 0; i < count; i ++)
	{
		VARIANT& varItem = m_listRecords->GetValue(i)->m_arrayVariant[index];
		if(var.vt == VT_EMPTY || var.Compare(varItem) > 0)
			var = varItem;
	}

	return S_OK;
}

static int s_posSort, s_bAscSort;
static CCriticalSection s_csSort;

static int __cdecl sortProc(const void *p1, const void *p2)
{
	CXRecord **pr1, **pr2;

	pr1 = (CXRecord**)p1;
	pr2 = (CXRecord**)p2;

	return s_bAscSort * ((CXVariant*)&(*pr1)->m_arrayVariant[s_posSort])->Compare(&(*pr2)->m_arrayVariant[s_posSort]);
}

STDMETHODIMP CXRecords::Sort(VARIANT key, VARIANT varAsc)
{
	if(!m_pFields)return SetErrorInfo(s_errInit);

	int pos, bAsc;

	int count = m_listRecords->GetCount();

	if(count == 0)
		return S_OK;

	pos = m_pFields->FindField(key);
	if(pos < 0 || pos >= (int)m_pFields->GetCount())
		return DISP_E_BADINDEX;

	bAsc = varGetNumbar(varAsc, 1);

	s_csSort.Enter();
	s_posSort = pos;
	s_bAscSort = bAsc ? 1 : -1;

	qsort(&m_listRecords->GetValue(0), count, sizeof(CXComPtr<CXRecord>), sortProc);

	s_csSort.Leave();

	return S_OK;
}

HRESULT CXRecords::SetFormat(__int64 nDocFormat)
{
	m_nDocFormat = nDocFormat;
	if(m_nDocFormat == -1)return E_NOTIMPL;

	m_listRecords->RemoveAll();

	m_pFields = CXTypeManager::GetFormat(m_nDocFormat);
	if(!m_pFields)return CXObject::GetMessageFromError(ERROR_BAD_FORMAT);

	return S_OK;
}

STDMETHODIMP CXRecords::Create(BSTR ver)
{
	return SetFormat(CXTypeManager::GetFormatHash(ver));
}

STDMETHODIMP CXRecords::get_Fields(IXList **pVal)
{
	if(!m_pFields)return SetErrorInfo(s_errInit);

	return m_pFields.QueryInterface(pVal);
}

STDMETHODIMP CXRecords::AddField(BSTR key, short type)
{
	if(m_nDocFormat != -1)return E_NOTIMPL;

	if(!m_pFields)
		m_pFields.CreateObject();

	HRESULT hr;

	hr = m_pFields->AddField(key, type);
	if(FAILED(hr))return hr;

	int i, count;
	VARIANT var = {type};

	count = m_listRecords->GetCount();

	for(i = 0; i < count; i ++)
		m_listRecords->GetValue(i)->m_arrayVariant.Add(var);

	return S_OK;
}

STDMETHODIMP CXRecords::InsertField(int pos, BSTR key, short type)
{
	if(m_nDocFormat != -1)return E_NOTIMPL;

	if(!m_pFields)
		m_pFields.CreateObject();

	HRESULT hr;

	hr = m_pFields->InsertField(pos, key, type);
	if(FAILED(hr))return hr;

	int i, count;
	VARIANT var = {type};

	count = m_listRecords->GetCount();

	for(i = 0; i < count; i ++)
		m_listRecords->GetValue(i)->m_arrayVariant.InsertAt(pos, var);

	return S_OK;
}

STDMETHODIMP CXRecords::RemoveField(int pos)
{
	if(m_nDocFormat != -1)return E_NOTIMPL;
	if(!m_pFields)return SetErrorInfo(s_errInit);

	HRESULT hr;

	hr = m_pFields->RemoveField(pos);
	if(FAILED(hr))return hr;

	int i, count;

	count = m_listRecords->GetCount();

	for(i = 0; i < count; i ++)
		m_listRecords->GetValue(i)->m_arrayVariant.RemoveAt(pos);

	return S_OK;
}

STDMETHODIMP CXRecords::get_FormatName(BSTR* pVal)
{
	*pVal = CXTypeManager::GetFormatName(m_nDocFormat);

	return S_OK;
}

STDMETHODIMP CXRecords::GetRows(VARIANT key, VARIANT* pVal)
{
	if(!m_pFields)return SetErrorInfo(s_errInit);

	CComSafeArray<VARIANT> bstrArray;
	VARIANT* pVar;
	HRESULT hr;
	int i = 0;
	int pos;
	int count = m_listRecords->GetCount();

	pos = m_pFields->FindField(key);
	if(pos < 0 || pos >= (int)m_pFields->GetCount())
		return DISP_E_BADINDEX;

	hr = bstrArray.Create(count);
	if(FAILED(hr))return hr;

	pVar = (VARIANT*)bstrArray.m_psa->pvData;

	for(i = 0; i < count; i ++)
	{
		hr = VariantCopy(&pVar[i], (VARIANT*)&m_listRecords->GetValue(i)->m_arrayVariant[pos]);
		if(FAILED(hr))
		{
			bstrArray.Destroy();
			return hr;
		}
	}

	pVal->vt = VT_ARRAY | VT_VARIANT;
	pVal->parray = bstrArray.Detach();

	return S_OK;
}

