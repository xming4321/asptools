#pragma once

#include "resource.h"       // main symbols
#include "asptools.h"
#include "atlib.h"
#include "XList.h"

class CXField : public CXDispatch<IXField>
{
public:
	STDMETHOD(get_Name)(BSTR* pVal)
	{
		*pVal = m_strName.AllocSysString();
		return S_OK;
	}

	STDMETHOD(get_Type)(short* pVal)
	{
		*pVal = m_nType;
		return S_OK;
	}

public:
	CXString m_strName;
	short m_nType;
	__int64 m_nFormat;
};

class CXFields : public CXList< CXComPtr<CXField> >
{
public:
	int FindField(LPCWSTR key);
	int FindField(VARIANT key);
	HRESULT AddField(LPCWSTR key, short type, __int64 fmt = -1);
	HRESULT AddRecordset(LPCWSTR key, __int64 fmt);
	HRESULT InsertField(int pos, BSTR key, short type);
	HRESULT RemoveField(int pos);

	__int64 GetHash();
};

class CXRecord : public CXDispatch<IXRecord>
{
public:
	CXRecord()
	{
		m_nDocFormat = -1;
	}

	~CXRecord()
	{
		int i;

		for(i = 0; i < (int)m_arrayVariant.GetCount(); i ++)
			::VariantClear(&m_arrayVariant[i]);
	}

	HRESULT InitField(CXFields* pFields);

public:
	STDMETHOD(get_Item)(VARIANT key, VARIANT* pVal);
	STDMETHOD(put_Item)(VARIANT key, VARIANT newVal);
	STDMETHOD(SetFormat)(BSTR newFmt);
	STDMETHOD(get_FormatName)(BSTR* pVal);

public:
	HRESULT Save(IStream* pFile);
	HRESULT Load(IStream* pFile, __int64 nDocFormat, CXFields* pFields = NULL);
	HRESULT Convert(CXRecord* pFrom);
	HRESULT	SetValue(LPCWSTR key, VARIANT& v);

public:
	CAtlArray<VARIANT> m_arrayVariant;
	CXComPtr<CXFields> m_pFields;
	__int64 m_nDocFormat;
};

class CXRecords : public CXDispatch<IXRecords>
{
public:
	CXRecords(void)
	{
		m_nDocFormat = -1;
		m_listRecords.CreateObject();
	}

public:
	STDMETHOD(get_Item)(long i, IXRecord** pVal);
	STDMETHOD(RecordCount)(LONG* pval);
	STDMETHOD(Join)(BSTR fmtString, BSTR* pval);
	STDMETHOD(AddNew)(IXRecord **pVal);
	STDMETHOD(Remove)(int pos);
	STDMETHOD(RemoveAll)();
	STDMETHOD(SelectWhere)(BSTR key, BSTR op, VARIANT v1, VARIANT v2, IXRecords** pVal);
	STDMETHOD(CountWhere)(BSTR key, BSTR op, VARIANT v1, VARIANT v2, long* pVal);
	STDMETHOD(ExistsWhere)(BSTR key, BSTR op, VARIANT v1, VARIANT v2, VARIANT_BOOL* pVal);
	STDMETHOD(Update)(BSTR key, VARIANT newVal);
	STDMETHOD(UpdateWhere)(BSTR key, VARIANT newVal, BSTR whereKey, BSTR op, VARIANT v1, VARIANT v2, long* pVal);
	STDMETHOD(DeleteWhere)(BSTR key, BSTR op, VARIANT v1, VARIANT v2, long* pVal);
	STDMETHOD(Max)(BSTR key, VARIANT* pVal);
	STDMETHOD(Min)(BSTR key, VARIANT* pVal);
	STDMETHOD(Sort)(VARIANT key, VARIANT varAsc);
	STDMETHOD(Create)(BSTR ver);
	STDMETHOD(get_Fields)(IXList **pVal);
	STDMETHOD(AddField)(BSTR key, short type);
	STDMETHOD(InsertField)(int pos, BSTR key, short type);
	STDMETHOD(RemoveField)(int pos);
	STDMETHOD(get_FormatName)(BSTR* pVal);
	STDMETHOD(GetRows)(VARIANT key, VARIANT* pVal);

public:
	HRESULT Save(IStream* pFile);
	HRESULT Load(IStream* pFile);
	HRESULT SetFormat(__int64 nDocFormat);

	HRESULT Convert(CXRecords* pFrom);

public:
	__int64 m_nDocFormat;

private:
	void deleteItem(CXRecord* pRecord);

private:
	CXComPtr<CXFields> m_pFields;
	CXComPtr< CXList< CXComPtr<CXRecord> > > m_listRecords;
	CXComPtr<CXRecords> m_pParent;
};
