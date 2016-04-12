// XList.h : Declaration of the CXList

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"
#include "atlib.h"
#include "XMemStream.h"


// CXList

template<class TYPE>
class CXList : public CXDispatch<IXList>
{
public:
	CXList()
	{
	}

	~CXList() 
	{
		RemoveAll();
	}

public:
	STDMETHOD(get_Item)(long i, VARIANT* pVal)
	{
		if(i >=0 && i < (int)m_arrayVariant.GetCount())
			*(CComVariant*)pVal = m_arrayVariant[i];
		else return DISP_E_BADINDEX;

		return S_OK;
	}

	STDMETHOD(Count)(LONG* pval)
	{
		*pval = (long)m_arrayVariant.GetCount();

		return S_OK;
	}

	STDMETHOD(Join)(BSTR fmtString, BSTR* pval)
	{
		int i, count;
		LPCWSTR ptr, ptr1, ptrEnd;
		CXMemStream ms;
		HRESULT hr;

		count = m_arrayVariant.GetCount();
		ptrEnd = fmtString + ::SysStringLen(fmtString);

		for(i = 0; i < count; i ++)
		{
			CXString strKey;
			CComVariant var, val;
			CComDispatchDriver pDisp;

			var = m_arrayVariant[i];

			if(var.vt != VT_DISPATCH || var.pdispVal == NULL)
				return E_NOTIMPL;

			pDisp = var.pdispVal;
			ptr = fmtString;

			while(ptr < ptrEnd)
			{
				DISPPARAMS dispparams = { &var, NULL, 1, 0};

				ptr1 = ptr;
				while(ptr < ptrEnd && *ptr != '$')ptr ++;

				hr = ms.Write(ptr1, (ptr - ptr1) * 2);
				if(FAILED(hr))return hr;

				if(ptr == ptrEnd)break;
				ptr ++;
				if(ptr == ptrEnd || *ptr == '$')
				{
					hr = ms.Write(L"$", 2);
					if(FAILED(hr))return hr;
					ptr ++;
					continue;
				}

				ptr1 = ptr;
				while(ptr < ptrEnd && *ptr != '$')ptr ++;

				strKey.SetString(ptr1, ptr - ptr1);
				var = strKey;

				pDisp->Invoke(0, IID_NULL, LOCALE_USER_DEFAULT, DISPATCH_PROPERTYGET, &dispparams, &val, NULL, NULL);

				if(val.vt != VT_BSTR)val.ChangeType(VT_BSTR);

				hr = ms.Write(val.bstrVal, ::SysStringByteLen(val.bstrVal));
				val.Clear();
				if(FAILED(hr))return hr;

				if(ptr == ptrEnd)break;
				ptr ++;
			}
		}

		ms.SeekToBegin();
		count = (int)ms.GetLength();

		*pval = ::SysAllocStringByteLen(NULL, count);
		ms.Read(*pval, count);

		return S_OK;
	}

	STDMETHOD(get__NewEnum)(IUnknown** pVal)
	{
		return getNewEnum(this, pVal);
	}

public:
	void AddValue(const TYPE& value)
	{
		m_arrayVariant.Add(value);
	}

	void InsertValue(int pos, const TYPE& value)
	{
		if(pos >= 0 && pos < (int)m_arrayVariant.GetCount())
			m_arrayVariant.InsertAt(pos, value);
	}

	size_t GetCount(void)
	{
		return m_arrayVariant.GetCount();
	}

	void SetAt(size_t iElement, const TYPE& value)
	{
		m_arrayVariant[iElement] = value;
	}

	void RemoveAll(void)
	{
		m_arrayVariant.RemoveAll();
	}

	void RemoveAt(size_t iElement)
	{
		m_arrayVariant.RemoveAt(iElement);
	}

	void FillEnum(CAtlArray<VARIANT>& arrayVariant)
	{
		int i;

		for (i = 0; i < (int)m_arrayVariant.GetCount(); i ++)
		{
			VARIANT var = {VT_EMPTY};

			*(CComVariant*)&var = m_arrayVariant[i];
			arrayVariant.Add(var);
		}
	}

	HRESULT GetValue(long i, TYPE* pVal)
	{
		if(i >=0 && i < (int)m_arrayVariant.GetCount())
			*pVal = m_arrayVariant[i];
		else return DISP_E_BADINDEX;

		return S_OK;
	}

	const TYPE& GetValue( size_t iElement ) const
	{
		return( m_arrayVariant[iElement] );
	}

	TYPE& GetValue( size_t iElement )
	{
		return( m_arrayVariant[iElement] );
	}

protected:
	CAtlArray<TYPE> m_arrayVariant;
};

