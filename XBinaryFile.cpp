// XBinaryFile.cpp : Implementation of CXBinaryFile

#include "stdafx.h"
#include "XBinaryFile.h"
#include "XMemstream.h"


// CXBinaryFile


STDMETHODIMP CXBinaryFile::Create(BSTR bstrName, VARIANT varOverwrite)
{
	return m_File.Create(bstrName, (VARIANT_BOOL)varGetNumbar(varOverwrite));
}

STDMETHODIMP CXBinaryFile::Open(BSTR bstrName, VARIANT varMode, VARIANT varShareMode)
{
	return m_File.Open(bstrName, (short)varGetNumbar(varMode, 1), (short)varGetNumbar(varShareMode, 3));
}

STDMETHODIMP CXBinaryFile::get_EOF(VARIANT_BOOL* pVal)
{
	*pVal = m_File.isEOF() ? VARIANT_TRUE : VARIANT_FALSE;

	return S_OK;
}

STDMETHODIMP CXBinaryFile::get_lastModify(DATE* pVal)
{
	*pVal = m_File.GetlastModify();

	return S_OK;
}

STDMETHODIMP CXBinaryFile::get_Position(DOUBLE* pVal)
{
	LARGE_INTEGER lVal = {0};
	HRESULT hr;

	hr = m_File.Seek(lVal, SEEK_CUR, (ULARGE_INTEGER*)&lVal);
	if(FAILED(hr))return hr;

	*pVal = (DOUBLE)lVal.QuadPart;

	return S_OK;
}

STDMETHODIMP CXBinaryFile::put_Position(DOUBLE newVal)
{
	LARGE_INTEGER lVal;
	lVal.QuadPart = (LONGLONG)newVal;

	return m_File.Seek(lVal, newVal >= 0 ? SEEK_SET : SEEK_END, NULL);
}

STDMETHODIMP CXBinaryFile::get_Size(DOUBLE* pVal)
{
	*pVal = (double)m_File.GetLength();

	return S_OK;
}

STDMETHODIMP CXBinaryFile::put_Size(DOUBLE newVal)
{
	m_File.SetSize((ULONGLONG)newVal);

	return S_OK;
}

STDMETHODIMP CXBinaryFile::get_Compressed(VARIANT_BOOL* pVal)
{
	*pVal = m_File.GetCompress() ? VARIANT_TRUE :VARIANT_FALSE;
	return S_OK;
}

STDMETHODIMP CXBinaryFile::put_Compressed(VARIANT_BOOL newVal)
{
	return m_File.SetCompress(newVal != VARIANT_FALSE);
}

STDMETHODIMP CXBinaryFile::Read(VARIANT varSize, VARIANT* pVal)
{
	long nSize = varGetNumbar(varSize, -1);

	if(nSize == 0)
		return S_OK;

	CXVarPtr varPtr;
	ULONG nRead;
	HRESULT hr;

	if(nSize < 0)
	{
		DOUBLE nTotalSize;
		DOUBLE nPosition;

		hr = get_Size(&nTotalSize);
		if(SUCCEEDED(hr))
		{
			hr = get_Position(&nPosition);
			if(SUCCEEDED(hr))
			{
				if(nTotalSize - nPosition > 0x7fffffff)
					return DISP_E_OVERFLOW;
				nSize = (long)(nTotalSize - nPosition);
			}
		}
	}

	if(nSize < 0)
	{
		ULARGE_INTEGER cb = {0xffffffff, 0xffffffff};
		CXMemStream mStream;

		hr = mStream.CopyFrom(&m_File, cb.QuadPart, NULL, &cb.QuadPart);
		if(FAILED(hr))return hr;

		if(cb.QuadPart == 0)
			return CXObject::GetMessageFromError(ERROR_HANDLE_EOF);

		if(cb.HighPart != 0)
			return DISP_E_OVERFLOW;

		return mStream.GetVariant(pVal);
	}

	hr = varPtr.Create(nSize);
	if(FAILED(hr))return hr;

	hr = m_File.Read(varPtr.m_pData, nSize, &nRead);
	if(FAILED(hr))return hr;

	return varPtr.GetVariant(pVal, nRead);
}

STDMETHODIMP CXBinaryFile::Write(VARIANT varData)
{
	CXVarPtr varPtr;
	HRESULT hr;

	hr = varPtr.Attach(varData);
	if(FAILED(hr))return hr;

	return m_File.Write(varPtr.m_pData, varPtr.m_nSize);
}

STDMETHODIMP CXBinaryFile::WriteText(BSTR strData, VARIANT varCodePage)
{
	CXAutoPtr<> bytePtr;
	UINT strLen = SysStringByteLen(strData);
	int nCodePage = varGetNumbar(varCodePage);

	if(nCodePage == -1)
		return m_File.Write(strData, strLen);

	if(nCodePage == 0)nCodePage = _AtlGetConversionACP();

	int _nTempCount = ::WideCharToMultiByte(nCodePage, 0, strData, strLen / 2, NULL, 0, NULL, NULL);
	bytePtr.Allocate(_nTempCount);

	::WideCharToMultiByte(nCodePage, 0, strData, strLen / 2, (LPSTR)bytePtr.m_p, _nTempCount, NULL, NULL);

	return m_File.Write(bytePtr.m_p, _nTempCount);
}

STDMETHODIMP CXBinaryFile::ReadAllText(VARIANT varCodePage, BSTR* pVal)
{
	long nSize = -1;

	CXVarPtr varPtr;
	ULONG nRead;
	HRESULT hr;

	if(nSize < 0)
	{
		DOUBLE nTotalSize;
		DOUBLE nPosition;

		hr = get_Size(&nTotalSize);
		if(SUCCEEDED(hr))
		{
			hr = get_Position(&nPosition);
			if(SUCCEEDED(hr))
			{
				if(nTotalSize - nPosition > 0x7fffffff)
					return DISP_E_OVERFLOW;
				nSize = (long)(nTotalSize - nPosition);
			}
		}
	}

	if(nSize < 0)
	{
		ULARGE_INTEGER cb = {0xffffffff, 0xffffffff};
		CXMemStream mStream;

		hr = mStream.CopyFrom(&m_File, cb.QuadPart, NULL, &cb.QuadPart);
		if(FAILED(hr))return hr;

		if(cb.QuadPart == 0)
			return CXObject::GetMessageFromError(ERROR_HANDLE_EOF);

		if(cb.HighPart != 0)
			return DISP_E_OVERFLOW;

		hr = varPtr.Create((UINT)mStream.GetLength());
		if(FAILED(hr))return hr;

		hr = m_File.Read(varPtr.m_pData, nSize, &nRead);
		if(FAILED(hr))return hr;
	}else
	{
		hr = varPtr.Create(nSize);
		if(FAILED(hr))return hr;

		hr = m_File.Read(varPtr.m_pData, nSize, &nRead);
		if(FAILED(hr))return hr;
	}

	int nCodePage = varGetNumbar(varCodePage, _AtlGetConversionACP());

	int _nTempCount = ::MultiByteToWideChar(nCodePage, 0, (LPCSTR)varPtr.m_pData, varPtr.m_nSize, NULL, NULL);
	*pVal = ::SysAllocStringLen(NULL, _nTempCount);
	if(*pVal != NULL)
		::MultiByteToWideChar(nCodePage, 0, (LPCSTR)varPtr.m_pData, varPtr.m_nSize, *pVal, _nTempCount);

	return S_OK;
}

STDMETHODIMP CXBinaryFile::Close(void)
{
	m_File.Close();

	return S_OK;
}

STDMETHODIMP CXBinaryFile::Truncate(void)
{
	return m_File.Truncate();
}

STDMETHODIMP CXBinaryFile::FlushBuffers(void)
{
	return m_File.FlushBuffers();
}
