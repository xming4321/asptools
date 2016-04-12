#pragma once

#include "atlib.h"

class CXBlockStream : public CXClass<IStream>
{
public:
	CXBlockStream(LPBYTE pData, DWORD dwSize) :
		m_pData(pData),
		m_dwSize(dwSize),
		m_pvarPtr(NULL),
		m_dwPosition(0),
		m_dwCount(0)
	{
	}

	CXBlockStream(void) :
		m_pData(NULL),
		m_dwSize(0),
		m_pvarPtr(NULL),
		m_dwPosition(0),
		m_dwCount(0)
	{
	}

	~CXBlockStream(void)
	{
		Clear();
	}

	HRESULT Attach(LPBYTE ptr, int size)
	{
		m_pData = ptr;
		m_dwSize = size;

		return S_OK;
	}

	HRESULT AttachVariant(VARIANT* pvar)
	{
		HRESULT hr;

		m_pvarPtr = new CXVarPtr();

		hr = m_pvarPtr->Attach(pvar);
		if(FAILED(hr))return hr;

		m_pData = m_pvarPtr->m_pData;
		m_dwSize = m_pvarPtr->m_nSize;

		return S_OK;
	}

	HRESULT AttachVariant(VARIANT& var)
	{
		return AttachVariant(&var);
	}

	void Clear(void)
	{
		m_pData = NULL;
		m_pBuffer.Free();
		m_dwSize = 0;
		m_dwPosition = 0;
		m_dwCount = 0;
		if(m_pvarPtr)
		{
			delete m_pvarPtr;
			m_pvarPtr = NULL;
		}
	}

	STDMETHOD(Read)(void *pDest, ULONG nMaxLen, ULONG *pnRead)
	{
		if(m_pData == NULL)
			return E_POINTER;

		DWORD dwToRead = min((m_dwSize - m_dwPosition), nMaxLen);

		memcpy(pDest, m_pData + m_dwPosition, dwToRead);
		m_dwPosition += dwToRead;
		if(m_dwCount < m_dwPosition)m_dwCount = m_dwPosition;

		if(pnRead != NULL)
			*pnRead = dwToRead;

		if(dwToRead < nMaxLen)
			return S_FALSE;

		return S_OK;
	}

	STDMETHOD(Write)(const void * pv, ULONG cb, ULONG * pcbWritten)
	{
		if(m_pData != NULL)
			return E_NOTIMPL;

		if(cb)
		{
			if(m_dwSize - m_dwPosition < (int)cb)
			{
				while(m_dwSize - m_dwPosition < (int)cb)
				{
					if(m_dwSize < 4096)m_dwSize += 1024;
					else m_dwSize += m_dwSize / 4;

					if(m_dwSize < 0)
						return DISP_E_OVERFLOW;
				}

				CXAutoPtr<BYTE> pBuffer;

				if(!pBuffer.Allocate(m_dwSize))
					return E_OUTOFMEMORY;

				if(m_dwPosition)
					CopyMemory(pBuffer, m_pBuffer, m_dwPosition);

				m_pBuffer = pBuffer;
			}

			CopyMemory(m_pBuffer + m_dwPosition, pv, cb);
			m_dwPosition += cb;
			if(m_dwCount < m_dwPosition)m_dwCount = m_dwPosition;
		}

		if(pcbWritten)*pcbWritten = cb;

		return S_OK;
	}


	HRESULT __stdcall Seek(LARGE_INTEGER /*dlibMove*/, DWORD /*dwOrigin*/, ULARGE_INTEGER * /*pLibNewPosition*/)
	{
		return E_NOTIMPL;
	}

	HRESULT __stdcall SetSize(ULARGE_INTEGER /*libNewSize*/)
	{
		return E_NOTIMPL;
	}

	HRESULT __stdcall CopyTo(IStream * /*pStream*/, ULARGE_INTEGER /*cb*/, ULARGE_INTEGER * /*pcbRead*/, ULARGE_INTEGER * /*pcbWritten*/)
	{
		return E_NOTIMPL;
	}

	HRESULT __stdcall Commit(DWORD /*grfCommitFlags*/)
	{
		return E_NOTIMPL;
	}

	HRESULT __stdcall Revert()
	{
		return E_NOTIMPL;
	}

	HRESULT __stdcall LockRegion(ULARGE_INTEGER /*libOffset*/, ULARGE_INTEGER /*cb*/, DWORD /*dwLockType*/)
	{
		return E_NOTIMPL;
	}

	HRESULT __stdcall UnlockRegion(ULARGE_INTEGER /*libOffset*/, ULARGE_INTEGER /*cb*/, DWORD /*dwLockType*/)
	{
		return E_NOTIMPL;
	}

	HRESULT __stdcall Stat(STATSTG * /*pstatstg*/, DWORD /*grfStatFlag*/)
	{
		return E_NOTIMPL;
	}

	HRESULT __stdcall Clone(IStream ** /*ppstm*/)
	{
		return E_NOTIMPL;
	}

public:
	CXAutoPtr<BYTE> m_pBuffer;
	DWORD	m_dwSize;
	DWORD	m_dwPosition;
	DWORD	m_dwCount;

private:
	LPBYTE	m_pData;
	CXVarPtr* m_pvarPtr;
};

class CXMemStream : public CXClass<IStream>
{
public:
	CXMemStream(void);

	enum SeekPosition { begin = 0x0, current = 0x1, end = 0x2 };

	ULONGLONG inline GetLength(void)
	{
		STATSTG ss;

		HRESULT hr = Stat(&ss, 0);
		if(FAILED(hr))return 0;

		return ss.cbSize.QuadPart;
	}

	ULONGLONG inline GetPosition(void)
	{
		return Seek(0, current);
	}

	ULONGLONG Seek(LONGLONG pos, UINT nFrom = begin)
	{
		LARGE_INTEGER lVal;
		ULARGE_INTEGER rVal = {0};

		lVal.QuadPart = pos;
		Seek(lVal, nFrom, &rVal);

		return rVal.QuadPart;
	}

	void SeekToBegin(void)
	{
		Seek(0);
	}

	ULONGLONG SeekToEnd(void)
	{
		return Seek(0, end);
	}

	HRESULT GetVariant(VARIANT* pvar)
	{
		CXVarPtr varPtr;

		HRESULT hr = varPtr.Create((ULONG)GetLength());
		if(FAILED(hr))return hr;

		SeekToBegin();
		Read(varPtr.m_pData, varPtr.m_nSize);

		return varPtr.GetVariant(pvar);
	}

	HRESULT GetVariant(VARIANT& var)
	{
		return GetVariant(&var);
	}

public:
	// ISequentialStream
	STDMETHOD(Read)(void *pv, ULONG cb, ULONG *pcbRead = NULL);
	STDMETHOD(Write)(const void *pv, ULONG cb, ULONG *pcbWritten = NULL);

public:
	// IStream
	STDMETHOD(Seek)(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER *plibNewPosition);
	STDMETHOD(SetSize)(ULARGE_INTEGER libNewSize);
	STDMETHOD(CopyTo)(IStream *pstm, ULARGE_INTEGER cb, ULARGE_INTEGER *pcbRead, ULARGE_INTEGER *pcbWritten);
	STDMETHOD(Commit)(DWORD grfCommitFlags);
	STDMETHOD(Revert)(void);
	STDMETHOD(LockRegion)(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType);
	STDMETHOD(UnlockRegion)(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType);
	STDMETHOD(Stat)(STATSTG *pstatstg, DWORD grfStatFlag);
	STDMETHOD(Clone)(IStream **ppstm);

public:
	HRESULT CopyFrom(IStream *pstm, unsigned __int64 cb = -1, unsigned __int64 *pcbRead = NULL, unsigned __int64 *pcbWritten = NULL);
	HRESULT CopyTo(IStream *pstm, unsigned __int64 cb = -1, unsigned __int64 *pcbRead = NULL, unsigned __int64 *pcbWritten = NULL);

private:
	CCriticalSection m_cs;
	ULONG m_ulPosition;
	CHandle m_hTempFile;

public:
	ULONG ReadAt(ULONG ulOffset, void* pv, ULONG cb);
	void WriteAt(ULONG ulOffset, void const* pv, ULONG cb);
	void SetSize(ULONG cb);
	ULONG GetSize(void)
	{
		return m_size;
	}

private:
	CAtlArray< CHeapPtr<BYTE> > m_pBuffers;
	ULONG m_size;
};
