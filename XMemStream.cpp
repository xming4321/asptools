#include "StdAfx.h"
#include "XMemStream.h"

#define LB_BLOCKSIZE	4096

void CXMemStream::SetSize(ULONG cb)
{
	int i;

	m_pBuffers.SetCount((cb + LB_BLOCKSIZE - 1) / LB_BLOCKSIZE);

	if(m_size < cb)
		for(i = (m_size + LB_BLOCKSIZE - 1) / LB_BLOCKSIZE; i < (int)m_pBuffers.GetCount(); i ++)
			m_pBuffers[i].Allocate(LB_BLOCKSIZE);

	m_size = cb;
}

ULONG CXMemStream::ReadAt(ULONG ulOffset, void* pv, ULONG cb)
{
	if(ulOffset >= m_size)
		return 0;

	int i, pos, size, cbRead;
	BYTE* ptr = (BYTE*)pv;

	cb += ulOffset;
	if(cb > m_size)
		cb = m_size;

	cbRead = cb - ulOffset;

	while(ulOffset < cb)
	{
		i = ulOffset / LB_BLOCKSIZE;
		pos = ulOffset % LB_BLOCKSIZE;
		size = cb - ulOffset;
		if(size + pos > LB_BLOCKSIZE)
			size = LB_BLOCKSIZE - pos;

		CopyMemory(ptr, m_pBuffers[i] + pos, size);

		ptr += size;
		ulOffset += size;
	}

	return cbRead;
}

void CXMemStream::WriteAt(ULONG ulOffset, void const* pv, ULONG cb)
{
	int i, pos, size;
	const BYTE* ptr = (const BYTE*)pv;

	if(ulOffset + cb > m_size)
	{
		m_pBuffers.SetCount((ulOffset + cb + LB_BLOCKSIZE - 1) / LB_BLOCKSIZE);

		for(i = (m_size + LB_BLOCKSIZE - 1) / LB_BLOCKSIZE; i < (int)m_pBuffers.GetCount(); i ++)
			m_pBuffers[i].Allocate(LB_BLOCKSIZE);

		m_size = ulOffset + cb;
	}

	cb += ulOffset;
	while(ulOffset < cb)
	{
		i = ulOffset / LB_BLOCKSIZE;
		pos = ulOffset % LB_BLOCKSIZE;
		size = cb - ulOffset;
		if(size + pos > LB_BLOCKSIZE)
			size = LB_BLOCKSIZE - pos;

		CopyMemory(m_pBuffers[i] + pos, ptr, size);

		ptr += size;
		ulOffset += size;
	}
}

#define MAX_MEMSIZE		0x2000000

CXMemStream::CXMemStream(void) :
	m_ulPosition(0),
	m_size(0)
{}

// ISequentialStream
STDMETHODIMP CXMemStream::Read(void *pv, ULONG cb, ULONG *pcbRead)
{
	if(pv == NULL)
		return STG_E_INVALIDPOINTER;

	ULONG n, cbRead;

	if(pcbRead == NULL)pcbRead = &cbRead;
	*pcbRead = 0;
	if(!cb)return S_OK;

	m_cs.Enter();

	if(m_ulPosition < MAX_MEMSIZE)
	{
		*pcbRead = n = ReadAt(m_ulPosition, pv, cb);
		m_ulPosition += n;

		cb -= n;
		pv = (void *)((char*)pv + n);
	}

	if(m_ulPosition >= MAX_MEMSIZE && cb && m_hTempFile)
	{
		if(SetFilePointer(m_hTempFile, m_ulPosition - MAX_MEMSIZE, NULL, FILE_BEGIN) == INVALID_SET_FILE_POINTER)
		{
			m_cs.Leave();
			return GetLastError();
		}

		ULONG cbRead1 = 0;

		if(!::ReadFile(m_hTempFile, pv, cb, &cbRead1, NULL))
		{
			m_cs.Leave();
			return GetLastError();
		}

		m_ulPosition += cbRead1;
		*pcbRead += cbRead1;
	}

	m_cs.Leave();

	if(!*pcbRead)
		return GetMessageFromError(ERROR_HANDLE_EOF);

	if(*pcbRead < cb)return S_FALSE;

	return S_OK;
}

STDMETHODIMP CXMemStream::Write(const void *pv, ULONG cb, ULONG *pcbWritten)
{
	if(pv == NULL)
		return STG_E_INVALIDPOINTER;

	ULONG cbWritten;

	if(pcbWritten == NULL)pcbWritten = &cbWritten;
	*pcbWritten = 0;

	m_cs.Enter();

	if(m_ulPosition < MAX_MEMSIZE)
	{
		ULONG cb1 = cb;

		if(MAX_MEMSIZE - m_ulPosition < cb1)
			cb1 = MAX_MEMSIZE - m_ulPosition;

		WriteAt(m_ulPosition, pv, cb1);
		m_ulPosition += cb1;
		*pcbWritten = cb1;

		cb -= cb1;
		pv = (const void *)((const char*)pv + cb1);
	}

	if(m_ulPosition >= MAX_MEMSIZE && cb)
	{
		if(!m_hTempFile)
		{
			char szTempPath[MAX_PATH];
			char szTempFile[MAX_PATH];

			::GetTempPathA(MAX_PATH, szTempPath);
			if(!::GetTempFileNameA(szTempPath, "~", 0, szTempFile))
			{
				::GetSystemDirectoryA(szTempPath, MAX_PATH);
				::GetTempFileNameA(szTempPath, "~", 0, szTempFile);
			}

			m_hTempFile.m_h = ::CreateFileA(szTempFile, GENERIC_READ | GENERIC_WRITE, 0, NULL,
				OPEN_EXISTING, FILE_FLAG_DELETE_ON_CLOSE, 0);

			if (m_hTempFile == INVALID_HANDLE_VALUE)
			{
				m_hTempFile.m_h = NULL;
				m_cs.Leave();

				return GetLastError();
			}

			SetSize(MAX_MEMSIZE);
		}

		if(SetFilePointer(m_hTempFile, m_ulPosition - MAX_MEMSIZE, NULL, FILE_BEGIN) == INVALID_SET_FILE_POINTER)
		{
			m_cs.Leave();
			return GetLastError();
		}

		ULONG cbWritten1 = 0;

		if(!::WriteFile(m_hTempFile, pv, cb, &cbWritten1, NULL))
		{
			m_cs.Leave();
			return GetLastError();
		}

		m_ulPosition += cbWritten1;
		*pcbWritten += cbWritten1;
	}

	m_cs.Leave();

	return S_OK;
}

// IStream
STDMETHODIMP CXMemStream::Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER *plibNewPosition)
{
	m_cs.Enter();

	if(dwOrigin == SEEK_CUR)
		dlibMove.QuadPart += m_ulPosition;
	else if(dwOrigin == SEEK_END)
	{
		dlibMove.QuadPart += GetSize();
		if(m_hTempFile != NULL)
		{
			ULARGE_INTEGER cbSize;

			cbSize.LowPart = GetFileSize(m_hTempFile, &cbSize.HighPart);
			if(cbSize.LowPart == INVALID_FILE_SIZE)
			{
				HRESULT hr = GetLastError();
				if(FAILED(hr))
				{
					m_cs.Leave();
					return hr;
				}
			}

			dlibMove.QuadPart += cbSize.QuadPart;
		}
	}

	if(dlibMove.HighPart)
	{
		m_cs.Leave();
		return DISP_E_OVERFLOW;
	}
	m_ulPosition = dlibMove.LowPart;

	m_cs.Leave();

	if(plibNewPosition)
		plibNewPosition->QuadPart = dlibMove.LowPart;

	return S_OK;
}

STDMETHODIMP CXMemStream::SetSize(ULARGE_INTEGER libNewSize)
{
	if(libNewSize.HighPart)
		return DISP_E_OVERFLOW;

	m_cs.Enter();
	if(libNewSize.LowPart <= MAX_MEMSIZE)
	{
		SetSize(libNewSize.LowPart);
		m_hTempFile.Close();
	}else
	{
		SetSize(MAX_MEMSIZE);

		if(SetFilePointer(m_hTempFile, libNewSize.LowPart - MAX_MEMSIZE, NULL, FILE_BEGIN) == INVALID_SET_FILE_POINTER)
		{
			m_cs.Leave();
			return GetLastError();
		}

		if(!SetEndOfFile(m_hTempFile))
		{
			m_cs.Leave();
			return GetLastError();
		}
	}
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXMemStream::CopyTo(IStream *pstm, ULARGE_INTEGER cb, ULARGE_INTEGER *pcbRead, ULARGE_INTEGER *pcbWritten)
{
	return E_NOTIMPL;
}

STDMETHODIMP CXMemStream::Commit(DWORD grfCommitFlags)
{
	return E_NOTIMPL;
}

STDMETHODIMP CXMemStream::Revert(void)
{
	return E_NOTIMPL;
}

STDMETHODIMP CXMemStream::LockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType)
{
	return E_NOTIMPL;
}

STDMETHODIMP CXMemStream::UnlockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType)
{
	return E_NOTIMPL;
}

STDMETHODIMP CXMemStream::Stat(STATSTG *pstatstg, DWORD grfStatFlag)
{
	ZeroMemory(pstatstg, sizeof(STATSTG));

	pstatstg->type = STGTY_LOCKBYTES;

	m_cs.Enter();
	pstatstg->cbSize.LowPart = GetSize();

	if(m_hTempFile)
	{
		LARGE_INTEGER cbSize = {0};

		cbSize.LowPart = SetFilePointer(m_hTempFile, 0, &cbSize.HighPart, FILE_END);

		if(cbSize.LowPart == INVALID_SET_FILE_POINTER)
		{
			HRESULT hr = GetLastError();
			if(FAILED(hr))
			{
				m_cs.Leave();
				return hr;
			}
		}

		pstatstg->cbSize.QuadPart += cbSize.QuadPart;
	}
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXMemStream::Clone(IStream **ppstm)
{
	return E_NOTIMPL;
}

#define COPY_BLOCK  65536

static HRESULT CopyStream(IStream* pdst, IStream* psrc, unsigned __int64 cb, unsigned __int64 *pcbRead, unsigned __int64 *pcbWritten)
{
	if(pdst == NULL || psrc == NULL)
		return STG_E_INVALIDPOINTER;

	__int64 cbRead = 0, cbWritten = 0;
	CXAutoPtr<> copyBuff;
	ULONG cbLeft;
	HRESULT hr = S_OK;

	copyBuff.Allocate(COPY_BLOCK);

	while(cb)
	{
		cbLeft = (ULONG)(cb > COPY_BLOCK ? COPY_BLOCK : cb);

		hr = psrc->Read(copyBuff, cbLeft, &cbLeft);
		if (FAILED(hr) || cbLeft == 0)break;

		cbRead += cbLeft;
		cb -= cbLeft;

		hr = pdst->Write(copyBuff, cbLeft, &cbLeft);
		if (FAILED(hr))break;
		cbWritten += cbLeft;
	}

	if(pcbRead != NULL)*pcbRead = cbRead;
	if(pcbWritten != NULL)*pcbWritten = cbWritten;

	if(hr == HRESULT_FROM_WIN32(ERROR_HANDLE_EOF))
		return S_FALSE;

	return hr;
}

HRESULT CXMemStream::CopyFrom(IStream *pstm, unsigned __int64 cb, unsigned __int64 *pcbRead, unsigned __int64 *pcbWritten)
{
	return CopyStream(this, pstm, cb, pcbRead, pcbWritten);
}

HRESULT CXMemStream::CopyTo(IStream *pstm, unsigned __int64 cb, unsigned __int64 *pcbRead, unsigned __int64 *pcbWritten)
{
	return CopyStream(pstm, this, cb, pcbRead, pcbWritten);
}

