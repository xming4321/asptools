#include "StdAfx.h"
#include "XFileStream.h"

#ifndef Z_BUFSIZE
#define Z_BUFSIZE 16384
#endif

HRESULT CXFileStream::SetCompress(BOOL bZip)
{
	HRESULT hr = S_OK;

	if(!((m_bufCompress.m_p != NULL) ^ bZip))
		return S_OK;

	if(bZip)
	{
		ZeroMemory(&m_streamCompress, sizeof(m_streamCompress));
		m_bufCompress.Allocate(Z_BUFSIZE);
	}else if(m_bCompressMode)
	{
		if(m_bCompressMode == 2)
		{
			int err;
			uInt len;
			int done = 0;

			for (;;)
			{
				len = Z_BUFSIZE - m_streamCompress.avail_out;

				if (len != 0)
				{
					DWORD cbWritten;

					if(!::WriteFile(m_hFile, m_bufCompress, len, &cbWritten, NULL))
					{
						hr = GetLastError();
						break;
					}

					m_streamCompress.next_out = m_bufCompress;
					m_streamCompress.avail_out = Z_BUFSIZE;
				}
				if (done) break;
				err = deflate(&m_streamCompress, Z_FINISH);

				if (len == 0 && err == Z_BUF_ERROR)err = Z_OK;

				done = (m_streamCompress.avail_out != 0 || err == Z_STREAM_END);

				if (err != Z_OK && err != Z_STREAM_END)
				{
					hr = SetErrorInfo(zError(err));
					break;
				}
			}

			deflateEnd(&m_streamCompress);
		}else
		{
			inflateEnd(&m_streamCompress);
		}

		m_bufCompress.Free();

		m_bCompressMode = 0;
	}

	return hr;
}

BOOL CXFileStream::GetCompress()
{
	return m_bufCompress.m_p != NULL;
}

HRESULT CXFileStream::FlushBuffers()
{
	if(m_hFile.m_h == NULL)
		return E_HANDLE;

	if(m_bCompressMode != 0)
		return E_NOTIMPL;

	if(!FlushFileBuffers(m_hFile.m_h))
		return GetLastError();

	return S_OK;
}

// ISequentialStream
STDMETHODIMP CXFileStream::Read(void *pv, ULONG cb, ULONG *pcbRead)
{
	if(pv == NULL)
		return STG_E_INVALIDPOINTER;

	if(m_hFile.m_h == NULL)
		return E_HANDLE;

	if(m_bCompressMode == 2)
		return E_NOTIMPL;

	ULONG cbRead;

	if(pcbRead == NULL)pcbRead = &cbRead;

	*pcbRead = 0;

	if(m_bufCompress)
	{
		int err = Z_OK;

		if(m_bCompressMode == 0)
		{
			err = inflateInit2(&m_streamCompress, -MAX_WBITS);
			if (err != Z_OK)return SetErrorInfo(zError(err));

			m_streamCompress.next_in = m_bufCompress;
			m_streamCompress.avail_in = 0;

			m_bCompressMode = 1;
		}

		m_streamCompress.next_out = (Bytef*)pv;
		m_streamCompress.avail_out = cb;

		while (err != Z_STREAM_END && err != Z_BUF_ERROR)
		{
			if(m_streamCompress.avail_in == 0)
			{
				if(!::ReadFile(m_hFile, m_bufCompress, Z_BUFSIZE, &cbRead, NULL))
					return GetLastError();

				m_streamCompress.next_in = m_bufCompress;
				m_streamCompress.avail_in = cbRead;
			}

			if (m_streamCompress.avail_out == 0)
				break;

			err = inflate(&m_streamCompress, Z_NO_FLUSH);
			if(err != Z_OK && err != Z_STREAM_END && err != Z_BUF_ERROR)
				return SetErrorInfo(zError(err));
		}

		*pcbRead = cb - m_streamCompress.avail_out;

		return S_OK;
	}

	if(!::ReadFile(m_hFile, pv, cb, pcbRead, NULL))
		return GetLastError();

	if(!*pcbRead)
		return GetMessageFromError(ERROR_HANDLE_EOF);

	if(*pcbRead < cb)return S_FALSE;

	return S_OK;
}

STDMETHODIMP CXFileStream::Write(const void *pv, ULONG cb, ULONG *pcbWritten)
{
	if(pv == NULL)
		return STG_E_INVALIDPOINTER;

	if(m_hFile.m_h == NULL)
		return E_HANDLE;

	if(m_bCompressMode == 1)
		return E_NOTIMPL;

	ULONG cbWritten = 0;

	if(pcbWritten == NULL)pcbWritten = &cbWritten;

	if(m_bufCompress)
	{
		int err = Z_OK;

		if(m_bCompressMode == 0)
		{
			err = deflateInit2(&m_streamCompress, Z_DEFAULT_COMPRESSION, Z_DEFLATED, -MAX_WBITS, DEF_MEM_LEVEL, Z_DEFAULT_STRATEGY);
			if (err != Z_OK)return SetErrorInfo(zError(err));

			m_streamCompress.next_out = m_bufCompress;
			m_streamCompress.avail_out = Z_BUFSIZE;

			m_bCompressMode = 2;
		}

		m_streamCompress.next_in = (Bytef*)pv;
		m_streamCompress.avail_in = cb;

		while (m_streamCompress.avail_in != 0)
		{
			if (m_streamCompress.avail_out == 0)
			{
				if(!::WriteFile(m_hFile, m_bufCompress, Z_BUFSIZE, pcbWritten, NULL))
					return GetLastError();

				m_streamCompress.next_out = m_bufCompress;
				m_streamCompress.avail_out = Z_BUFSIZE;
			}

			if (deflate(&m_streamCompress, Z_NO_FLUSH) != Z_OK)
				return SetErrorInfo(zError(err));
		}

		*pcbWritten = cb;

		return S_OK;
	}

	if(!::WriteFile(m_hFile, pv, cb, pcbWritten, NULL))
		return GetLastError();

	return S_OK;
}

// IStream
STDMETHODIMP CXFileStream::Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER *plibNewPosition)
{
	if(m_hFile.m_h == NULL)
		return E_HANDLE;

	if(m_bufCompress)
		return E_NOTIMPL;

	dlibMove.LowPart = SetFilePointer(m_hFile, dlibMove.LowPart, &dlibMove.HighPart, dwOrigin);

	if(dlibMove.LowPart == INVALID_SET_FILE_POINTER)
	{
		HRESULT hr = GetLastError();
		if(FAILED(hr))return hr;
	}

	if(plibNewPosition != NULL)
		*(LARGE_INTEGER*)plibNewPosition = dlibMove;

	return S_OK;
}

HRESULT CXFileStream::Truncate()
{
	if(m_hFile.m_h == NULL)
		return E_HANDLE;

	if(m_bufCompress)
		return E_NOTIMPL;

	if(!SetEndOfFile(m_hFile))
		return GetLastError();

	return S_OK;
}

STDMETHODIMP CXFileStream::SetSize(ULARGE_INTEGER libNewSize)
{
	if(m_hFile.m_h == NULL)
		return E_HANDLE;

	if(m_bufCompress)
		return E_NOTIMPL;

	LARGE_INTEGER lVal = {0};
	HRESULT hr;

	hr = Seek(lVal, SEEK_CUR, (ULARGE_INTEGER*)&lVal);
	if(FAILED(hr))return hr;

	hr = Seek(*(LARGE_INTEGER*)&libNewSize, SEEK_SET, NULL);
	if(FAILED(hr))return hr;

	if(!SetEndOfFile(m_hFile))
		return GetLastError();

	return Seek(lVal, SEEK_SET, NULL);
}

STDMETHODIMP CXFileStream::CopyTo(IStream *pstm, ULARGE_INTEGER cb, ULARGE_INTEGER *pcbRead, ULARGE_INTEGER *pcbWritten)
{
	return E_NOTIMPL;
}

STDMETHODIMP CXFileStream::Commit(DWORD grfCommitFlags)
{
	return E_NOTIMPL;
}

STDMETHODIMP CXFileStream::Revert(void)
{
	return E_NOTIMPL;
}

STDMETHODIMP CXFileStream::LockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType)
{
	if(m_hFile.m_h == NULL)
		return E_HANDLE;

	if(m_bufCompress)
		return E_NOTIMPL;

	if (!::LockFile(m_hFile, libOffset.LowPart, libOffset.HighPart, cb.LowPart, cb.HighPart))
		return GetLastError();

	return S_OK;
}

STDMETHODIMP CXFileStream::UnlockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType)
{
	if(m_hFile.m_h == NULL)
		return E_HANDLE;

	if(m_bufCompress)
		return E_NOTIMPL;

	if (!::UnlockFile(m_hFile, libOffset.LowPart, libOffset.HighPart, cb.LowPart, cb.HighPart))
		return GetLastError();

	return S_OK;
}

STDMETHODIMP CXFileStream::Stat(STATSTG *pstatstg, DWORD grfStatFlag)
{
	if(pstatstg == NULL)
		return STG_E_INVALIDPOINTER;

	if(m_hFile.m_h == NULL)
		return E_HANDLE;

	if(m_bufCompress)
		return E_NOTIMPL;

	ZeroMemory(pstatstg, sizeof(STATSTG));
	pstatstg->type = STGTY_STORAGE;
	if(!GetFileTime(m_hFile, &pstatstg->ctime, &pstatstg->atime, &pstatstg->mtime))
		return GetLastError();

	pstatstg->cbSize.LowPart = GetFileSize(m_hFile, &pstatstg->cbSize.HighPart);
	if(pstatstg->cbSize.LowPart == INVALID_FILE_SIZE)
	{
		HRESULT hr = GetLastError();
		if(FAILED(hr))return hr;
	}

	return S_OK;
}

STDMETHODIMP CXFileStream::Clone(IStream **ppstm)
{
	return E_NOTIMPL;
}

HRESULT CXFileStream::Close(void)
{
	SetCompress(FALSE);
	m_hFile.Close();
	return S_OK;
}

HRESULT CXFileStream::OpenFile(BSTR bstrName,  DWORD dwDesiredAccess, DWORD dwShareMode, DWORD dwCreationDisposition)
{
	Close();

	m_hFile.m_h = ::CreateFileW(bstrName, dwDesiredAccess, dwShareMode, NULL, dwCreationDisposition, FILE_ATTRIBUTE_NORMAL, 0);

	if (m_hFile == INVALID_HANDLE_VALUE)
	{
		HRESULT hr = GetLastError();

		m_hFile.m_h = NULL;

		if((hr == HRESULT_FROM_WIN32(ERROR_PATH_NOT_FOUND)) && (dwCreationDisposition != OPEN_EXISTING))
		{
			CXString str;
			int p = 3;

			str = bstrName;
			do
			{
				p = str.Find('\\', p);
				if(p == -1)break;
				CreateDirectoryW(str.Left(p), NULL);
				p ++;
			}while(TRUE);

			m_hFile.m_h = ::CreateFileW(bstrName, dwDesiredAccess, dwShareMode, NULL, dwCreationDisposition, FILE_ATTRIBUTE_NORMAL, 0);

			if (m_hFile == INVALID_HANDLE_VALUE)
			{
				m_hFile.m_h = NULL;
				return GetLastError();
			}

			return S_OK;
		}else return hr;
	}

	return S_OK;
}

HRESULT CXFileStream::Create(BSTR bstrName, VARIANT_BOOL bOverwrite)
{
	return OpenFile(bstrName,  GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ, bOverwrite ? CREATE_ALWAYS : CREATE_NEW);
}

HRESULT CXFileStream::Open(BSTR bstrName, short nMode, short nShare)
{
	static DWORD modes[] = {0, GENERIC_READ, GENERIC_WRITE, GENERIC_READ | GENERIC_WRITE};

	HRESULT hr = OpenFile(bstrName,  modes[nMode & modeReadWrite], nShare, nMode & modeCreate ? OPEN_ALWAYS : OPEN_EXISTING);
	if(FAILED(hr))return hr;

	if(nMode & modeAppend)
		::SetFilePointer(m_hFile, 0, NULL, FILE_END);

	return S_OK;
}

