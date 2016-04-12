#pragma once

#include "atlib.h"
#include "zlib\zlib.h"
#include "zlib\zutil.h"

class CXFileStream : public CXClass<IStream>
{
public:
	enum OpenFlags
	{
		modeRead =			(int) 0x00001,
		modeWrite =			(int) 0x00002,
		modeReadWrite =		(int) 0x00003,
		modeAppend =		(int) 0x00008,
		modeCreate =		(int) 0x00010,

		shareRead =			(int) 0x00001,
		shareWrite =		(int) 0x00002,
		shareReadWrite =	(int) 0x00003,
		shareDelete =		(int) 0x00004,
		shareAll =			(int) 0x00007
	};

	enum SeekPosition { begin = 0x0, current = 0x1, end = 0x2 };

	CXFileStream() : m_bCompressMode(0)
	{}

	~CXFileStream()
	{
		Close();
	}

public:
	HRESULT Create(BSTR bstrName, VARIANT_BOOL bOverwrite = -1);
	HRESULT Open(BSTR bstrName, short nMode = modeRead, short nShare = shareReadWrite);
	HRESULT Close(void);

	ULONGLONG inline GetLength(void)
	{
		ULARGE_INTEGER lVal;

		lVal.LowPart = GetFileSize(m_hFile, &lVal.HighPart);
		return lVal.QuadPart;
	}

	void inline SetSize(ULONGLONG size)
	{
		ULARGE_INTEGER lVal;

		lVal.QuadPart = size;
		SetSize(lVal);
	}

	ULONGLONG inline GetPosition(void)
	{
		return Seek(0, current);
	}

	ULONGLONG inline Seek(LONGLONG pos, UINT nFrom = begin)
	{
		LARGE_INTEGER lVal;
		ULARGE_INTEGER rVal = {0};

		lVal.QuadPart = pos;
		Seek(lVal, nFrom, &rVal);

		return rVal.QuadPart;
	}

	void inline SeekToBegin(void)
	{
		Seek(0);
	}

	ULONGLONG inline SeekToEnd(void)
	{
		return Seek(0, end);
	}

	BOOL inline isEOF()
	{
		return GetPosition() >= GetLength();
	}

	DATE GetlastModify()
	{
		FILETIME ft, ft1;
		SYSTEMTIME st;
		DATE d;

		GetFileTime(m_hFile, NULL, NULL, &ft);
		FileTimeToLocalFileTime(&ft, &ft1);
		FileTimeToSystemTime(&ft1, &st);
		SystemTimeToVariantTime(&st, &d);

		return d;
	}

	HRESULT SetCompress(BOOL bZip);
	BOOL GetCompress();

	HRESULT FlushBuffers();
	HRESULT Truncate();

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

private:
	HRESULT OpenFile(BSTR pstrName,  DWORD dwDesiredAccess, DWORD dwShareMode, DWORD dwCreationDisposition);

	CHandle m_hFile;

private:
	z_stream m_streamCompress;
	CXAutoPtr<> m_bufCompress;
	BOOL m_bCompressMode;
};
