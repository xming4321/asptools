// XMessageCast.h : Declaration of the CXMessageCast

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"
#include "atlib.h"


// CXMessageCast

class CXMessageCast
{
public:
	CXMessageCast()
	{
		m_id = GetLocalFileTime();
	}

public:
	HRESULT Open(BSTR v1, VARIANT v2, VARIANT v3);
	HRESULT Close();
	void Send(const void* buffer, int size);
	HRESULT Recv(VARIANT* pVal);

private:
	__int64 m_id;
	CCriticalSection m_csOut;
	CAtlArray< CAtlArray<BYTE> > m_OutBuffers;
};
