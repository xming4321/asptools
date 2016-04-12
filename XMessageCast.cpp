// XMessageCast.cpp : Implementation of CXMessageCast

#include "stdafx.h"
#include "XMessageCast.h"


// CXMessageCast


/*
	"server", 123								TCP Server
	"192.168.100.0", 123						UDP ¹ã²¥
	"192.168.100.0", "123.32.32.32", 123		UDP ¶à²¥
*/

HRESULT CXMessageCast::Open(BSTR v1, VARIANT v2, VARIANT v3)
{
	WSADATA wsd;
	WSAStartup(MAKEWORD(2,2),&wsd);

	SOCKADDR_IN s_dest;
	CXStringA str(v1);

	s_dest.sin_family = AF_INET;
	s_dest.sin_addr.s_addr = inet_addr(str);

	if(s_dest.sin_addr.s_addr == INADDR_NONE)
	{
		LPHOSTENT lphost;
		lphost = gethostbyname(str);
		if (lphost == NULL)return CXObject::WSAGetLastError();

		s_dest.sin_addr.s_addr = ((LPIN_ADDR)lphost->h_addr)->s_addr;
	}

	if(s_dest.sin_addr.S_un.S_un_b.s_b4 == 0)
	{
		return S_OK;



	}else
	{
		return S_OK;



	}

	return S_OK;
}

HRESULT CXMessageCast::Close()
{

	return S_OK;
}

void CXMessageCast::Send(const void* buffer, int size)
{
	m_csOut.Enter();

	m_OutBuffers.Add();
	m_OutBuffers[m_OutBuffers.GetCount() - 1].SetCount(size);
	memcpy(&m_OutBuffers[m_OutBuffers.GetCount() - 1][0], buffer, size);

	m_csOut.Leave();
}

HRESULT CXMessageCast::Recv(VARIANT* pVal)
{


	return S_OK;
}

