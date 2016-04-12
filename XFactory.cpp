// XFactory.cpp : Implementation of CXFactory

#include "stdafx.h"
#include "XFactory.h"
#include "XDictionary.h"


// CXFactory


HRESULT CXFactory::CreateObject(BSTR objName, IDispatch** retVal)
{
	if(!_wcsicmp(objName, L"dictionary"))
	{
		CXComPtr<CXDictionary> pDict;

		pDict.CoCreateObject();
		*retVal = pDict.Detach();
	}

	return S_OK;
}
