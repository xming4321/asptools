// XGeoTools.cpp : Implementation of CXGeoTools

#include "stdafx.h"
#include "XGeoTools.h"

extern CXStringA s_pathRuntime;

// CXGeoTools

template <typename T>
void CSVParser(const T *str, CAtlArray< CStringT< T, StrTraitATL< T, ChTraitsCRT< T > > > >& arrayResult)
{
	const T *str1;
	CStringT< T, StrTraitATL< T, ChTraitsCRT< T > > > strData;

	while(*str)
	{
		str1 = str;

		while(*str && *str != ',')str ++;

		if(*str1 == '\"')
		{
			str1 ++;
			if(*(str-1) == '\"')
				strData.SetString(str1, str - str1 - 1);
			else strData.SetString(str1, str - str1);
		}else strData.SetString(str1, str - str1);

		arrayResult.Add(strData);

		if(*str)
		{
			str ++;
			if(!*str)
			{
				arrayResult.Add("");
				break;
			}
		}
	}
}

HRESULT CXGeoTools::FinalConstruct()
{
	return CoCreateFreeThreadedMarshaler(
		GetControllingUnknown(), &m_pUnkMarshaler.p);
}

STDMETHODIMP CXGeoTools::Point(int Datium, VARIANT Latitude, VARIANT Longitude, IXGeoPoint** pVal)
{
	CXComPtr<CXGeoPoint> pPoint;

	pPoint.CreateObject();
	pPoint->m_Datium = Datium;
	pPoint->m_latitude = varGetFloat(Latitude);
	pPoint->m_longitude = varGetFloat(Longitude);

	if(pVal)*pVal = pPoint.Detach();

	return S_OK;
}

