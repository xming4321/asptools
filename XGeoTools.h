// XGeoTools.h : Declaration of the CXGeoTools

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"
#include "atlib.h"

// CXGeoTools

struct CityInfoStruct
{
	int id;
	LPCWSTR name;
	LPCWSTR sname;
	LPCWSTR ename;
	LPCWSTR path;
	int level;	// 0: 洲 1: 国家 2: 地区 3: 省 4: 城市 5: 区县
};

extern CityInfoStruct s_cityname[3731];
#define CITYCOUNT	(sizeof(s_cityname) / sizeof(s_cityname[0]))

inline int nextcity(int cityid)
{
	if(cityid == 142500000) // 重庆市
		return 142510000;

	if(cityid == 142499999) // 西南地区
		return 142600000;

	if(cityid % 100000000 == 0) // 洲
		return cityid + 100000000;

	if(cityid % 1000000 == 0) // 国家
		return cityid + 1000000;

	if(cityid % 100000 == 0) // 地区
		return cityid + 100000;

	if(cityid % 10000 == 0) // 省
		return cityid + 10000;

	if(cityid % 100 == 0) // 城市
		return cityid + 100;

	return cityid + 1;
}

inline int parentcity(int cityid)
{
	if(cityid == 142500000) // 重庆市
		return 142499999;

	if(cityid == 142499999) // 西南地区
		return 142000000;

	if(cityid % 100000000 == 0) // 洲
		return 0;

	if(cityid % 1000000 == 0) // 国家
		return cityid / 100000000 * 100000000;

	if(cityid % 100000 == 0) // 地区
		return cityid / 1000000 * 1000000;

	if(cityid % 10000 == 0) // 省
	{
		cityid = cityid / 100000 * 100000;
		if(cityid == 142500000) // 重庆市
			return 142499999; // 西南地区
		return cityid;
	}

	if(cityid % 100 == 0) // 城市
		return cityid / 10000 * 10000;

	return cityid / 100 * 100;
}

class ATL_NO_VTABLE CXGeoTools :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXGeoTools, &CLSID_XGeoTools>,
	public IDispatchImpl<IXGeoTools, &IID_IXGeoTools, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
private:

	class CXGeoPoint : public CXDispatch<IXGeoPoint>
	{
	public:
		CXGeoPoint()
		{
			m_latitude = 0;
			m_longitude = 0;
			m_Datium = 0;
		}

	public:
		STDMETHOD(get_Latitude)(double* pVal);
		STDMETHOD(put_Latitude)(double newVal);
		STDMETHOD(get_Longitude)(double* pVal);
		STDMETHOD(put_Longitude)(double newVal);
		STDMETHOD(get_Datium)(int* pVal);
		STDMETHOD(ConvertDatium)(int newVal);
		STDMETHOD(GetDistance)(IXGeoPoint* pTo, double* pVal);
		STDMETHOD(CityInfo)(VARIANT* pVal);

	public:
		double m_latitude;
		double m_longitude;
		int m_Datium;

	private:
		static short s_LPMap[451][2635][2];
	};

public:
	CXGeoTools()
	{
		m_pUnkMarshaler = NULL;
	}

DECLARE_REGISTRY_RESOURCEID(IDR_XGEOTOOLS)


BEGIN_COM_MAP(CXGeoTools)
	COM_INTERFACE_ENTRY(IXGeoTools)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
END_COM_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()
	DECLARE_GET_CONTROLLING_UNKNOWN()

	HRESULT FinalConstruct();
	void FinalRelease()
	{
		m_pUnkMarshaler.Release();
	}

	CComPtr<IUnknown> m_pUnkMarshaler;

public:
	STDMETHOD(Point)(int Datium, VARIANT Latitude, VARIANT Longitude, IXGeoPoint** pVal);

private:
	CCriticalSection m_cs;
};

OBJECT_ENTRY_AUTO(__uuidof(XGeoTools), CXGeoTools)
