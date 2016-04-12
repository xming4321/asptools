// asptools.cpp : Implementation of DLL Exports.

#include "stdafx.h"
#include "asptools.h"
#include "resource.h"
#include "atlib.h"
#include "XFolderMan.h"

CLSID CLSID_FreeThreadedMarshaler;
CXStringA s_pathRuntime;

class CasptoolsModule : public CAtlDllModuleT< CasptoolsModule >
{
public :
	DECLARE_LIBID(LIBID_asptoolsLib)
	DECLARE_REGISTRY_APPID_RESOURCEID(IDR_ASPTOOLS, "{B3DC402F-88F9-4C9F-83B6-A9F075603538}")

public:
	CasptoolsModule()
	{
		HRESULT hr;
		CComPtr<IUnknown> pUnkMarshaler;
		CXObject obj;

		hr = CoCreateFreeThreadedMarshaler(&obj, &pUnkMarshaler.p);
		if(FAILED(hr))return;

		CComQIPtr<IMarshal> pMarshal;

		if(pMarshal = pUnkMarshaler)
			pMarshal->GetUnmarshalClass(IID_IUnknown, &obj, MSHCTX_INPROC, NULL, MSHLFLAGS_NORMAL, &CLSID_FreeThreadedMarshaler);
	}

	~CasptoolsModule()
	{
#ifdef _DEBUG
		_CrtDumpMemoryLeaks();
#endif
	}

public:
	HINSTANCE m_hInstance;
};

CasptoolsModule _AtlModule;


// DLL Entry Point
extern "C" BOOL WINAPI DllMain(HINSTANCE hInstance, DWORD dwReason, LPVOID lpReserved)
{
	_AtlModule.m_hInstance = hInstance;
	if(s_pathRuntime.IsEmpty())
	{
		s_pathRuntime.ReleaseBuffer(::GetModuleFileName(hInstance, s_pathRuntime.GetBuffer(2048), 2048));
		s_pathRuntime = s_pathRuntime.Left(s_pathRuntime.ReverseFind('\\') + 1);
	}

	return _AtlModule.DllMain(dwReason, lpReserved); 
}


// Used to determine whether the DLL can be unloaded by OLE
STDAPI DllCanUnloadNow(void)
{
    return _AtlModule.DllCanUnloadNow();
}


// Returns a class factory to create an object of the requested type
STDAPI DllGetClassObject(REFCLSID rclsid, REFIID riid, LPVOID* ppv)
{
    return _AtlModule.DllGetClassObject(rclsid, riid, ppv);
}


// DllRegisterServer - Adds entries to the system registry
STDAPI DllRegisterServer(void)
{
	// registers object, typelib and all interfaces in typelib
    HRESULT hr = _AtlModule.DllRegisterServer();

	TCHAR szModule1[MAX_PATH];
	TCHAR szModule2[MAX_PATH];
	TCHAR* pszFileName;

	DWORD dwFLen = ::GetModuleFileName(_AtlModule.m_hInstance, szModule1, MAX_PATH);
	if ( dwFLen != 0 && dwFLen != MAX_PATH )
	{
		dwFLen = ::GetFullPathName(szModule1, MAX_PATH, szModule2, &pszFileName);
		if (dwFLen != 0)
		{
			HKEY hk;
			static char str[] = "SYSTEM\\CurrentControlSet\\Services\\Eventlog\\Application\\asptools";

			if (RegOpenKey(HKEY_LOCAL_MACHINE, str, &hk ) == ERROR_SUCCESS ||
				RegCreateKey(HKEY_LOCAL_MACHINE, str, &hk) == ERROR_SUCCESS)
				RegSetValueEx(hk, _T("EventMessageFile"), 0, REG_SZ, (CONST BYTE*)(LPCTSTR)szModule2, (dwFLen + 1) * sizeof(TCHAR));

			RegCloseKey(hk);
		}
	}

	return hr;
}


// DllUnregisterServer - Removes entries from the system registry
STDAPI DllUnregisterServer(void)
{
	HRESULT hr = _AtlModule.DllUnregisterServer();
	return hr;
}

CXTypeInfoHolder CXDispatch<IXDocItem>::_tih(&IID_IXDocItem);
CXTypeInfoHolder CXDispatch<IXField>::_tih(&IID_IXField);
CXTypeInfoHolder CXDispatch<IXRecord>::_tih(&IID_IXRecord);
CXTypeInfoHolder CXDispatch<IXRecords>::_tih(&IID_IXRecords);
CXTypeInfoHolder CXDispatch<IXList>::_tih(&IID_IXList);
CXTypeInfoHolder CXDispatch<IXDocContent>::_tih(&IID_IXDocContent);
CXTypeInfoHolder CXDispatch<IXUploadData>::_tih(&IID_IXUploadData);
CXTypeInfoHolder CXDispatch<IXUploadList>::_tih(&IID_IXUploadList);
CXTypeInfoHolder CXDispatch<IXChannel>::_tih(&IID_IXChannel);
CXTypeInfoHolder CXDispatch<IXSession>::_tih(&IID_IXSession);
CXTypeInfoHolder CXDispatch<IXClass>::_tih(&IID_IXClass);
CXTypeInfoHolder CXDispatch<IXUrl>::_tih(&IID_IXUrl);
CXTypeInfoHolder CXDispatch<IXFile>::_tih(&IID_IXFile);
CXTypeInfoHolder CXDispatch<IXPathBand>::_tih(&IID_IXPathBand);
CXTypeInfoHolder CXDispatch<IXGeoPoint>::_tih(&IID_IXGeoPoint);


#include "asptools_i.c"
