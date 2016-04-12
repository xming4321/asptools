// XTypeManager.h : Declaration of the CXTypeManager

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"
#include "atlib.h"
#include "XRecords.h"

// CXTypeManager

class ATL_NO_VTABLE CXTypeManager :
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXTypeManager, &CLSID_XTypeManager>,
	public IDispatchImpl<IXTypeManager, &IID_IXTypeManager, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	~CXTypeManager();

	class CXClass : public CXDispatch<IXClass>
	{
		STDMETHOD(AddField)(BSTR key, VARIANT type);
		STDMETHOD(AddRecordset)(BSTR key, VARIANT type);

	public:
		int m_nIndex;
	};

DECLARE_REGISTRY_RESOURCEID(IDR_XTYPEMANAGER)


BEGIN_COM_MAP(CXTypeManager)
	COM_INTERFACE_ENTRY(IXTypeManager)
	COM_INTERFACE_ENTRY(IDispatch)
END_COM_MAP()



	DECLARE_PROTECT_FINAL_CONSTRUCT()

	HRESULT FinalConstruct()
	{
		return S_OK;
	}

	void FinalRelease()
	{
	}

public:
	STDMETHOD(newClass)(BSTR Ver, IXClass** pVal);
	STDMETHOD(newRecordset)(VARIANT varFmt, IXRecords** pVal);

public:
	static CXFields* GetFormat(__int64 hash);
	static BSTR GetFormatName(__int64 hash);
	static __int64 GetFormatHash(BSTR index);
};

OBJECT_ENTRY_AUTO(__uuidof(XTypeManager), CXTypeManager)
