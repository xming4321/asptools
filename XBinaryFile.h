// XBinaryFile.h : Declaration of the CXBinaryFile

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"
#include "XFileStream.h"

// CXBinaryFile

class ATL_NO_VTABLE CXBinaryFile : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXBinaryFile, &CLSID_XBinaryFile>,
	public IDispatchImpl<IXBinaryFile, &IID_IXBinaryFile, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
DECLARE_REGISTRY_RESOURCEID(IDR_XBINARYFILE)


BEGIN_COM_MAP(CXBinaryFile)
	COM_INTERFACE_ENTRY(IXBinaryFile)
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

	STDMETHOD(Create)(BSTR bstrName, VARIANT varOverwrite);
	STDMETHOD(Open)(BSTR bstrName, VARIANT varMode, VARIANT varShareMode);
	STDMETHOD(get_EOF)(VARIANT_BOOL* pVal);
	STDMETHOD(get_lastModify)(DATE* pVal);
	STDMETHOD(get_Position)(DOUBLE* pVal);
	STDMETHOD(put_Position)(DOUBLE newVal);
	STDMETHOD(get_Size)(DOUBLE* pVal);
	STDMETHOD(put_Size)(DOUBLE newVal);
	STDMETHOD(get_Compressed)(VARIANT_BOOL* pVal);
	STDMETHOD(put_Compressed)(VARIANT_BOOL newVal);
	STDMETHOD(Read)(VARIANT varSize, VARIANT* pVal);
	STDMETHOD(Write)(VARIANT varData);
	STDMETHOD(WriteText)(BSTR strData, VARIANT varCodePage);
	STDMETHOD(ReadAllText)(VARIANT varCodePage, BSTR* pVal);
	STDMETHOD(Truncate)(void);
	STDMETHOD(FlushBuffers)(void);
	STDMETHOD(Close)(void);

private:
	CXFileStream m_File;
};

OBJECT_ENTRY_AUTO(__uuidof(XBinaryFile), CXBinaryFile)
