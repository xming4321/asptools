// atLib.h : Declaration of the CXStream

#pragma once

#include <atlbase.h>
#include <atlcoll.h>
#include <atlstr.h>
#include <atlsafe.h>
#include <atlsync.h>
#include <atlcom.h>
#include <winsock2.h>
#include <ws2tcpip.h>

typedef CStringT< wchar_t, StrTraitATL< wchar_t, ChTraitsCRT< wchar_t > > > CXString;
typedef CStringT< char, StrTraitATL< char, ChTraitsCRT< char > > > CXStringA;

template <class T>
class CXKeyStringT
{
public:
	CXKeyStringT() throw()
	{}

	CXKeyStringT(const CXKeyStringT& src) throw()
	{
		m_str = src.m_str;
	}

	CXKeyStringT(const CXString& src) throw()
	{
		m_str = src;
	}

	CXKeyStringT(const CXStringA& src) throw()
	{
		m_str = src;
	}

	CXKeyStringT(LPCWSTR str) throw() :
		m_str(str)
	{
	}

	CXKeyStringT(LPCSTR str) throw() :
		m_str(str)
	{
	}

	BSTR AllocSysString(void) const
	{
		return m_str.AllocSysString();
	}

	friend bool operator==( const CXKeyStringT& str1, const CXKeyStringT& str2 ) throw()
	{
		return(str1.m_str.CompareNoCase(str2.m_str) == 0);
	}

	friend bool operator<( const CXKeyStringT& str1, const CXKeyStringT& str2 ) throw()
	{
		return(str1.m_str.CompareNoCase(str2.m_str) < 0);
	}

	friend bool operator>( const CXKeyStringT& str1, const CXKeyStringT& str2 ) throw()
	{
		return(str1.m_str.CompareNoCase(str2.m_str) > 0);
	}

	T m_str;
};

typedef CXKeyStringT< CXString > CXKeyString;
typedef CXKeyStringT< CXStringA > CXKeyStringA;

DATE inline GetVariantTime(void)
{
	SYSTEMTIME st;
	DATE d;

	GetLocalTime(&st);
	SystemTimeToVariantTime(&st, &d);

	return d;
}

__int64 inline GetLocalFileTime(void)
{
	SYSTEMTIME st;
	FILETIME d;

	GetLocalTime(&st);
	SystemTimeToFileTime(&st, &d);

	return *(__int64*)&d / 10000 - 9435312000000;
}

unsigned int inline rand32()
{
	return rand() | (rand() << 15) | (rand() << 30);
}

extern CLSID CLSID_FreeThreadedMarshaler;

template <class T>
class CXClass : public T
{
public:
	CXClass(void) : m_dwRef(1)
	{}

	virtual ~CXClass(void)
	{}

	static HRESULT SetErrorInfo(LPCOLESTR lpszDesc)
	{
		CComPtr<ICreateErrorInfo> pICEI;

		if (SUCCEEDED(CreateErrorInfo(&pICEI)))
		{
			CComPtr<IErrorInfo> pErrorInfo;
			pICEI->SetGUID(IID_NULL);
			pICEI->SetSource(L"asptools");
			pICEI->SetDescription((LPOLESTR)lpszDesc);
			if (SUCCEEDED(pICEI->QueryInterface(__uuidof(IErrorInfo), (void**)&pErrorInfo)))
				::SetErrorInfo(0, pErrorInfo);
		}

		return DISP_E_EXCEPTION;
	}

	static HRESULT SetErrorInfo(LPCSTR lpszDesc)
	{
		return SetErrorInfo(CXString(lpszDesc));
	}

	static HRESULT GetMessageFromError(long errorno)
	{
		if(errorno <= 0)
			return (HRESULT)errorno;

		CXString strErrorString;

		LPSTR pMsgBuf = NULL;
		if (FormatMessage(FORMAT_MESSAGE_ALLOCATE_BUFFER | FORMAT_MESSAGE_FROM_SYSTEM | FORMAT_MESSAGE_IGNORE_INSERTS,
				NULL, errorno, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), (LPSTR) &pMsgBuf, 0, NULL ) && pMsgBuf)
		{
			strErrorString = pMsgBuf;
			LocalFree(pMsgBuf);
		}

		SetErrorInfo(strErrorString);

		return HRESULT_FROM_WIN32(errorno);
	}

	static HRESULT GetLastError(void)
	{
		return GetMessageFromError(::GetLastError());
	}

	static HRESULT WSAGetMessageFromError(long errorno)
	{
		if(errorno <= 0)
			return (HRESULT)errorno;

		return SetErrorInfo(CXString(gai_strerror(errorno)));
	}

	static HRESULT WSAGetLastError(void)
	{
		return WSAGetMessageFromError(::WSAGetLastError());
	}

public:
	// IUnknown
	STDMETHOD_(ULONG, AddRef)() throw()
	{
		return InterlockedIncrement(&m_dwRef);
	}

	STDMETHOD_(ULONG, Release)() throw()
	{
		long l = InterlockedDecrement(&m_dwRef);
		if(l == 0)
			delete this;

		return l;
	}

	template<class Q>
	HRESULT QueryInterface(Q** pp)
	{
		return QueryInterface(__uuidof(Q), (void **)pp);
	}

	STDMETHOD(QueryInterface)(REFIID iid, void ** ppvObject) throw()
	{
		if(ppvObject == NULL)
			return E_POINTER;

		*ppvObject = NULL;
		if (iid == IID_IUnknown || iid == __uuidof(T))
			*ppvObject = (void*)(T*)this;
		else return E_NOTIMPL;
		AddRef();

		return S_OK;
	}

private:
	long m_dwRef;
};

typedef CXClass<IUnknown> CXObject;

class CXTypeInfoHolder : public CComTypeInfoHolder
{
public:
	CXTypeInfoHolder(const GUID* pguid, const GUID* plibid = &CAtlModule::m_libid, WORD wMajor = 1, WORD wMinor = 0)
	{
		m_pguid = pguid;
		m_plibid = plibid;
		m_wMajor = wMajor;
		m_wMinor = wMinor;
	}

	~CXTypeInfoHolder()
	{
		if (m_pInfo != NULL)
			m_pInfo->Release();
		m_pInfo = NULL;

		for (int j=m_nCount-1; j>=0; j--)
			m_pMap[j].bstr.Empty();

		delete [] m_pMap;
		m_pMap = NULL;
	}
};

template <class T>
class CXDispatch : public CXClass<T>
{
public:
	STDMETHOD(QueryInterface)(REFIID iid, void ** ppvObject) throw()
	{
		if(ppvObject == NULL)
			return E_POINTER;

		*ppvObject = NULL;
		if (iid == IID_IUnknown || iid == IID_IDispatch || iid == __uuidof(T))
			*ppvObject = (void*)(T*)this;
		else return E_NOTIMPL;
		AddRef();

		return S_OK;
	}

public:
// IDispatch
	STDMETHOD(GetTypeInfoCount)(UINT* pctinfo)
	{
		if (pctinfo == NULL) 
			return E_POINTER; 
		*pctinfo = 1;
		return S_OK;
	}
	STDMETHOD(GetTypeInfo)(UINT itinfo, LCID lcid, ITypeInfo** pptinfo)
	{
		return _tih.GetTypeInfo(itinfo, lcid, pptinfo);
	}
	STDMETHOD(GetIDsOfNames)(REFIID riid, LPOLESTR* rgszNames, UINT cNames,
		LCID lcid, DISPID* rgdispid)
	{
		return _tih.GetIDsOfNames(riid, rgszNames, cNames, lcid, rgdispid);
	}
	STDMETHOD(Invoke)(DISPID dispidMember, REFIID riid,
		LCID lcid, WORD wFlags, DISPPARAMS* pdispparams, VARIANT* pvarResult,
		EXCEPINFO* pexcepinfo, UINT* puArgErr)
	{
		return _tih.Invoke((IDispatch*)this, dispidMember, riid, lcid,
		wFlags, pdispparams, pvarResult, pexcepinfo, puArgErr);
	}

protected:
	static CXTypeInfoHolder _tih;
};

template <class T>
class CXDispatchFree : public CXDispatch<T>
{
public:
	CXDispatchFree()
	{
		CoCreateFreeThreadedMarshaler(this, &m_pUnkMarshaler.p);
	}

public:
	STDMETHOD(QueryInterface)(REFIID iid, void ** ppvObject) throw()
	{
		if(ppvObject == NULL)
			return E_POINTER;

		*ppvObject = NULL;
		if (iid == IID_IMarshal)
			return m_pUnkMarshaler.p->QueryInterface(iid, ppvObject);

		return CXDispatch<T>::QueryInterface(iid, ppvObject);
	}

private:
	CComPtr<IUnknown> m_pUnkMarshaler;
};

template <class T>
class CXComPtr
{
public:
	CXComPtr(T* lp) throw()
	{
		if(lp != NULL)
		{
			lp->AddRef();
			p = lp;
		}else p = NULL;
	}

	CXComPtr(void) throw() : p(NULL)
	{}

	CXComPtr(const CXComPtr<T>& lp) throw() : p(NULL)
	{
		operator=(lp);
	}

	~CXComPtr() throw()
	{
		if(p)
			p->Release();
	}

	HRESULT CreateObject() throw()
	{
		Attach(new T);

		return S_OK;
	}

	HRESULT CoCreateObject() throw()
	{
		HRESULT hRes = S_OK;

		CComObject<T>* lp = NULL;
		ATLTRY(lp = new CComObject<T>())
		if (lp != NULL)
		{
			lp->SetVoid(NULL);
			lp->InternalFinalConstructAddRef();
			hRes = lp->FinalConstruct();
			if (SUCCEEDED(hRes))
				hRes = lp->_AtlFinalConstruct();
			lp->InternalFinalConstructRelease();
			if (hRes != S_OK)
			{
				delete lp;
				lp = NULL;
			}
		}

		if(lp)
		{
			lp->AddRef();
			Attach(lp);
		}

		return hRes;
	}

	HRESULT CoCreateObject(const IID & id) throw()
	{
		return CoCreateInstance(id, NULL, CLSCTX_INPROC_SERVER, IID_IActiveScript, (void **)&p); 
	}

	template <class Q>
	HRESULT QueryInterface(Q** pp) const throw()
	{
		return QueryInterface(__uuidof(Q), (void**)pp);
	}

	HRESULT QueryInterface(REFIID iid, void ** ppvObject) const throw()
	{
		if(p == NULL)return E_NOTIMPL;
		return p->QueryInterface(iid, (void**)ppvObject);
	}

	T* operator=(T* lp) throw()
	{
		if(lp != NULL)
			lp->AddRef();

		return Attach(lp);
	}

	T* operator=(const CXComPtr<T>& lp) throw()
	{
		return operator=(lp.p);
	}

	template <class Q>
	T* operator=(Q* lp) throw()
	{
		T* q = NULL;

		if(lp != NULL)
			lp->QueryInterface(__uuidof(T), (void**)&q);

		return Attach(q);
	}

	template <class Q>
	T* operator=(const CXComPtr<Q>& lp) throw()
	{
		return operator=(lp.p);
	}

	operator T*() const throw()
	{
		return p;
	}

	T& operator*() const throw()
	{
		return *p;
	}

	T** operator&() throw()
	{
		return &p;
	}

	T* operator->()
	{
		return p;
	}

	void Release() throw()
	{
		T* pTemp = Detach();
		if(pTemp)pTemp->Release();
	}

	T* Attach(T* p2) throw()
	{
		T* p1 = (T*)InterlockedExchangePointer((void**)&p, p2);
		if(p1)p1->Release();

		return p2;
	}

	T* Detach() throw()
	{
		return (T*)InterlockedExchangePointer((void**)&p, NULL);
	}

	T* p;
};

template< typename T = BYTE, size_t size = 0 >
class CXAutoPtr
{
public:
	CXAutoPtr() throw() :
		m_p( NULL )
	{
		if( size )Allocate(size);
	}

	~CXAutoPtr() throw()
	{
		Free();
	}

	operator T*() const throw()
	{
		return( m_p );
	}

	CXAutoPtr< T >& operator=( CXAutoPtr< T >& p ) throw()
	{
		Free();
		Attach( p.Detach() );  // Transfer ownership

		return( *this );
	}

	// Allocate the vector
	bool Allocate( size_t nElements ) throw()
	{
		Free();
		ATLTRY( m_p = new T[nElements] );
		if( m_p == NULL )
		{
			return( false );
		}

		return( true );
	}

	void Copy(const void* p, size_t nElements) throw()
	{
		Allocate(nElements);
		CopyMemory(m_p, p, nElements);
	}

	// Attach to an existing pointer (takes ownership)
	void Attach( T* p ) throw()
	{
		Free();
		m_p = p;
	}
	// Detach the pointer (releases ownership)
	T* Detach() throw()
	{
		T* p;

		p = m_p;
		m_p = NULL;

		return( p );
	}
	// Delete the vector pointed to, and set the pointer to NULL
	void Free() throw()
	{
		if(m_p)
		{
			delete[] m_p;
			m_p = NULL;
		}
	}

public:
	T* m_p;
};

template <class T>
class CXEnumVARIANT : public CXClass<IEnumVARIANT>
{
public:
	CXEnumVARIANT(T* pCollection, ULONG ulPos = 0) : m_nPos(ulPos)
	{
		pCollection->FillEnum(m_arrayVariant);
	}

	~CXEnumVARIANT()
	{
		int i;

		for (i = 0; i < (int)m_arrayVariant.GetCount(); i ++)
			VariantClear(&m_arrayVariant[i]);
	}

public:
	// IEnumVARIANT
	STDMETHOD(Next)(ULONG celt, VARIANT *rgVar, ULONG *pCeltFetched)
	{
		if(rgVar == NULL)
			return E_INVALIDARG;

		if(pCeltFetched != NULL)
			*pCeltFetched = 0;
		else if(celt != 1)
			return E_INVALIDARG;

		ZeroMemory(rgVar, sizeof(tagVARIANT));

		if(m_nPos < (int)m_arrayVariant.GetCount())
		{
			VariantCopy(rgVar, &m_arrayVariant[m_nPos]);
			m_nPos ++;

			if(pCeltFetched != NULL)*pCeltFetched = 1;

			return S_OK;
		}

		return S_FALSE;
	}

	STDMETHOD(Skip)(ULONG celt)
	{
		m_nPos += celt;
		return S_OK;
	}

	STDMETHOD(Reset)(void)
	{
		m_nPos = 0;
		return S_OK;
	}

	STDMETHOD(Clone)(IEnumVARIANT **ppEnum)
	{
		return E_NOTIMPL;
	}

private:
	ULONG m_nPos;
	CAtlArray<VARIANT> m_arrayVariant;
};

template <class T>
HRESULT getNewEnum(T* pThis, IUnknown **ppEnumReturn)
{
	CXComPtr< CXEnumVARIANT<T> > pEnum;

	pEnum.Attach(new CXEnumVARIANT<T>(pThis));
	return pEnum.QueryInterface(ppEnumReturn);
}

class CXVarPtr
{
public:
	CXVarPtr(void) :
		m_pData(NULL),
		m_nSize(0)
	{
	}

	~CXVarPtr(void)
	{
	}

	HRESULT Attach(VARIANT* pvar)
	{
		while(pvar && pvar->vt == (VT_VARIANT | VT_BYREF))
			pvar = pvar->pvarVal;

		if(!pvar)return DISP_E_TYPEMISMATCH;

		if(pvar->vt == VT_UNKNOWN || pvar->vt == VT_DISPATCH)
		{
			HRESULT hr;
			CComDispatchDriver pDisp;

			if(pvar->punkVal == NULL)
				return DISP_E_UNKNOWNINTERFACE;

			hr = pvar->punkVal->QueryInterface(&pDisp);
			if(FAILED(hr))return hr;

			hr = pDisp.GetProperty(0, &m_varTemp);
			if(FAILED(hr))return hr;

			pvar = &m_varTemp;
		}

		if(pvar->vt == VT_BSTR)
		{
			if(m_pData = (BYTE *)pvar->pcVal)
				m_nSize = SysStringByteLen(pvar->bstrVal);
		}else if(pvar->vt & VT_ARRAY)
		{
			if(pvar->parray == NULL)
				return TYPE_E_TYPEMISMATCH;

			LPSAFEARRAY pArray = pvar->vt & VT_BYREF ? *pvar->pparray : pvar->parray;
			ULONG i;

			m_nSize = 0;
			switch (pvar->vt & VT_TYPEMASK)
			{
			case VT_UI1:
			case VT_I1:
				m_nSize = sizeof(BYTE);
				break;
			case VT_I2:
			case VT_UI2:
			case VT_BOOL:
				m_nSize = sizeof(short);
				break;
			case VT_I4:
			case VT_UI4:
			case VT_R4:
			case VT_INT:
			case VT_UINT:
			case VT_ERROR:
				m_nSize = sizeof(long);
				break;
			case VT_I8:
			case VT_UI8:
				m_nSize = sizeof(LONGLONG);
				break;
			case VT_R8:
			case VT_CY:
			case VT_DATE:
				m_nSize = sizeof(double);
				break;
			default:
				return TYPE_E_TYPEMISMATCH;
			}

			if(pArray->cDims > 0)
				for(i = 0; i < pArray->cDims; i ++)
					m_nSize *= pArray->rgsabound[i].cElements;
			else m_nSize = 0;

			m_pData = (BYTE *)pArray->pvData;
		}else if(pvar->vt != VT_EMPTY)
			return TYPE_E_TYPEMISMATCH;

		return S_OK;
	}

	HRESULT Attach(VARIANT& var)
	{
		return Attach(&var);
	}

	HRESULT Create(UINT nSize)
	{
		m_Array.Destroy();

		m_Array.Create(nSize);
		LPSAFEARRAY pArray = m_Array;

		if(pArray == NULL)
			return E_OUTOFMEMORY;

		m_pData = (BYTE *)pArray->pvData;
		m_nSize = nSize;

		return S_OK;
	}

	HRESULT GetVariant(VARIANT& var, int nSize = -1)
	{
		if(nSize != -1 && nSize != m_nSize)
			m_Array.Resize(nSize);
		var.parray = m_Array.Detach();
		var.vt = VT_ARRAY | VT_UI1;

		return S_OK;
	}

	HRESULT GetVariant(VARIANT* var, int nSize = -1)
	{
		return GetVariant(*var, nSize);
	}

	BYTE * m_pData;
	UINT m_nSize;

private:
	CComVariant m_varTemp;
	CComSafeArray<BYTE> m_Array;
};

template <class T>
HRESULT inline varGetString(VARIANT* pvar, T& str, LPCWSTR strDefault = NULL)
{
	while(pvar && pvar->vt == (VT_VARIANT | VT_BYREF))
		pvar = pvar->pvarVal;

	if(pvar->vt == VT_EMPTY)
	{
		str.Empty();
		return S_OK;
	}

	if(pvar->vt == VT_ERROR)
	{
		if(strDefault)str = strDefault;
		return S_FALSE;
	}

	if(pvar->vt == VT_BSTR)
	{
		if(pvar->bstrVal)str = pvar->bstrVal;
		return S_OK;
	}

	HRESULT hr;
	CComVariant varTemp;

	hr = VariantChangeType(&varTemp, pvar, VARIANT_ALPHABOOL, VT_BSTR);
	if(FAILED(hr))return hr;

	str = varTemp.bstrVal;

	return S_OK;
}

template <class T>
HRESULT inline varGetString(VARIANT& var, T& str, LPCWSTR strDefault = NULL)
{
	return varGetString(&var, str, strDefault);
}

long inline varGetNumbar(VARIANT* pvar, long nDefault = 0)
{
	while(pvar && pvar->vt == (VT_VARIANT | VT_BYREF))
		pvar = pvar->pvarVal;

	if(pvar->vt == VT_ERROR)
		return nDefault;

	switch (pvar->vt)
	{
	case VT_UI1:
	case VT_I1:
		return pvar->ulVal & 0xff;
	case VT_I2:
	case VT_UI2:
	case VT_BOOL:
		return pvar->ulVal & 0xffff;
	case VT_I4:
	case VT_UI4:
	case VT_INT:
	case VT_UINT:
		return pvar->ulVal;
	}

	CComVariant var;

	if(FAILED(var.ChangeType(VT_I4, pvar)))
		return nDefault;

	return var.lVal;
}

long inline varGetNumbar(VARIANT& v, long nDefault = 0)
{
	return varGetNumbar(&v, nDefault);
}

double inline varGetFloat(VARIANT* pvar, double nDefault = 0)
{
	while(pvar && pvar->vt == (VT_VARIANT | VT_BYREF))
		pvar = pvar->pvarVal;

	if(pvar->vt == VT_ERROR)
		return nDefault;

	switch (pvar->vt)
	{
	case VT_UI1:
	case VT_I1:
		return pvar->ulVal & 0xff;
	case VT_I2:
	case VT_UI2:
	case VT_BOOL:
		return pvar->ulVal & 0xffff;
	case VT_I4:
	case VT_UI4:
	case VT_INT:
	case VT_UINT:
		return pvar->ulVal;
	case VT_R4:
		return pvar->fltVal;
	case VT_R8:
		return pvar->dblVal;
	}

	CComVariant var;

	if(FAILED(var.ChangeType(VT_R8, pvar)))
		return nDefault;

	return var.dblVal;
}

double inline varGetFloat(VARIANT& v, double nDefault = 0)
{
	return varGetFloat(&v, nDefault);
}

class CXVariant : public tagVARIANT
{
public:
	CXVariant() throw()
	{
		ZeroMemory(this, sizeof(tagVARIANT));
	}

	CXVariant(const CXVariant& src) throw()
	{
		vt = VT_EMPTY;
		HRESULT hr = ::VariantCopy(this, (VARIANT*)&src);
		if (FAILED(hr))vt = VT_ERROR;
	}

	~CXVariant() throw()
	{
		::VariantClear(this);
	}

	HRESULT ChangeType(VARTYPE vtNew, const VARIANT* pSrc = NULL)
	{
		VARIANT* pVar = const_cast<VARIANT*>(pSrc);
		// Convert in place if pSrc is NULL
		if (pVar == NULL)
			pVar = this;
		// Do nothing if doing in place convert and vts not different
		return ::VariantChangeType(this, pVar, 0, vtNew);
	}

private:
	int GetClass(void) const throw()
	{
		switch(vt)
		{
		case VT_UNKNOWN:
		case VT_DISPATCH:
			return VT_UNKNOWN;
		case VT_BSTR:
			return VT_BSTR;
		case VT_DATE:
			return VT_DATE;
		case VT_BOOL:
			return VT_BOOL;
		case VT_R4:
		case VT_R8:
			return VT_R8;
		case VT_I1:
		case VT_UI1:
		case VT_I2:
		case VT_UI2:
		case VT_I4:
		case VT_UI4:
		case VT_INT:
		case VT_UINT:
		case VT_I8:
		case VT_UI8:
			return VT_UI8;
		default:
			return 0;
		}
	}

	template <typename T>
	static int _check_value(const T v)
	{
		if(v < 0)return -1;
		if(v > 0)return 1;
		return 0;
	}

	double _get_float(void) const throw()
	{
		if(vt == VT_R4)
			return fltVal;
		if(vt == VT_R8)
			return dblVal;

		return 0;
	}

	__int64 _get_int(void) const throw()
	{
		switch(vt)
		{
		case VT_I1:
		case VT_UI1:
			return bVal;
		case VT_I2:
		case VT_UI2:
			return iVal;
		case VT_I4:
		case VT_UI4:
		case VT_INT:
		case VT_UINT:
			return lVal;
		case VT_I8:
		case VT_UI8:
			return llVal;
		}

		return 0;
	}

	static IUnknown* _getobjaddr(IUnknown* p)
	{
		IUnknown* p1 = NULL;

		if(p)
		{
			p->QueryInterface(&p1);
			p1->Release();
		}
		return p1;
	}

public:
	int Compare(const CXVariant* varSrc) const throw()
	{
		int vc, vc1;

		vc = GetClass();
		vc1 = varSrc->GetClass();

		if(vc != vc1 || vc == 0)
			return vc - vc1;

		switch(vc)
		{
		case VT_UNKNOWN:return _check_value(_getobjaddr(punkVal) - _getobjaddr(varSrc->punkVal));
		case VT_BSTR:	if(!bstrVal && !varSrc->bstrVal)return 0;
						if(bstrVal && varSrc->bstrVal)return _wcsicmp(bstrVal, varSrc->bstrVal);
						if(bstrVal)return 1;
						return -1;
		case VT_DATE:	return _check_value(date - varSrc->date);
		case VT_BOOL:	return _check_value(boolVal - varSrc->boolVal);
		case VT_R8:		return _check_value(_get_float() - varSrc->_get_float());
		case VT_UI8:	return _check_value(_get_int() - varSrc->_get_int());
		}

		return 0;
	}

	int Compare(const CXVariant& varSrc) const throw()
	{
		return Compare(&varSrc);
	}

	int Compare(const VARIANT* varSrc) const throw()
	{
		return Compare((const CXVariant*)varSrc);
	}

	int Compare(const VARIANT& varSrc) const throw()
	{
		return Compare((const CXVariant*)&varSrc);
	}

	bool operator==(const VARIANT& varSrc) const throw()
	{
		return Compare(&varSrc) == 0;
	}

	bool operator<(const VARIANT& varSrc) const throw()
	{
		return Compare(&varSrc) < 0;
	}

	bool operator>(const VARIANT& varSrc) const throw()
	{
		return Compare(&varSrc) > 0;
	}

	CXVariant& operator=(const CXVariant& varSrc)
	{
		VariantCopy(this, (VARIANT*)&varSrc);
		return *this;
	}

	CXVariant& operator=(const VARIANT& varSrc)
	{
		VariantCopy(this, (VARIANT*)&varSrc);
		return *this;
	}
};

class CXAutoLock
{
// Constructors
public:
	CXAutoLock(CCriticalSection& Object, BOOL bInitialLock = FALSE)
	{
		m_pObject = &Object;
		m_nAcquired = 0;

		if (bInitialLock)
			Enter();
	}

	CXAutoLock(CCriticalSection* pObject, BOOL bInitialLock = FALSE)
	{
		m_pObject = pObject;
		m_nAcquired = 0;

		if (bInitialLock)
			Enter();
	}

// Operations
public:
	void Enter()
	{
		m_nAcquired ++;
		m_pObject->Enter();
	}

	void Leave()
	{
		m_pObject->Leave();
		m_nAcquired --;
	}

	BOOL IsLocked()
	{
		return m_nAcquired > 0;
	}

// Implementation
public:
	~CXAutoLock()
	{
		while(m_nAcquired)
			Leave();
	}

protected:
	CCriticalSection* m_pObject;
	int    m_nAcquired;
};

class CXStream
{
public:
	CXStream(IStream* pStream)
	{
		m_pStream = pStream;
	}

public:
	enum SeekPosition { begin = 0x0, current = 0x1, end = 0x2 };

	ULONGLONG GetLength(void)
	{
		STATSTG ss;

		HRESULT hr = Stat(&ss, 0);
		if(FAILED(hr))return 0;

		return ss.cbSize.QuadPart;
	}

	void inline SetSize(ULONGLONG size)
	{
		ULARGE_INTEGER lVal;

		lVal.QuadPart = size;
		SetSize(lVal);
	}

	ULONGLONG inline Seek(ULONGLONG pos, UINT nFrom = begin)
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


	HRESULT inline ReadStream(void* pv, long cb)
	{
		char *pBuffer = (char*)pv;
		HRESULT hr;
		ULONG cbRead;

		while(cb)
		{
			hr = Read(pBuffer, cb, &cbRead);
			if(FAILED(hr))return hr;

			cb -= cbRead;
			pBuffer += cbRead;
		}

		return S_OK;
	}

	template<class TYPE>
	HRESULT inline ReadStream(TYPE& v)
	{
		return ReadStream(&v, sizeof(TYPE));
	}

	HRESULT inline ReadStream(CXString& str)
	{
		HRESULT hr;
		int nSize;

		hr = ReadStream(nSize);
		if(FAILED(hr))return hr;

		if(nSize < 0 || nSize > 10 * 1024 * 1024)
			return DISP_E_OVERFLOW;

		hr = ReadStream(str.GetBuffer(nSize), nSize * 2);
		str.ReleaseBuffer(nSize);

		return hr;
	}

	HRESULT inline ReadStream(VARIANT& var)
	{
		HRESULT hr;

		if(var.vt == VT_BSTR)
		{
			int nSize;

			hr = ReadStream(nSize);
			if(FAILED(hr))return hr;

			if(nSize < 0 || nSize > 10 * 1024 * 1024)
				return DISP_E_OVERFLOW;

			if(var.bstrVal)::SysFreeString(var.bstrVal);
			var.bstrVal = ::SysAllocStringLen(NULL, nSize);
			return ReadStream(var.bstrVal, nSize * 2);
		}

		switch(var.vt)
		{
		case VT_BOOL:
		case VT_I4:
		case VT_R4: hr = ReadStream(var.lVal);break;
		case VT_DATE:
		case VT_R8: hr = ReadStream(var.dblVal);break;
		case VT_I2: hr = ReadStream(var.iVal);break;
		default: return E_INVALIDARG;
		};

		return hr;
	}

	HRESULT inline ReadStream(CComVariant& var)
	{
		return ReadStream(*(VARIANT*)&var);
	}

	template<class TYPE>
	HRESULT inline WriteStream(const TYPE& v)
	{
		return Write(&v, sizeof(TYPE));
	}

	HRESULT inline WriteStream(LPCWSTR str)
	{
		HRESULT hr;
		int nSize = (int)wcslen(str);

		hr = WriteStream(nSize);
		if(FAILED(hr))return hr;

		return Write((LPCWSTR)str, nSize * 2);
	}

	HRESULT inline WriteStream(const CXString& str)
	{
		HRESULT hr;
		int nSize = str.GetLength();

		hr = WriteStream(nSize);
		if(FAILED(hr))return hr;

		return Write((LPCWSTR)str, nSize * 2);
	}

	HRESULT inline WriteStream(const VARIANT& var)
	{
		HRESULT hr;

		if(var.vt == VT_BSTR)
		{
			int nSize = 0;

			if(var.bstrVal)nSize = ::SysStringLen(var.bstrVal);
			hr = WriteStream(nSize);
			if(FAILED(hr))return hr;

			if(!nSize)return S_OK;

			return Write(var.bstrVal, nSize * 2);
		}

		switch(var.vt)
		{
		case VT_BOOL:
		case VT_I4:
		case VT_R4: hr = WriteStream(var.lVal);break;
		case VT_DATE:
		case VT_R8: hr = WriteStream(var.dblVal);break;
		case VT_I2: hr = WriteStream(var.iVal);break;
		default: return E_INVALIDARG;
		};

		return hr;
	}

	HRESULT inline WriteStream(const CComVariant& var)
	{
		return WriteStream(*(VARIANT*)&var);
	}

	BOOL inline isEOF()
	{
		return Seek(0, current) >= GetLength();
	}

public:
	// ISequentialStream
	HRESULT Read(void *pv, ULONG cb, ULONG *pcbRead = NULL)
	{
		if(!m_pStream)return E_NOTIMPL;
		return m_pStream->Read(pv, cb, pcbRead);
	}

	HRESULT Write(const void *pv, ULONG cb, ULONG *pcbWritten = NULL)
	{
		if(!m_pStream)return E_NOTIMPL;
		return m_pStream->Write(pv, cb, pcbWritten);
	}

public:
	// IStream
	HRESULT Seek(LARGE_INTEGER dlibMove, DWORD dwOrigin, ULARGE_INTEGER *plibNewPosition)
	{
		if(!m_pStream)return E_NOTIMPL;
		return m_pStream->Seek(dlibMove, dwOrigin, plibNewPosition);
	}

	HRESULT SetSize(ULARGE_INTEGER libNewSize)
	{
		if(!m_pStream)return E_NOTIMPL;
		return m_pStream->SetSize(libNewSize);
	}

	HRESULT CopyTo(IStream *pstm, ULARGE_INTEGER cb, ULARGE_INTEGER *pcbRead, ULARGE_INTEGER *pcbWritten)
	{
		if(!m_pStream)return E_NOTIMPL;
		return m_pStream->CopyTo(pstm, cb, pcbRead, pcbWritten);
	}

	HRESULT Commit(DWORD grfCommitFlags)
	{
		if(!m_pStream)return E_NOTIMPL;
		return m_pStream->Commit(grfCommitFlags);
	}

	HRESULT Revert(void)
	{
		if(!m_pStream)return E_NOTIMPL;
		return m_pStream->Revert();
	}

	HRESULT LockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType)
	{
		if(!m_pStream)return E_NOTIMPL;
		return m_pStream->LockRegion(libOffset, cb, dwLockType);
	}

	HRESULT UnlockRegion(ULARGE_INTEGER libOffset, ULARGE_INTEGER cb, DWORD dwLockType)
	{
		if(!m_pStream)return E_NOTIMPL;
		return m_pStream->UnlockRegion(libOffset, cb, dwLockType);
	}

	HRESULT Stat(STATSTG *pstatstg, DWORD grfStatFlag)
	{
		if(!m_pStream)return E_NOTIMPL;
		return m_pStream->Stat(pstatstg, grfStatFlag);
	}

	HRESULT Clone(IStream **ppstm)
	{
		if(!m_pStream)return E_NOTIMPL;
		return m_pStream->Clone(ppstm);
	}

private:
	CXComPtr<IStream> m_pStream;
};

BOOL inline isVBType(VARTYPE vt)
{
	return vt == VT_BSTR || vt == VT_BOOL || vt == VT_DATE || vt == VT_R4 || vt == VT_R8 || vt == VT_I2 || vt == VT_I4;
}

BOOL inline isNumber(int n)
{
	return n >= '0' && n <= '9';
}