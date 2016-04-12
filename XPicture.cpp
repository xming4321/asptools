// XPicture.cpp : Implementation of CXPicture

#include "stdafx.h"
#include "XPicture.h"
#include "XMemStream.h"
#include "XFileStream.h"


// CXPicture

class CxStreamFile : public CxFile
	{
public:
	CxStreamFile(IStream* pStream)
	{
		m_pStream = pStream;
	}

	virtual bool Close()
	{
		m_pStream.Release();
		return TRUE;
	}
	virtual size_t	Read(void *buffer, size_t size, size_t count)
	{
		if (!m_pStream) return 0;

		if (size==0) return 0;
		ULONG n;
		
		HRESULT hr = m_pStream->Read(buffer, (ULONG)size*count, &n);
		if FAILED(hr) return 0;

		return n/size;
	}
//////////////////////////////////////////////////////////
	virtual size_t	Write(const void *buffer, size_t size, size_t count)
	{
		if (!m_pStream) return 0;

		if (size==0) return 0;
		ULONG n;
		
		HRESULT hr = m_pStream->Write(buffer, (ULONG)size*count, &n);
		if FAILED(hr) return 0;

		return n/size;
	}
//////////////////////////////////////////////////////////
	virtual bool Seek(long offset, int origin)
	{
		if (!m_pStream) return false;

		LARGE_INTEGER u;
		ULARGE_INTEGER n;

		n.QuadPart = 0;
		u.QuadPart = offset;
		HRESULT hr = m_pStream->Seek(u, origin, &n);
		if FAILED(hr) return false;
		return true;
	}
//////////////////////////////////////////////////////////
	virtual long Tell()
	{
		if (!m_pStream) return 0;

		LARGE_INTEGER u;
		ULARGE_INTEGER n;

		n.QuadPart = 0;
		u.QuadPart = 0;
		HRESULT hr = m_pStream->Seek(u, 1, &n);
		if FAILED(hr) return 0;
		return (long)n.QuadPart;
	}
//////////////////////////////////////////////////////////
	virtual long	Size()
	{
		if (!m_pStream) return 0;

		STATSTG stat;

		HRESULT hr = m_pStream->Stat(&stat, 1);
		if FAILED(hr) return 0;

		return (long)stat.cbSize.QuadPart;
	}
//////////////////////////////////////////////////////////
	virtual bool	Flush()
	{
		return false;
	}
//////////////////////////////////////////////////////////
	virtual bool	Eof()
	{
		return (Tell()>=Size());
	}
//////////////////////////////////////////////////////////
	virtual long	Error()
	{
		if (!m_pStream) return -1;
		return 0;
	}
//////////////////////////////////////////////////////////
	virtual bool PutC(unsigned char c)
	{
		if (!m_pStream) return false;

		ULONG n;
		
		HRESULT hr = m_pStream->Write((const void *)&c, (ULONG)1, &n);
		if FAILED(hr) return 0;

		return (n==1);
	}
//////////////////////////////////////////////////////////
	virtual long	GetC()
	{
		if (!m_pStream) return 0;
		if (Eof()) return EOF;

		ULONG n;
		unsigned char c;

		HRESULT hr = m_pStream->Read(&c, (ULONG)1, &n);
		if FAILED(hr) return 0;

		return (long)c;
	}
protected:
	CXComPtr<IStream> m_pStream;
	};

CXPicture::CImagePaint::CImagePaint(CXPicture* pImage) : m_pImage(pImage)
{
	LOGBRUSH lb = {(m_pImage->m_nBrushStyle > 0) ? BS_HATCHED : BS_SOLID, m_pImage->m_colocBrush, m_pImage->m_nBrushStyle - 1};

	if(m_pImage->m_nPenStyle == -1)m_hPen = (HPEN)GetStockObject(NULL_PEN);
	else m_hPen = CreatePen(m_pImage->m_nPenStyle, m_pImage->m_nPenWidth, m_pImage->m_colorPen);

	if(m_pImage->m_nBrushStyle < 0)m_hBrush = (HBRUSH)GetStockObject(NULL_BRUSH);
	else m_hBrush = CreateBrushIndirect(&lb);

	m_pImage->m_cs.Enter();
	m_hOldPen = (HPEN)SelectObject(m_pImage->m_Image.hMemDC, m_hPen);
	m_hOldBrush = (HBRUSH)SelectObject(m_pImage->m_Image.hMemDC, m_hBrush);
}

CXPicture::CImagePaint::~CImagePaint()
{
	if(m_hOldPen)SelectObject(m_pImage->m_Image.hMemDC, m_hOldPen);
	if(m_hOldBrush)SelectObject(m_pImage->m_Image.hMemDC, m_hOldBrush);
	m_pImage->m_cs.Leave();

	DeleteObject(m_hPen);
	DeleteObject(m_hBrush);
}

CXPicture::CImagePoly::CImagePoly(VARIANT var) :
	m_pPoint(NULL), m_nCount(0), m_hr(S_OK)
{
	VARIANT* pvar = &var;
	CComVariant varTemp;

	if(pvar->vt == VT_UNKNOWN || pvar->vt == VT_DISPATCH)
	{
		CComDispatchDriver pDisp;

		if(pvar->punkVal == NULL)
		{
			m_hr = DISP_E_UNKNOWNINTERFACE;
			return;
		}

		m_hr = pvar->punkVal->QueryInterface(&pDisp);
		if(FAILED(m_hr))return;

		m_hr = pDisp.GetProperty(0, &varTemp);
		if(FAILED(m_hr))return;

		pvar = &varTemp;
	}

	if(pvar->vt != (VT_ARRAY | VT_VARIANT))
	{
		m_hr = DISP_E_TYPEMISMATCH;
		return;
	}

	LPSAFEARRAY pArray = pvar->vt & VT_BYREF ? *pvar->pparray : pvar->parray;

	if((pArray->cDims != 1) || (pArray->rgsabound[0].cElements & 1))
	{
		m_hr = DISP_E_TYPEMISMATCH;
		return;
	}

	if(m_nCount = pArray->rgsabound[0].cElements / 2)
	{
		m_pPoint = new POINT[m_nCount];

		pvar = (VARIANT *)pArray->pvData;

		for(int i = 0; i < m_nCount; i ++)
		{
			m_pPoint[i].x = varGetNumbar(pvar[i * 2]);
			m_pPoint[i].y = varGetNumbar(pvar[i * 2 + 1]);
		}
	}
}

CXPicture::CImagePoly::~CImagePoly()
{
	if(m_pPoint)delete m_pPoint;
}

CXPicture::CXPicture()
{
	m_pUnkMarshaler = NULL;

	memset(&m_lfFont, 0, sizeof(m_lfFont));

	m_lfFont.lfHeight        = 10; 
	m_lfFont.lfCharSet       = DEFAULT_CHARSET;
	m_lfFont.lfWeight        = FW_NORMAL;
	m_lfFont.lfWidth         = 0; 
	m_lfFont.lfEscapement    = 0; 
	m_lfFont.lfOrientation   = 0; 
	m_lfFont.lfItalic        = FALSE; 
	m_lfFont.lfUnderline     = FALSE; 
	m_lfFont.lfStrikeOut     = FALSE; 
	m_lfFont.lfOutPrecision  = OUT_DEFAULT_PRECIS; 
	m_lfFont.lfClipPrecision = CLIP_DEFAULT_PRECIS; 
	m_lfFont.lfQuality       = PROOF_QUALITY; 
	m_lfFont.lfPitchAndFamily= DEFAULT_PITCH | FF_DONTCARE ; 
    strcpy_s(m_lfFont.lfFaceName, sizeof(m_lfFont.lfFaceName), _T("Arial"));

	m_colocBrush = RGB(255, 255, 255);
	m_nBrushStyle = 0;
	m_colorPen = RGB(255, 255, 255);
	m_nPenWidth = 1;
	m_nPenStyle = PS_SOLID;
}

STDMETHODIMP CXPicture::Load(BSTR strFile, VARIANT varType)
{
	CXComPtr<CXFileStream> pStream;
	CXComPtr<IPicture> pPicture;
	HRESULT hr;
	int nType = varGetNumbar(varType);

	pStream.CreateObject();
	hr = pStream->Open(strFile);
	if(FAILED(hr))return hr;

	CxStreamFile sf(pStream);

	m_cs.Enter();

	if(!nType)
	{
		DWORD pos = sf.Tell();
		if(!m_Image.Decode(&sf, CxImage::getFormatByName(CXStringA(strFile))))
		{
			sf.Seek(pos,SEEK_SET);
			if(!m_Image.Decode(&sf, 0))
				hr = CXObject::SetErrorInfo(m_Image.GetLastError());
		}
	}else if(!m_Image.Decode(&sf, nType))
		hr = CXObject::SetErrorInfo(m_Image.GetLastError());
	m_cs.Leave();

	return hr;
}

STDMETHODIMP CXPicture::Save(BSTR strFile, VARIANT varType)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	HRESULT hr;
	CXComPtr<CXFileStream> pStream;
	int nType = varGetNumbar(varType);

	if(!nType)nType = CxImage::getFormatByName(CXStringA(strFile));

	pStream.CreateObject();
	hr = pStream->Create(strFile);
	if(FAILED(hr))return hr;

	CxStreamFile sf(pStream);

	m_cs.Enter();
	if(nType == 0)
		nType = (short)m_Image.GetType();
	if(!m_Image.Encode(&sf, nType))
		hr = CXObject::SetErrorInfo(m_Image.GetLastError());
	m_cs.Leave();

	return hr;
}

STDMETHODIMP CXPicture::LoadFromData(VARIANT varData, VARIANT varType)
{
	CXBlockStream blockStream;
	CXComPtr<IPicture> pPicture;
	HRESULT hr;
	int nType = varGetNumbar(varType);

	hr = blockStream.AttachVariant(varData);
	if(FAILED(hr))return hr;

	CxStreamFile sf(&blockStream);

	m_cs.Enter();
	if(!m_Image.Decode(&sf, nType))
		hr = CXObject::SetErrorInfo(m_Image.GetLastError());
	m_cs.Leave();

	return hr;
}

STDMETHODIMP CXPicture::GetImageData(VARIANT varType, VARIANT* pVal)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	HRESULT hr;
	CXMemStream memStream;
	int nType = varGetNumbar(varType);

	CxStreamFile sf(&memStream);

	m_cs.Enter();
	if(nType == 0)
		nType = (short)m_Image.GetType();
	if(!m_Image.Encode(&sf, nType))
		hr = CXObject::SetErrorInfo(m_Image.GetLastError());
	m_cs.Leave();

	return memStream.GetVariant(pVal);
}

STDMETHODIMP CXPicture::LoadRawData(VARIANT varData)
{
	CXVarPtr varPtr;
	HRESULT hr;
	UINT n;

	hr = varPtr.Attach(varData);
	if(FAILED(hr))return hr;

	n = varPtr.m_nSize;

	m_cs.Enter();

	n = m_Image.GetEffWidth() * m_Image.GetHeight();
	if(n > varPtr.m_nSize)n = varPtr.m_nSize;
	memcpy(m_Image.GetBits(), varPtr.m_pData, n);

	m_cs.Leave();

	return hr;
}

STDMETHODIMP CXPicture::GetRawData(VARIANT* pVal)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	CXVarPtr varPtr;

	m_cs.Enter();

	varPtr.Create(m_Image.GetEffWidth() * m_Image.GetHeight());
	memcpy(varPtr.m_pData, m_Image.GetBits(), varPtr.m_nSize);

	m_cs.Leave();

	return varPtr.GetVariant(pVal);
}

STDMETHODIMP CXPicture::LoadInfo(BSTR strFile, VARIANT varType)
{
	m_Image.info.bLoadImage = FALSE;
	return Load(strFile, varType);
}

STDMETHODIMP CXPicture::LoadInfoFromData(VARIANT varData, VARIANT varType)
{
	m_Image.info.bLoadImage = FALSE;
	return LoadInfoFromData(varData, varType);
}

STDMETHODIMP CXPicture::Create(long nWidth, long nHeight, short Bpp)
{
	HRESULT hr = S_OK;

	m_cs.Enter();
	if(!m_Image.Create(nWidth, nHeight, Bpp))
		hr = CXObject::SetErrorInfo(m_Image.GetLastError());
	m_cs.Leave();

	return hr;
}

STDMETHODIMP CXPicture::Clear(void)
{
	m_cs.Enter();
	m_Image.Clear();
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXPicture::get_Type(short *pvar)
{
	m_cs.Enter();
	*pvar = (short)m_Image.GetType();
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXPicture::get_Height(long *pvar)
{
	m_cs.Enter();
	*pvar = m_Image.GetHeight();
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXPicture::get_Width(long *pvar)
{
	m_cs.Enter();
	*pvar = m_Image.GetWidth();
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXPicture::get_Bpp(short *pvar)
{
	*pvar = m_Image.GetBpp();
	return S_OK;
}

STDMETHODIMP CXPicture::put_Bpp(short var)
{
	HRESULT hr = S_OK;

	m_cs.Enter();
	if (var > m_Image.GetBpp())
	{
		if (!m_Image.IncreaseBpp(var))
			hr = CXObject::SetErrorInfo(m_Image.GetLastError());
	}
	else if(var < m_Image.GetBpp())
	{
		if (!m_Image.DecreaseBpp(var, true))
			hr = CXObject::SetErrorInfo(m_Image.GetLastError());
	}
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXPicture::get_Quality(short *pvar)
{
	*pvar = m_Image.GetJpegQuality();
	return S_OK;
}

STDMETHODIMP CXPicture::put_Quality(short var)
{
	if(var < 0 || var > 100)
		return DISP_E_BADCALLEE;

	m_Image.SetJpegQuality((BYTE)var);
	return S_OK;
}

STDMETHODIMP CXPicture::get_TransparentColor(long *pvar)
{
	*(RGBQUAD*)pvar = m_Image.GetTransColor();

	return S_OK;
}

STDMETHODIMP CXPicture::put_TransparentColor(long var)
{
	if((var & 0xffff0000) == 0xffff0000)
		var &= 0xffff;
	var &= 0xffffff;

	m_Image.SetTransColor(*(RGBQUAD*)&var);

	return S_OK;
}

STDMETHODIMP CXPicture::IsTransparent(VARIANT_BOOL *pvar)
{
	*pvar = m_Image.IsTransparent() ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}

STDMETHODIMP CXPicture::Resample(long nWidth, long nHeight, VARIANT varMode, IXPicture **pvar)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	HRESULT hr = S_OK;
	CXComPtr<CXPicture> pNewImage;
	int nMode = varGetNumbar(varMode, 1);

	pNewImage.CoCreateObject();

	m_cs.Enter();
	if(!m_Image.Resample(nWidth, nHeight, nMode, &pNewImage->m_Image))
		hr = CXObject::SetErrorInfo(m_Image.GetLastError());
	m_cs.Leave();

	if(FAILED(hr))return hr;

	return pNewImage.QueryInterface(pvar);
}

STDMETHODIMP CXPicture::Thumbnail(long nWidth, long nHeight, IXPicture **pvar)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	HRESULT hr = S_OK;
	CXComPtr<CXPicture> pNewImage;
	long l, t, w, h, iw, ih;

	pNewImage.CoCreateObject();

	m_cs.Enter();
	iw = m_Image.GetWidth();
	ih = m_Image.GetHeight();
	if(nWidth * ih > nHeight * iw)
	{
		h = nHeight * iw / nWidth;
		w = iw;
		l = 0;
		t = (ih - h) / 4;
	}else
	{
		h = ih;
		w = nWidth * ih / nHeight;
		l = (iw - w) / 2;
		t = 0;
	}

	if(!m_Image.Crop(l, t, w + l, h + t, &pNewImage->m_Image))
		hr = CXObject::SetErrorInfo(m_Image.GetLastError());
	else if (pNewImage->m_Image.GetBpp() < 24 && !pNewImage->m_Image.IncreaseBpp(24))
		hr = CXObject::SetErrorInfo(pNewImage->m_Image.GetLastError());
	else if(!pNewImage->m_Image.Resample(nWidth, nHeight, 3))
		hr = CXObject::SetErrorInfo(pNewImage->m_Image.GetLastError());
	else pNewImage->m_Image.info.dwType = 3;

	m_cs.Leave();

	if(FAILED(hr))return hr;

	return pNewImage.QueryInterface(pvar);
}

STDMETHODIMP CXPicture::Clone(IXPicture **pvar)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	CXComPtr<CXPicture> pNewImage;

	pNewImage.CoCreateObject();

	m_cs.Enter();
	pNewImage->m_Image.Copy(m_Image);
	m_cs.Leave();

	return pNewImage.QueryInterface(pvar);
}

STDMETHODIMP CXPicture::Crop(long left, long top, long right, long bottom)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	HRESULT hr = S_OK;

	m_cs.Enter();
	if(!m_Image.Crop(left, top, right, bottom))
		hr = CXObject::SetErrorInfo(m_Image.GetLastError());
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXPicture::Flip(VARIANT varDirection)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	int intDirection = varGetNumbar(varDirection, 1);
	HRESULT hr = S_OK;

	m_cs.Enter();
	if(intDirection == 2)
	{
		if(!m_Image.Flip())
			hr = CXObject::SetErrorInfo(m_Image.GetLastError());
	}else if(intDirection == 1)
	{
		if(!m_Image.Mirror())
			hr = CXObject::SetErrorInfo(m_Image.GetLastError());
	}
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXPicture::Resize(long nWidth, long nHeight, VARIANT varMode)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	int nMode = varGetNumbar(varMode, 1);
	HRESULT hr = S_OK;

	m_cs.Enter();
	if(!m_Image.Resample(nWidth, nHeight, nMode))
		hr = CXObject::SetErrorInfo(m_Image.GetLastError());
	m_cs.Leave();

	return hr;
}

STDMETHODIMP CXPicture::Rotate(long fAngle)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	HRESULT hr = S_OK;

	m_cs.Enter();
	if(!m_Image.Rotate(((float)-fAngle) / 10))
		hr = CXObject::SetErrorInfo(m_Image.GetLastError());
	m_cs.Leave();

	return hr;
}

STDMETHODIMP CXPicture::get_FillColor(long *pvar)
{
	*pvar = m_colocBrush;
	return S_OK;
}

STDMETHODIMP CXPicture::put_FillColor(long var)
{
	if((var & 0xffff0000) == 0xffff0000)
		var &= 0xffff;
	var &= 0xffffff;

	m_colocBrush = var;
	return S_OK;
}

STDMETHODIMP CXPicture::get_FillStyle(short *pvar)
{
	*pvar = m_nBrushStyle;
	return S_OK;
}

STDMETHODIMP CXPicture::put_FillStyle(short var)
{
	m_nBrushStyle = var;
	return S_OK;
}

STDMETHODIMP CXPicture::get_FontBold(VARIANT_BOOL *pvar)
{
	*pvar = (m_lfFont.lfWeight == FW_BOLD) ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}

STDMETHODIMP CXPicture::put_FontBold(VARIANT_BOOL var)
{
	m_lfFont.lfWeight = (var != VARIANT_FALSE) ? FW_BOLD : FW_NORMAL;
	return S_OK;
}

STDMETHODIMP CXPicture::get_FontBkColor(long *pvar)
{
	m_cs.Enter();
	*pvar = GetBkColor(m_Image.hMemDC);
	m_cs.Leave();
	return S_OK;
}

STDMETHODIMP CXPicture::put_FontBkColor(long var)
{
	if((var & 0xffff0000) == 0xffff0000)
		var &= 0xffff;
	var &= 0xffffff;

	m_cs.Enter();
	SetBkColor(m_Image.hMemDC, var);
	m_cs.Leave();
	return S_OK;
}

STDMETHODIMP CXPicture::get_FontBkMode(VARIANT_BOOL *pvar)
{
	m_cs.Enter();
	*pvar = (GetBkMode(m_Image.hMemDC) == OPAQUE) ? VARIANT_TRUE : VARIANT_FALSE;
	m_cs.Leave();
	return S_OK;
}

STDMETHODIMP CXPicture::put_FontBkMode(VARIANT_BOOL var)
{
	m_cs.Enter();
	SetBkMode(m_Image.hMemDC, var != VARIANT_FALSE ? OPAQUE : TRANSPARENT);
	m_cs.Leave();
	return S_OK;
}

STDMETHODIMP CXPicture::get_FontColor(long *pvar)
{
	m_cs.Enter();
	*pvar = GetTextColor(m_Image.hMemDC);
	m_cs.Leave();
	return S_OK;
}

STDMETHODIMP CXPicture::put_FontColor(long var)
{
	if((var & 0xffff0000) == 0xffff0000)
		var &= 0xffff;
	var &= 0xffffff;

	m_cs.Enter();
	SetTextColor(m_Image.hMemDC, var);
	m_cs.Leave();
	return S_OK;
}

STDMETHODIMP CXPicture::get_FontEscapement(long *pvar)
{
	*pvar = m_lfFont.lfEscapement;
	return S_OK;
}

STDMETHODIMP CXPicture::put_FontEscapement(long var)
{
	m_lfFont.lfEscapement = var;
	return S_OK;
}

STDMETHODIMP CXPicture::get_FontItalic(VARIANT_BOOL *pvar)
{
	*pvar = m_lfFont.lfItalic ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}

STDMETHODIMP CXPicture::put_FontItalic(VARIANT_BOOL var)
{
	m_lfFont.lfItalic = (var != VARIANT_FALSE);
	return S_OK;
}

STDMETHODIMP CXPicture::get_FontName(BSTR *pvar)
{
	CXString str;

	m_cs.Enter();
	str = m_lfFont.lfFaceName;
	m_cs.Leave();

	*pvar = str.AllocSysString();

	return S_OK;
}

STDMETHODIMP CXPicture::put_FontName(BSTR var)
{
	CXStringA stra(var);

	m_cs.Enter();
	strncpy_s(m_lfFont.lfFaceName, sizeof(m_lfFont.lfFaceName), stra, sizeof(m_lfFont.lfFaceName));
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXPicture::get_FontSize(short *pvar)
{
	*pvar = (short)m_lfFont.lfHeight;
	return S_OK;
}

STDMETHODIMP CXPicture::put_FontSize(short var)
{
	m_lfFont.lfHeight = var;

	return S_OK;
}

STDMETHODIMP CXPicture::get_FontUnderline(VARIANT_BOOL *pvar)
{
	*pvar = m_lfFont.lfUnderline ? VARIANT_TRUE : VARIANT_FALSE;
	return S_OK;
}

STDMETHODIMP CXPicture::put_FontUnderline(VARIANT_BOOL var)
{
	m_lfFont.lfUnderline = (var != VARIANT_FALSE);
	return S_OK;
}


STDMETHODIMP CXPicture::get_PenColor(long *pvar)
{
	*pvar = m_colorPen;
	return S_OK;
}

STDMETHODIMP CXPicture::put_PenColor(long var)
{
	if((var & 0xffff0000) == 0xffff0000)
		var &= 0xffff;
	var &= 0xffffff;

	m_colorPen = GetNearestColor(m_Image.hMemDC, var);
	return S_OK;
}

STDMETHODIMP CXPicture::get_PenStyle(short *pvar)
{
	*pvar = m_nPenStyle;
	return S_OK;
}

STDMETHODIMP CXPicture::put_PenStyle(short var)
{
	m_nPenStyle = var;
	return S_OK;
}

STDMETHODIMP CXPicture::get_PenWidth(short *pvar)
{
	*pvar = m_nPenWidth;
	return S_OK;
}

STDMETHODIMP CXPicture::put_PenWidth(short var)
{
	m_nPenWidth = var;
	return S_OK;
}

STDMETHODIMP CXPicture::get_TextAlignment(short *pvar)
{
	m_cs.Enter();
	*pvar = GetTextAlign(m_Image.hMemDC);
	m_cs.Leave();
	return S_OK;
}

STDMETHODIMP CXPicture::put_TextAlignment(short var)
{
	m_cs.Enter();
	SetTextAlign(m_Image.hMemDC, var);
	m_cs.Leave();
	return S_OK;
}

STDMETHODIMP CXPicture::get_Palette(VARIANT index, VARIANT* pvar)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	CComSafeArray<VARIANT> bstrArray;
	VARIANT* pVar;
	HRESULT hr;
	int n = varGetNumbar(index, -1);
	int s, i;
	RGBQUAD* pPal;

	m_cs.Enter();
	s = m_Image.GetPaletteSize() / sizeof(RGBQUAD);
	if(n < 0 || n > s - 1)
	{
		hr = bstrArray.Create(s);
		if(FAILED(hr))
		{
			m_cs.Leave();
			return hr;
		}

		pVar = (VARIANT*)bstrArray.m_psa->pvData;

		pPal = m_Image.GetPalette();

		for(i = 0; i < s; i ++)
		{
			pVar->lVal = 0;
			*(&pVar->bVal) = pPal[i].rgbRed;
			(&pVar->bVal)[1] = pPal[i].rgbGreen;
			(&pVar->bVal)[2] = pPal[i].rgbBlue;
			pVar->vt = VT_I4;
			pVar ++;
		}

		pvar->vt = VT_ARRAY | VT_VARIANT;
		pvar->parray = bstrArray.Detach();
	}else
	{
		pvar->lVal = 0;
		m_Image.GetPaletteColor(n, (&pvar->bVal), (&pvar->bVal) + 1, (&pvar->bVal) + 2);
		pvar->vt = VT_I4;
	}
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXPicture::put_Palette(VARIANT index, VARIANT var)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	int n = varGetNumbar(index, -1);
	int s, i;
	RGBQUAD* pPal;

	m_cs.Enter();
	s = m_Image.GetPaletteSize() / sizeof(RGBQUAD);
	if(n < 0 || n > s - 1)
	{
		VARIANT* pvar = &var;

		while(pvar->vt == (VT_VARIANT | VT_BYREF))
			pvar = pvar->pvarVal;

		if(pvar->vt != (VT_ARRAY | VT_VARIANT))
		{
			m_cs.Leave();
			return DISP_E_TYPEMISMATCH;;
		}

		pvar = (VARIANT*)pvar->parray->pvData;

		pPal = m_Image.GetPalette();

		for(i = 0; i < s; i ++)
		{
			pPal[i].rgbRed = *(&pvar->bVal);
			pPal[i].rgbGreen = (&pvar->bVal)[1];
			pPal[i].rgbBlue = (&pvar->bVal)[2];
			pvar ++;
		}
	}else
	{
		DWORD dwColor = varGetNumbar(var, -1);
		if((dwColor & 0xffff0000) == 0xffff0000)
			dwColor &= 0xffff;
		dwColor &= 0xffffff;

		m_Image.SetPaletteColor(n, *(COLORREF*)&dwColor);
	}
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXPicture::get_Pixel(long x, long y, long *pvar)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	int n;
	RGBQUAD c;

	m_cs.Enter();
	c = m_Image.GetPixelColor(x, m_Image.GetHeight() - y - 1);
	n = c.rgbBlue;
	c.rgbBlue = c.rgbRed;
	c.rgbRed = n;
	c.rgbReserved = 0;
	*(RGBQUAD*)pvar = c;
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXPicture::put_Pixel(long x, long y, long var)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	if((var & 0xffff0000) == 0xffff0000)
		var &= 0xffff;
	var &= 0xffffff;

	m_cs.Enter();
	SetPixel(m_Image.hMemDC, x, y, var);
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXPicture::GrayScale(void)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	m_cs.Enter();
	m_Image.GrayScale();
	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXPicture::Colorize(long color)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	int s, i;
	RGBQUAD* pPal;

	if((color & 0xffff0000) == 0xffff0000)
		color &= 0xffff;
	color &= 0xffffff;

	RGBQUAD rgb = {((BYTE*)&color)[2], ((BYTE*)&color)[1], *((BYTE*)&color)};
	RGBQUAD hsl = CxImage::RGBtoHSL(rgb);

	m_cs.Enter();

	s = m_Image.GetPaletteSize() / sizeof(RGBQUAD);
	pPal = m_Image.GetPalette();

	for(i = 0; i < s; i ++)
	{
		RGBQUAD hsl1 = CxImage::RGBtoHSL(pPal[i]);

		hsl1.rgbRed = hsl.rgbRed;
		hsl1.rgbGreen = hsl.rgbGreen;
		if(hsl1.rgbBlue < 128)hsl1.rgbBlue = hsl1.rgbBlue * hsl.rgbBlue / 128;
		else hsl1.rgbBlue = 255 - (255 - hsl1.rgbBlue) * (255 - hsl.rgbBlue) / 128;
		pPal[i] = CxImage::HSLtoRGB(hsl1);
	}

	m_cs.Leave();

	return S_OK;
}

STDMETHODIMP CXPicture::DrawChord(long x1, long y1, long x2, long y2, long fromDegree, long toDegree)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	double f, t;

	f = fromDegree * (3.1415926535 / 1800);
	t = toDegree * (3.1415926535 / 1800);

	CImagePaint p(this);

	Chord(m_Image.hMemDC, x1, y1, x2, y2, 
			(x1 + x2) / 2 + (long)floor(1024 * cos(f)), 
			(y1 + y2) / 2 - (long)floor(1024 * sin(f)), 
			(x1 + x2) / 2 + (long)floor(1024 * cos(t)), 
			(y1 + y2) / 2 - (long)floor(1024 * sin(t)));

	return S_OK;
}

STDMETHODIMP CXPicture::DrawArc(long x1, long y1, long x2, long y2, long fromDegree, long toDegree)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	double f, t;

	f = fromDegree * (3.1415926535 / 1800);
	t = toDegree * (3.1415926535 / 1800);

	CImagePaint p(this);

	Arc(m_Image.hMemDC, x1, y1, x2, y2, 
			(x1 + x2) / 2 + (long)floor(1024 * cos(f)), 
			(y1 + y2) / 2 - (long)floor(1024 * sin(f)), 
			(x1 + x2) / 2 + (long)floor(1024 * cos(t)), 
			(y1 + y2) / 2 - (long)floor(1024 * sin(t)));

	return S_OK;
}

STDMETHODIMP CXPicture::DrawEllipse(long x1, long y1, long x2, long y2)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	CImagePaint p(this);

	Ellipse(m_Image.hMemDC, x1, y1, x2, y2);

	return S_OK;
}

STDMETHODIMP CXPicture::DrawLine(long x1, long y1, long x2, long y2)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	CImagePaint p(this);

	MoveToEx(m_Image.hMemDC, x1, y1, NULL);
	LineTo(m_Image.hMemDC, x2, y2);

	return S_OK;
}

STDMETHODIMP CXPicture::DrawPie(long x1, long y1, long x2, long y2, long fromDegree, long toDegree)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	double f, t;

	f = fromDegree * (3.1415926535 / 1800);
	t = toDegree * (3.1415926535 / 1800);

	CImagePaint p(this);

	Pie(m_Image.hMemDC, x1, y1, x2, y2, 
			(x1 + x2) / 2 + (long)floor(1024 * cos(f)), 
			(y1 + y2) / 2 - (long)floor(1024 * sin(f)), 
			(x1 + x2) / 2 + (long)floor(1024 * cos(t)), 
			(y1 + y2) / 2 - (long)floor(1024 * sin(t)));

	return S_OK;
}

STDMETHODIMP CXPicture::DrawPolyBezier(VARIANT var)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	CImagePoly ps(var);

	if(ps.m_pPoint)
	{
        CImagePaint p(this);

		PolyBezier(m_Image.hMemDC, ps.m_pPoint, ps.m_nCount);
	}

	return ps.m_hr;
}

STDMETHODIMP CXPicture::DrawPolygon(VARIANT var)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	CImagePoly ps(var);

	if(ps.m_pPoint)
	{
        CImagePaint p(this);

		Polygon(m_Image.hMemDC, ps.m_pPoint, ps.m_nCount);
	}

	return ps.m_hr;
}

STDMETHODIMP CXPicture::DrawPolyline(VARIANT var)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	CImagePoly ps(var);

	if(ps.m_pPoint)
	{
        CImagePaint p(this);

		Polyline(m_Image.hMemDC, ps.m_pPoint, ps.m_nCount);
	}

	return ps.m_hr;
}

STDMETHODIMP CXPicture::DrawRectangle(long x1, long y1, long x2, long y2)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	CImagePaint p(this);

	Rectangle(m_Image.hMemDC, x1, y1, x2, y2);

	return S_OK;
}

STDMETHODIMP CXPicture::DrawText(long x, long y, BSTR strText)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	HFONT hFont;
	HFONT pOldFont;
	HRESULT hr = S_OK;

	hFont = CreateFontIndirect( &m_lfFont );

	m_cs.Enter();
	pOldFont = (HFONT)SelectObject(m_Image.hMemDC, hFont);
	TextOutW(m_Image.hMemDC, x, y, strText, ::SysStringLen(strText));
	if (pOldFont) SelectObject(m_Image.hMemDC, pOldFont);
	m_cs.Leave();

	DeleteObject(hFont);

	return hr;
}

STDMETHODIMP CXPicture::FloodFill(long x, long y)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	CImagePaint p(this);

	ExtFloodFill(m_Image.hMemDC, x, y, GetPixel(m_Image.hMemDC, x, y), FLOODFILLSURFACE);

	return S_OK;
}

STDMETHODIMP CXPicture::Paint(long x, long y, IXPicture *pImage)
{
	if(!m_Image.pDib)return DISP_E_BADCALLEE;

	CXComPtr<CXPicture> pSrcImage;

	pImage->QueryInterface(CLSID_XPicture, (void**)&pSrcImage);
	if(pSrcImage == NULL)return DISP_E_TYPEMISMATCH;

	pSrcImage->m_cs.Enter();
	m_cs.Enter();

	int mx, my;

	mx = pSrcImage->m_Image.head.biWidth;
	if(mx + x > m_Image.head.biWidth)
		mx = m_Image.head.biWidth - x;

	y = m_Image.head.biHeight - pSrcImage->m_Image.head.biHeight - y;

	my = pSrcImage->m_Image.head.biHeight;
	if(my + y > m_Image.head.biHeight)
		my = m_Image.head.biHeight - y;

	if(pSrcImage->m_Image.IsTransparent())
	{
		RGBQUAD c, c1;

		c1 = pSrcImage->m_Image.GetTransColor();
		for(int iy = 0; iy < my; iy ++)
			for(int ix = 0; ix < mx; ix++)
			{
				c = pSrcImage->m_Image.GetPixelColor(ix, iy);
				if(*(DWORD*)&c1 != *(DWORD*)&c)
					m_Image.SetPixelColor(x + ix, y + iy, c);
			}
	}else
	{
		for(int iy = 0; iy < my; iy ++)
			for(int ix = 0; ix < mx; ix++)
				m_Image.SetPixelColor(x + ix, y + iy, pSrcImage->m_Image.GetPixelColor(ix, iy));
	}

	m_cs.Leave();
	pSrcImage->m_cs.Leave();

	return S_OK;
}

