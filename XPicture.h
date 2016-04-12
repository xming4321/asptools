// XPicture.h : Declaration of the CXPicture

#pragma once
#include "resource.h"       // main symbols

#include "asptools.h"
#include "atlib.h"

#include "CxImage\ximage.h"

// CXPicture

class ATL_NO_VTABLE CXPicture : 
	public CComObjectRootEx<CComMultiThreadModel>,
	public CComCoClass<CXPicture, &CLSID_XPicture>,
	public IDispatchImpl<IXPicture, &IID_IXPicture, &LIBID_asptoolsLib, /*wMajor =*/ 1, /*wMinor =*/ 0>
{
public:
	CXPicture();

DECLARE_REGISTRY_RESOURCEID(IDR_XPICTURE)


BEGIN_COM_MAP(CXPicture)
	COM_INTERFACE_ENTRY_IID(CLSID_XPicture, CXPicture)
	COM_INTERFACE_ENTRY(IXPicture)
	COM_INTERFACE_ENTRY(IDispatch)
	COM_INTERFACE_ENTRY_AGGREGATE(IID_IMarshal, m_pUnkMarshaler.p)
END_COM_MAP()


	DECLARE_PROTECT_FINAL_CONSTRUCT()
	DECLARE_GET_CONTROLLING_UNKNOWN()

	HRESULT FinalConstruct()
	{
		return CoCreateFreeThreadedMarshaler(
			GetControllingUnknown(), &m_pUnkMarshaler.p);
	}
	
	void FinalRelease() 
	{
		m_pUnkMarshaler.Release();
	}

	CComPtr<IUnknown> m_pUnkMarshaler;

private:
	class CImagePaint
	{
	public:
		CImagePaint(CXPicture* pImage);
		~CImagePaint();

	private:
		CXPicture* m_pImage;
		HPEN m_hPen;
		HPEN m_hOldPen;
		HBRUSH m_hBrush;
		HBRUSH m_hOldBrush;
	};

	class CImagePoly
	{
	public:
		CImagePoly(VARIANT var);
		~CImagePoly();

		LPPOINT m_pPoint;
		int m_nCount;
		HRESULT m_hr;
	};

public:
	STDMETHOD(Load)(BSTR strFile, VARIANT varType);
	STDMETHOD(Save)(BSTR strFile, VARIANT varType);
	STDMETHOD(LoadFromData)(VARIANT varData, VARIANT varType);
	STDMETHOD(GetImageData)(VARIANT varType, VARIANT* pVal);
	STDMETHOD(LoadRawData)(VARIANT varData);
	STDMETHOD(GetRawData)(VARIANT* pVal);
	STDMETHOD(LoadInfo)(BSTR strFile, VARIANT varType);
	STDMETHOD(LoadInfoFromData)(VARIANT varData, VARIANT varType);
	STDMETHOD(Create)(long nWidth, long nHeight, short Bpp);
	STDMETHOD(Clear)(void);
	STDMETHOD(get_Type)(short *pvar);
	STDMETHOD(get_Height)(long *pvar);
	STDMETHOD(get_Width)(long *pvar);
	STDMETHOD(get_Bpp)(short *pvar);
	STDMETHOD(put_Bpp)(short var);
	STDMETHOD(get_Quality)(short *pvar);
	STDMETHOD(put_Quality)(short var);
	STDMETHOD(get_TransparentColor)(long *pvar);
	STDMETHOD(put_TransparentColor)(long var);
	STDMETHOD(IsTransparent)(VARIANT_BOOL *pvar);
	STDMETHOD(Resample)(long nWidth, long nHeight, VARIANT varMode, IXPicture **pvar);
	STDMETHOD(Thumbnail)(long nWidth, long nHeight, IXPicture **pvar);
	STDMETHOD(Clone)(IXPicture **pvar);
	STDMETHOD(Crop)(long left, long top, long right, long bottom);
	STDMETHOD(Flip)(VARIANT varDirection);
	STDMETHOD(Resize)(long nWidth, long nHeight, VARIANT varMode);
	STDMETHOD(Rotate)(long fAngle);
	STDMETHOD(GrayScale)(void);
	STDMETHOD(Colorize)(long color);
	STDMETHOD(get_FillColor)(long *pvar);
	STDMETHOD(put_FillColor)(long var);
	STDMETHOD(get_FillStyle)(short *pvar);
	STDMETHOD(put_FillStyle)(short var);
	STDMETHOD(get_FontBold)(VARIANT_BOOL *pvar);
	STDMETHOD(put_FontBold)(VARIANT_BOOL var);
	STDMETHOD(get_FontBkColor)(long *pvar);
	STDMETHOD(put_FontBkColor)(long var);
	STDMETHOD(get_FontBkMode)(VARIANT_BOOL *pvar);
	STDMETHOD(put_FontBkMode)(VARIANT_BOOL var);
	STDMETHOD(get_FontColor)(long *pvar);
	STDMETHOD(put_FontColor)(long var);
	STDMETHOD(get_FontEscapement)(long *pvar);
	STDMETHOD(put_FontEscapement)(long var);
	STDMETHOD(get_FontItalic)(VARIANT_BOOL *pvar);
	STDMETHOD(put_FontItalic)(VARIANT_BOOL var);
	STDMETHOD(get_FontName)(BSTR *pvar);
	STDMETHOD(put_FontName)(BSTR var);
	STDMETHOD(get_FontSize)(short *pvar);
	STDMETHOD(put_FontSize)(short var);
	STDMETHOD(get_FontUnderline)(VARIANT_BOOL *pvar);
	STDMETHOD(put_FontUnderline)(VARIANT_BOOL var);
	STDMETHOD(get_PenColor)(long *pvar);
	STDMETHOD(put_PenColor)(long var);
	STDMETHOD(get_PenStyle)(short *pvar);
	STDMETHOD(put_PenStyle)(short var);
	STDMETHOD(get_PenWidth)(short *pvar);
	STDMETHOD(put_PenWidth)(short var);
	STDMETHOD(get_TextAlignment)(short *pvar);
	STDMETHOD(put_TextAlignment)(short var);
	STDMETHOD(get_Palette)(VARIANT index, VARIANT* pvar);
	STDMETHOD(put_Palette)(VARIANT index, VARIANT var);
	STDMETHOD(get_Pixel)(long x, long y, long *pvar);
	STDMETHOD(put_Pixel)(long x, long y, long var);
	STDMETHOD(DrawArc)(long x1, long y1, long x2, long y2, long fromDegree, long toDegree);
	STDMETHOD(DrawChord)(long x1, long y1, long x2, long y2, long fromDegree, long toDegree);
	STDMETHOD(DrawEllipse)(long x1, long y1, long x2, long y2);
	STDMETHOD(DrawLine)(long x1, long y1, long x2, long y2);
	STDMETHOD(DrawPie)(long x1, long y1, long x2, long y2, long fromDegree, long toDegree);
	STDMETHOD(DrawPolyBezier)(VARIANT var);
	STDMETHOD(DrawPolygon)(VARIANT var);
	STDMETHOD(DrawPolyline)(VARIANT var);
	STDMETHOD(DrawRectangle)(long x1, long y1, long x2, long y2);
	STDMETHOD(DrawText)(long x, long y, BSTR strText);
	STDMETHOD(FloodFill)(long x, long y);
	STDMETHOD(Paint)(long x, long y, IXPicture *pImage);

private:
	CCriticalSection m_cs;
	CxImage m_Image;
	LOGFONT m_lfFont;
	short m_align;
	COLORREF m_colocBrush;
	short m_nBrushStyle;
	COLORREF m_colorPen;
	short m_nPenWidth;
	short m_nPenStyle;
};

OBJECT_ENTRY_AUTO(__uuidof(XPicture), CXPicture)
