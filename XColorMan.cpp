// XColorMan.cpp : Implementation of CXColorMan

#include "stdafx.h"
#include "XColorMan.h"


// CXColorMan

////////////////////////////////////////////////////////////////////////////////
#define  HSLMAX   255	/* H,L, and S vary over 0-HSLMAX */
#define  RGBMAX   255   /* R,G, and B vary over 0-RGBMAX */
                        /* HSLMAX BEST IF DIVISIBLE BY 6 */
                        /* RGBMAX, HSLMAX must each fit in a BYTE. */
/* Hue is undefined if Saturation is 0 (grey-scale) */
/* This value determines where the Hue scrollbar is */
/* initially set for achromatic colors */
#define HSLUNDEFINED (HSLMAX*2/3)

void CXColorMan::toHSL()
{
	BYTE cMax,cMin;				/* max and min RGB values */
	WORD Rdelta,Gdelta,Bdelta;	/* intermediate value: % of spread from max*/

	cMax = (BYTE)max( max(r,g), b);	/* calculate lightness */
	cMin = (BYTE)min( min(r,g), b);
	l = (BYTE)((((cMax+cMin)*HSLMAX)+RGBMAX)/(2*RGBMAX));

	if (cMax==cMin){			/* r=g=b --> achromatic case */
		s = 0;					/* saturation */
		h = HSLUNDEFINED;		/* hue */
	} else {					/* chromatic case */
		if (l <= (HSLMAX/2))	/* saturation */
			s = (BYTE)((((cMax-cMin)*HSLMAX)+((cMax+cMin)/2))/(cMax+cMin));
		else
			s = (BYTE)((((cMax-cMin)*HSLMAX)+((2*RGBMAX-cMax-cMin)/2))/(2*RGBMAX-cMax-cMin));
		/* hue */
		Rdelta = (WORD)((((cMax-r)*(HSLMAX/6)) + ((cMax-cMin)/2) ) / (cMax-cMin));
		Gdelta = (WORD)((((cMax-g)*(HSLMAX/6)) + ((cMax-cMin)/2) ) / (cMax-cMin));
		Bdelta = (WORD)((((cMax-b)*(HSLMAX/6)) + ((cMax-cMin)/2) ) / (cMax-cMin));

		if (r == cMax)
			h = (BYTE)(Bdelta - Gdelta);
		else if (g == cMax)
			h = (BYTE)((HSLMAX/3) + Rdelta - Bdelta);
		else /* B == cMax */
			h = (BYTE)(((2*HSLMAX)/3) + Gdelta - Rdelta);

		if (h > HSLMAX) h -= HSLMAX;
	}
}

float inline HueToRGB(float n1,float n2, float hue)
{
	//<F. Livraghi> fixed implementation for HSL2RGB routine
	float rValue;

	if (hue > 360)
		hue = hue - 360;
	else if (hue < 0)
		hue = hue + 360;

	if (hue < 60)
		rValue = n1 + (n2-n1)*hue/60.0f;
	else if (hue < 180)
		rValue = n2;
	else if (hue < 240)
		rValue = n1+(n2-n1)*(240-hue)/60;
	else
		rValue = n1;

	return rValue;
}

void CXColorMan::toRGB()
{
	float H, S, L;
	float m1,m2;

	H = (float)h * 360.0f/255.0f;
	S = (float)s/255.0f;
	L = (float)l/255.0f;

	if (L <= 0.5)	m2 = L * (1+S);
	else			m2 = L + S - L*S;

	m1 = 2 * L - m2;

	if (S == 0) {
		r=g=b=(BYTE)(L*255.0f);
	} else {
		r = (BYTE)(HueToRGB(m1,m2,H+120) * 255.0f);
		g = (BYTE)(HueToRGB(m1,m2,H) * 255.0f);
		b = (BYTE)(HueToRGB(m1,m2,H-120) * 255.0f);
	}
}

STDMETHODIMP CXColorMan::get_R(SHORT* pVal)
{
	*pVal = r;

	return S_OK;
}

STDMETHODIMP CXColorMan::put_R(SHORT newVal)
{
	r = min(255,max(0,newVal));
	toHSL();

	return S_OK;
}

STDMETHODIMP CXColorMan::get_G(SHORT* pVal)
{
	*pVal = g;

	return S_OK;
}

STDMETHODIMP CXColorMan::put_G(SHORT newVal)
{
	g = min(255,max(0,newVal));
	toHSL();

	return S_OK;
}

STDMETHODIMP CXColorMan::get_B(SHORT* pVal)
{
	*pVal = b;

	return S_OK;
}

STDMETHODIMP CXColorMan::put_B(SHORT newVal)
{
	b = min(255,max(0,newVal));
	toHSL();

	return S_OK;
}

STDMETHODIMP CXColorMan::get_H(SHORT* pVal)
{
	*pVal = h;

	return S_OK;
}

STDMETHODIMP CXColorMan::put_H(SHORT newVal)
{
	h = min(255,max(0,newVal));
	toRGB();

	return S_OK;
}

STDMETHODIMP CXColorMan::get_S(SHORT* pVal)
{
	*pVal = s;

	return S_OK;
}

STDMETHODIMP CXColorMan::put_S(SHORT newVal)
{
	s = min(255,max(0,newVal));
	toRGB();

	return S_OK;
}

STDMETHODIMP CXColorMan::get_L(SHORT* pVal)
{
	*pVal = l;

	return S_OK;
}

STDMETHODIMP CXColorMan::put_L(SHORT newVal)
{
	l = min(255,max(0,newVal));
	toRGB();

	return S_OK;
}

STDMETHODIMP CXColorMan::get_RGB(LONG* pVal)
{
	RGBQUAD rgb = {(BYTE)b, (BYTE)g, (BYTE)r};

	*pVal = ((LONG)b << 16) + ((LONG)g << 8) + r;

	return S_OK;
}

STDMETHODIMP CXColorMan::put_RGB(LONG newVal)
{
	if((newVal & 0xffff0000) == 0xffff0000)
		newVal &= 0xffff;
	newVal &= 0xffffff;

	r = ((BYTE*)&newVal)[0];
	g = ((BYTE*)&newVal)[1];
	b = ((BYTE*)&newVal)[2];
	toHSL();

	return S_OK;
}

STDMETHODIMP CXColorMan::get_RGBString(BSTR* pVal)
{
	static char hexChar[] = "0123456789ABCDEF";
	WCHAR strRGB[7];

	strRGB[0] = hexChar[r / 16];
	strRGB[1] = hexChar[r % 16];
	strRGB[2] = hexChar[g / 16];
	strRGB[3] = hexChar[g % 16];
	strRGB[4] = hexChar[b / 16];
	strRGB[5] = hexChar[b % 16];
	strRGB[6] = 0;

	*pVal = ::SysAllocString(strRGB);

	return S_OK;
}

static struct _colorName{
	const WCHAR* name; UINT color;
}s_colors[] = 
{
	{L"aliceblue", 0xF0F8FF}, 
	{L"antiquewhite", 0xFAEBD7}, 
	{L"aqua", 0x00FFFF}, 
	{L"aquamarine", 0x7FFFD4}, 
	{L"azure", 0xF0FFFF}, 
	{L"beige", 0xF5F5DC}, 
	{L"bisque", 0xFFE4C4}, 
	{L"black", 0x000000}, 
	{L"blanchedalmond", 0xFFEBCD}, 
	{L"blue", 0x0000FF}, 
	{L"blueviolet", 0x8A2BE2}, 
	{L"brown", 0xA52A2A}, 
	{L"burlywood", 0xDEB887}, 
	{L"cadetblue", 0x5F9EA0}, 
	{L"chartreuse", 0x7FFF00}, 
	{L"chocolate", 0xD2691E}, 
	{L"coral", 0xFF7F50}, 
	{L"cornflowerblue", 0x6495ED}, 
	{L"cornsilk", 0xFFF8DC}, 
	{L"crimson", 0xDC143C}, 
	{L"cyan", 0x00FFFF}, 
	{L"darkblue", 0x00008B}, 
	{L"darkcyan", 0x008B8B}, 
	{L"darkgoldenrod", 0xB8860B}, 
	{L"darkgray", 0xA9A9A9}, 
	{L"darkgreen", 0x006400}, 
	{L"darkkhaki", 0xBDB76B}, 
	{L"darkmagenta", 0x8B008B}, 
	{L"darkolivegreen", 0x556B2F}, 
	{L"darkorange", 0xFF8C00}, 
	{L"darkorchid", 0x9932CC}, 
	{L"darkred", 0x8B0000}, 
	{L"darksalmon", 0xE9967A}, 
	{L"darkseagreen", 0x8FBC8B}, 
	{L"darkslateblue", 0x483D8B}, 
	{L"darkslategray", 0x2F4F4F}, 
	{L"darkturquoise", 0x00CED1}, 
	{L"darkviolet", 0x9400D3}, 
	{L"deeppink", 0xFF1493}, 
	{L"deepskyblue", 0x00BFFF}, 
	{L"dimgray", 0x696969}, 
	{L"dodgerblue", 0x1E90FF}, 
	{L"firebrick", 0xB22222}, 
	{L"floralwhite", 0xFFFAF0}, 
	{L"forestgreen", 0x228B22}, 
	{L"fuchsia", 0xFF00FF}, 
	{L"gainsboro", 0xDCDCDC}, 
	{L"ghostwhite", 0xF8F8FF}, 
	{L"gold", 0xFFD700}, 
	{L"goldenrod", 0xDAA520}, 
	{L"gray", 0x808080}, 
	{L"green", 0x008000}, 
	{L"greenyellow", 0xADFF2F}, 
	{L"honeydew", 0xF0FFF0}, 
	{L"hotpink", 0xFF69B4}, 
	{L"indianred", 0xCD5C5C}, 
	{L"indigo", 0x4B0082}, 
	{L"ivory", 0xFFFFF0}, 
	{L"khaki", 0xF0E68C}, 
	{L"lavender", 0xE6E6FA}, 
	{L"lavenderblush", 0xFFF0F5}, 
	{L"lawngreen", 0x7CFC00}, 
	{L"lemonchiffon", 0xFFFACD}, 
	{L"lightblue", 0xADD8E6}, 
	{L"lightcoral", 0xF08080}, 
	{L"lightcyan", 0xE0FFFF}, 
	{L"lightgoldenrodyellow", 0xFAFAD2}, 
	{L"lightgreen", 0x90EE90}, 
	{L"lightgrey", 0xD3D3D3}, 
	{L"lightpink", 0xFFB6C1}, 
	{L"lightsalmon", 0xFFA07A}, 
	{L"lightseagreen", 0x20B2AA}, 
	{L"lightskyblue", 0x87CEFA}, 
	{L"lightslategray", 0x778899}, 
	{L"lightsteelblue", 0xB0C4DE}, 
	{L"lightyellow", 0xFFFFE0}, 
	{L"lime", 0x00FF00}, 
	{L"limegreen", 0x32CD32}, 
	{L"linen", 0xFAF0E6}, 
	{L"magenta", 0xFF00FF}, 
	{L"maroon", 0x800000}, 
	{L"mediumaquamarine", 0x66CDAA}, 
	{L"mediumblue", 0x0000CD}, 
	{L"mediumorchid", 0xBA55D3}, 
	{L"mediumpurple", 0x9370DB}, 
	{L"mediumseagreen", 0x3CB371}, 
	{L"mediumslateblue", 0x7B68EE}, 
	{L"mediumspringgreen", 0x00FA9A}, 
	{L"mediumturquoise", 0x48D1CC}, 
	{L"mediumvioletred", 0xC71585}, 
	{L"midnightblue", 0x191970}, 
	{L"mintcream", 0xF5FFFA}, 
	{L"mistyrose", 0xFFE4E1}, 
	{L"moccasin", 0xFFE4B5}, 
	{L"navajowhite", 0xFFDEAD}, 
	{L"navy", 0x000080}, 
	{L"oldlace", 0xFDF5E6}, 
	{L"olive", 0x808000}, 
	{L"olivedrab", 0x6B8E23}, 
	{L"orange", 0xFFA500}, 
	{L"orangered", 0xFF4500}, 
	{L"orchid", 0xDA70D6}, 
	{L"palegoldenrod", 0xEEE8AA}, 
	{L"palegreen", 0x98FB98}, 
	{L"paleturquoise", 0xAFEEEE}, 
	{L"palevioletred", 0xDB7093}, 
	{L"papayawhip", 0xFFEFD5}, 
	{L"peachpuff", 0xFFDAB9}, 
	{L"peru", 0xCD853F}, 
	{L"pink", 0xFFC0CB}, 
	{L"plum", 0xDDA0DD}, 
	{L"powderblue", 0xB0E0E6}, 
	{L"purple", 0x800080}, 
	{L"red", 0xFF0000}, 
	{L"rosybrown", 0xBC8F8F}, 
	{L"royalblue", 0x4169E1}, 
	{L"saddlebrown", 0x8B4513}, 
	{L"salmon", 0xFA8072}, 
	{L"sandybrown", 0xF4A460}, 
	{L"seagreen", 0x2E8B57}, 
	{L"seashell", 0xFFF5EE}, 
	{L"sienna", 0xA0522D}, 
	{L"silver", 0xC0C0C0}, 
	{L"skyblue", 0x87CEEB}, 
	{L"slateblue", 0x6A5ACD}, 
	{L"slategray", 0x708090}, 
	{L"snow", 0xFFFAFA}, 
	{L"springgreen", 0x00FF7F}, 
	{L"steelblue", 0x4682B4}, 
	{L"tan", 0xD2B48C}, 
	{L"teal", 0x008080}, 
	{L"thistle", 0xD8BFD8}, 
	{L"tomato", 0xFF6347}, 
	{L"turquoise", 0x40E0D0}, 
	{L"violet", 0xEE82EE}, 
	{L"wheat", 0xF5DEB3}, 
	{L"white", 0xFFFFFF}, 
	{L"whitesmoke", 0xF5F5F5}, 
	{L"yellow", 0xFFFF00}, 
	{L"yellowgreen", 0x9ACD32}
};

static int compareKey( WCHAR **arg1, _colorName *arg2 )
{
	return _wcsicmp( *arg1, arg2->name );
}

STDMETHODIMP CXColorMan::put_RGBString(BSTR newVal)
{
	UINT n = 0, pos = 0, len = SysStringLen(newVal);
	WCHAR ch1, ch2;
	struct _colorName* pName = NULL;

	if(len > 0)
	{
		if(newVal[0] != '#')
		{
			pName = (struct _colorName*)bsearch((void*)&newVal, (void*)&s_colors, sizeof(s_colors) / sizeof(s_colors[0]), sizeof(s_colors[0]), (int (*)(const void*, const void*))compareKey);
			if(pName)
				n = pName->color;
		}

		if(!pName)
			while(pos < len)
			{
				ch1 = newVal[pos++];

				if((ch1 >= 'a' && ch1 <= 'f') || (ch1 >= 'A' && ch1 <= 'F'))
					ch1 = (ch1 & 0xf) + 9;
				else if(ch1 >= '0' && ch1 <= '9')
					ch1 &= 0xf;
				else continue;

				if(pos < len)
				{
					ch2 = newVal[pos++];

					if((ch2 >= 'a' && ch2 <= 'f') || (ch2 >= 'A' && ch2 <= 'F'))
						ch2 = (ch2 & 0xf) + 9;
					else if(ch2 >= '0' && ch2 <= '9')
						ch2 &= 0xf;
					else
					{
						ch2 = ch1;
						ch1 = 0;
					}
				}else
				{
					ch2 = ch1;
					ch1 = 0;
				}

				n = (n << 8) + (ch1 << 4) + ch2;
			}
	}

	r = ((BYTE*)&n)[2];
	g = ((BYTE*)&n)[1];
	b = ((BYTE*)&n)[0];
	toHSL();

	return S_OK;
}

