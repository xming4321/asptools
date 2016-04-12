/*
 * File:	ximagif.h
 * Purpose:	GIF Image Class Loader and Writer
 */
/* ==========================================================
 * CxImageGIF (c) 07/Aug/2001 Davide Pizzolato - www.xdp.it
 * For conditions of distribution and use, see copyright notice in ximage.h
 *
 * Special thanks to Troels Knakkergaard for new features, enhancements and bugfixes
 *
 * original CImageGIF  and CImageIterator implementation are:
 * Copyright:	(c) 1995, Alejandro Aguilar Sierra <asierra(at)servidor(dot)unam(dot)mx>
 *
 * 6/15/97 Randy Spann: Added GIF87a writing support
 *         R.Spann@ConnRiver.net
 *
 * DECODE.C - An LZW decoder for GIF
 * Copyright (C) 1987, by Steven A. Bennett
 * Copyright (C) 1994, C++ version by Alejandro Aguilar Sierra
 *
 * In accordance with the above, I want to credit Steve Wilhite who wrote
 * the code which this is heavily inspired by...
 *
 * GIF and 'Graphics Interchange Format' are trademarks (tm) of
 * Compuserve, Incorporated, an H&R Block Company.
 *
 * Release Notes: This file contains a decoder routine for GIF images
 * which is similar, structurally, to the original routine by Steve Wilhite.
 * It is, however, somewhat noticably faster in most cases.
 *
 * ==========================================================
 */

#if !defined(__ximaGIF_h)
#define __ximaGIF_h

#include "ximage.h"

#if CXIMAGE_SUPPORT_GIF


class CImageIterator;
class DLL_EXP CxImageGIF: public CxImage
{
public:
	CxImageGIF(): CxImage(CXIMAGE_FORMAT_GIF) {}

	bool Decode(CxFile * fp);
	bool Decode(FILE *fp) { CxIOFile file(fp); return Decode(&file); }

#if CXIMAGE_SUPPORT_ENCODE
	bool Encode(CxFile * fp);
	bool Encode(CxFile * fp, CxImage ** pImages, int pagecount, bool bLocalColorMap = false);
	bool Encode(FILE *fp) { CxIOFile file(fp); return Encode(&file); }
	bool Encode(FILE *fp, CxImage ** pImages, int pagecount, bool bLocalColorMap = false)
				{ CxIOFile file(fp); return Encode(&file, pImages, pagecount, bLocalColorMap); }
#endif // CXIMAGE_SUPPORT_ENCODE
};

#endif

#endif
