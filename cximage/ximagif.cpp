/*
 * File:	ximagif.cpp
 * Purpose:	Platform Independent GIF Image Class Loader and Writer
 * 07/Aug/2001 Davide Pizzolato - www.xdp.it
 * CxImage version 5.99c 17/Oct/2004
 */

#include "ximagif.h"

#if CXIMAGE_SUPPORT_GIF

#include "ximaiter.h"

#if CXIMAGE_SUPPORT_WINCE
	#define assert(s)
#else
	#include <assert.h>
#endif

#include "../gif/gif_lib.h"


int readCxFile(GifFileType *gFile, GifByteType * Buffer, int size)
{
	return ((CxFile*)gFile->UserData)->Read(Buffer, 1, size);
}

int writeCxFile(GifFileType *gFile, const GifByteType * Buffer, int size)
{
	return ((CxFile*)gFile->UserData)->Write(Buffer, 1, size);
}

////////////////////////////////////////////////////////////////////////////////
bool CxImageGIF::Decode(CxFile *fp)
{
	if (fp == NULL) return false;

	GifFileType *GifFile = NULL;
	GifRecordType RecordType;
	GifByteType *Extension;
	GifColorType *ColorMap;
	int i, j, Count, Row, Col, Width, Height, ExtCode, ColorMapSize;
	static int InterlacedOffset[] = { 0, 4, 2, 1 }, /* The way Interlaced image should. */
	    InterlacedJumps[] = { 8, 8, 4, 2 };    /* be read - offsets and jumps... */

	try
	{
		GifFile = DGifOpen(fp, readCxFile);

		if (!Create(GifFile->SWidth,GifFile->SHeight,8,CXIMAGE_FORMAT_GIF))
			throw "Can't allocate memory";

		do
		{
			DGifGetRecordType(GifFile, &RecordType);

			switch (RecordType)
			{
			case IMAGE_DESC_RECORD_TYPE:
				DGifGetImageDesc(GifFile);

				Row = GifFile->Image.Top; /* Image Position relative to Screen. */
				Col = GifFile->Image.Left;
				Width = GifFile->Image.Width;
				Height = GifFile->Image.Height;

				if (GifFile->Image.Left + GifFile->Image.Width > GifFile->SWidth ||
					GifFile->Image.Top + GifFile->Image.Height > GifFile->SHeight)
						throw "Image is not confined to screen dimension, aborted.\n";

				if (GifFile->Image.Interlace)
				{
					/* Need to perform 4 passes on the images: */
					for (Count = i = 0; i < 4; i++)
						for (j = Row + InterlacedOffset[i]; j < Row + Height; j += InterlacedJumps[i])
								DGifGetLine(GifFile, GetBits(GifFile->SHeight - j - 1) + Col, Width);
				}
				else {
					for (i = 0; i < Height; i++)
						DGifGetLine(GifFile,GetBits(GifFile->SHeight - Row++ - 1) + Col,Width);
			    }
				break;
			case EXTENSION_RECORD_TYPE:
				/* Skip any extension blocks in file: */
				DGifGetExtension(GifFile, &ExtCode, &Extension);
				while (Extension != NULL)
					DGifGetExtensionNext(GifFile, &Extension);
				break;
			default:
				break;
			}
		}while (RecordType != TERMINATE_RECORD_TYPE);

		ColorMap = (GifFile->Image.ColorMap ? GifFile->Image.ColorMap->Colors : GifFile->SColorMap->Colors);
		ColorMapSize = GifFile->Image.ColorMap ? GifFile->Image.ColorMap->ColorCount : GifFile->SColorMap->ColorCount;

		if(ColorMap && ColorMapSize)
		{
			RGBQUAD* ppal=GetPalette();

			for (i=0; i < ColorMapSize; i++) {
				ppal[i].rgbRed = ColorMap[i].Red;
				ppal[i].rgbGreen = ColorMap[i].Green;
				ppal[i].rgbBlue = ColorMap[i].Blue;
			}

			head.biClrUsed = ColorMapSize;

			if(GifFile->SBackGroundColor)
			{
				info.nBkgndIndex = GifFile->SBackGroundColor;
				info.nBkgndColor.rgbRed = ColorMap[info.nBkgndIndex].Red;
				info.nBkgndColor.rgbGreen = ColorMap[info.nBkgndIndex].Green;
				info.nBkgndColor.rgbBlue = ColorMap[info.nBkgndIndex].Blue;
			}
		}

		DGifCloseFile(GifFile);
		GifFile = NULL;
	} catch (int errid) {
		strncpy(info.szLastError,GifGetErrorMessage(errid),255);
		if(GifFile != NULL) DGifCloseFile(GifFile);
		return false;
	} catch (char *message) {
		strncpy(info.szLastError,message,255);
		if(GifFile != NULL) DGifCloseFile(GifFile);
		return false;
	}

	return true;
}

////////////////////////////////////////////////////////////////////////////////
#if CXIMAGE_SUPPORT_ENCODE
////////////////////////////////////////////////////////////////////////////////

bool CxImageGIF::Encode(CxFile * fp)
{
	if (EncodeSafeCheck(fp)) return false;

	GifFileType *GifFile = NULL;
	ColorMapObject *OutputColorMap = NULL;
	int i, ColorMapSize;

	if(head.biBitCount != 8)
	{
		if(head.biBitCount < 8)IncreaseBpp(8);
		else DecreaseBpp(8, true);
	}

	try
	{
		GifFile = EGifOpen(fp, writeCxFile);

		ColorMapSize = head.biClrUsed;
		OutputColorMap = MakeMapObject(ColorMapSize, NULL);

		RGBQUAD* pPal = GetPalette();
		for(i=0; i<ColorMapSize; ++i) 
		{
			OutputColorMap->Colors[i].Red = pPal[i].rgbRed;
			OutputColorMap->Colors[i].Green = pPal[i].rgbGreen;
			OutputColorMap->Colors[i].Blue = pPal[i].rgbBlue;
		}

		EGifPutScreenDesc(GifFile, head.biWidth, head.biHeight, OutputColorMap->ColorCount, info.nBkgndIndex== -1 ? 0 : info.nBkgndIndex, OutputColorMap);

		FreeMapObject(OutputColorMap);
		OutputColorMap = NULL;

		if(info.nBkgndIndex != -1)
		{
			unsigned char ExtStr[4] = { 1, 0, 0, info.nBkgndIndex };
			EGifPutExtension(GifFile, GRAPHICS_EXT_FUNC_CODE, 4, ExtStr);
		}

		EGifPutImageDesc(GifFile, 0, 0, head.biWidth, head.biHeight, FALSE, NULL);

		for (i = 0; i < head.biHeight; i++)
			EGifPutLine(GifFile, GetBits(head.biHeight - i - 1), head.biWidth);

		EGifCloseFile(GifFile);
		GifFile = NULL;
	} catch (int errid) {
		strncpy(info.szLastError,GifGetErrorMessage(errid),255);
		if(OutputColorMap)
		{
			FreeMapObject(OutputColorMap);
			OutputColorMap = NULL;
		}
		if(GifFile)
		{
			EGifCloseFile(GifFile);
			GifFile = NULL;
		}
		return false;
	} catch (char *message) {
		strncpy(info.szLastError,message,255);
		if(OutputColorMap)
		{
			FreeMapObject(OutputColorMap);
			OutputColorMap = NULL;
		}
		if(GifFile)
		{
			EGifCloseFile(GifFile);
			GifFile = NULL;
		}
		return false;
	}

	return true;
}
////////////////////////////////////////////////////////////////////////////////
bool CxImageGIF::Encode(CxFile * fp, CxImage ** pImages, int pagecount, bool bLocalColorMap)
{

	return true;
}
////////////////////////////////////////////////////////////////////////////////
#endif // CXIMAGE_SUPPORT_ENCODE
////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////////////
#endif // CXIMAGE_SUPPORT_GIF
