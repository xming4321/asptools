/*
 * File:	ximajpg.h
 * Purpose:	JPG Image Class Loader and Writer
 */
/* ==========================================================
 * CxImageJPG (c) 07/Aug/2001 Davide Pizzolato - www.xdp.it
 * For conditions of distribution and use, see copyright notice in ximage.h
 *
 * Special thanks to Troels Knakkergaard for new features, enhancements and bugfixes
 *
 * Special thanks to Chris Shearer Cooper for CxFileJpg tips & code
 *
 * EXIF support based on jhead-1.8 by Matthias Wandel <mwandel(at)rim(dot)net>
 *
 * original CImageJPG  and CImageIterator implementation are:
 * Copyright:	(c) 1995, Alejandro Aguilar Sierra <asierra(at)servidor(dot)unam(dot)mx>
 *
 * This software is based in part on the work of the Independent JPEG Group.
 * Copyright (C) 1991-1998, Thomas G. Lane.
 * ==========================================================
 */
#if !defined(__ximaJPEG_h)
#define __ixmaJPEG_h

#include "ximage.h"

#if CXIMAGE_SUPPORT_JPG

extern "C" {
 #include "../jpeg/jpeglib.h"
 #include "../jpeg/jerror.h"
}

class DLL_EXP CxImageJPG: public CxImage
{
public:
	CxImageJPG();
	~CxImageJPG();

//	bool Load(const char * imageFileName){ return CxImage::Load(imageFileName,CXIMAGE_FORMAT_JPG);}
//	bool Save(const char * imageFileName){ return CxImage::Save(imageFileName,CXIMAGE_FORMAT_JPG);}
	bool Decode(CxFile * hFile);
	bool Decode(FILE *hFile) { CxIOFile file(hFile); return Decode(&file); }

#if CXIMAGE_SUPPORT_ENCODE
	bool Encode(CxFile * hFile);
	bool Encode(FILE *hFile) { CxIOFile file(hFile); return Encode(&file); }
#endif // CXIMAGE_SUPPORT_ENCODE


////////////////////////////////////////////////////////////////////////////////////////
//////////////////////        C x F i l e J p g         ////////////////////////////////
////////////////////////////////////////////////////////////////////////////////////////

// thanks to Chris Shearer Cooper <cscooper(at)frii(dot)com>
class CxFileJpg : public jpeg_destination_mgr, public jpeg_source_mgr
	{
public:
	enum { eBufSize = 4096 };

	CxFileJpg(CxFile* pFile)
	{
        m_pFile = pFile;

		init_destination = InitDestination;
		empty_output_buffer = EmptyOutputBuffer;
		term_destination = TermDestination;

		init_source = InitSource;
		fill_input_buffer = FillInputBuffer;
		skip_input_data = SkipInputData;
		resync_to_restart = jpeg_resync_to_restart; // use default method
		term_source = TermSource;
		next_input_byte = NULL; //* => next byte to read from buffer 
		bytes_in_buffer = 0;	//* # of bytes remaining in buffer 

		m_pBuffer = new unsigned char[eBufSize];
	}
	~CxFileJpg()
	{
		delete [] m_pBuffer;
	}

	static void InitDestination(j_compress_ptr cinfo)
	{
		CxFileJpg* pDest = (CxFileJpg*)cinfo->dest;
		pDest->next_output_byte = pDest->m_pBuffer;
		pDest->free_in_buffer = eBufSize;
	}

	static boolean EmptyOutputBuffer(j_compress_ptr cinfo)
	{
		CxFileJpg* pDest = (CxFileJpg*)cinfo->dest;
		if (pDest->m_pFile->Write(pDest->m_pBuffer,1,eBufSize)!=(size_t)eBufSize)
			ERREXIT(cinfo, JERR_FILE_WRITE);
		pDest->next_output_byte = pDest->m_pBuffer;
		pDest->free_in_buffer = eBufSize;
		return TRUE;
	}

	static void TermDestination(j_compress_ptr cinfo)
	{
		CxFileJpg* pDest = (CxFileJpg*)cinfo->dest;
		size_t datacount = eBufSize - pDest->free_in_buffer;
		/* Write any data remaining in the buffer */
		if (datacount > 0) {
			if (!pDest->m_pFile->Write(pDest->m_pBuffer,1,datacount))
				ERREXIT(cinfo, JERR_FILE_WRITE);
		}
		pDest->m_pFile->Flush();
		/* Make sure we wrote the output file OK */
		if (pDest->m_pFile->Error()) ERREXIT(cinfo, JERR_FILE_WRITE);
		return;
	}

	static void InitSource(j_decompress_ptr cinfo)
	{
		CxFileJpg* pSource = (CxFileJpg*)cinfo->src;
		pSource->m_bStartOfFile = TRUE;
	}

	static boolean FillInputBuffer(j_decompress_ptr cinfo)
	{
		size_t nbytes;
		CxFileJpg* pSource = (CxFileJpg*)cinfo->src;
		nbytes = pSource->m_pFile->Read(pSource->m_pBuffer,1,eBufSize);
		if (nbytes <= 0){
			if (pSource->m_bStartOfFile)	//* Treat empty input file as fatal error 
				ERREXIT(cinfo, JERR_INPUT_EMPTY);
			WARNMS(cinfo, JWRN_JPEG_EOF);
			// Insert a fake EOI marker 
			pSource->m_pBuffer[0] = (JOCTET) 0xFF;
			pSource->m_pBuffer[1] = (JOCTET) JPEG_EOI;
			nbytes = 2;
		}
		pSource->next_input_byte = pSource->m_pBuffer;
		pSource->bytes_in_buffer = nbytes;
		pSource->m_bStartOfFile = FALSE;
		return TRUE;
	}

	static void SkipInputData(j_decompress_ptr cinfo, long num_bytes)
	{
		CxFileJpg* pSource = (CxFileJpg*)cinfo->src;
		if (num_bytes > 0){
			while (num_bytes > (long)pSource->bytes_in_buffer){
				num_bytes -= (long)pSource->bytes_in_buffer;
				FillInputBuffer(cinfo);
				// note we assume that fill_input_buffer will never return FALSE,
				// so suspension need not be handled.
			}
			pSource->next_input_byte += (size_t) num_bytes;
			pSource->bytes_in_buffer -= (size_t) num_bytes;
		}
	}

	static void TermSource(j_decompress_ptr cinfo)
	{
		return;
	}
protected:
    CxFile  *m_pFile;
	unsigned char *m_pBuffer;
	bool m_bStartOfFile;
};

public:
	enum CODEC_OPTION
	{
		ENCODE_BASELINE = 0x1,
		ENCODE_ARITHMETIC = 0x2,
		ENCODE_GRAYSCALE = 0x4,
		ENCODE_OPTIMIZE = 0x8,
		ENCODE_PROGRESSIVE = 0x10,
		ENCODE_LOSSLESS = 0x20,
		ENCODE_SMOOTHING = 0x40,
		DECODE_GRAYSCALE = 0x80,
		DECODE_QUANTIZE = 0x100,
		DECODE_DITHER = 0x200,
		DECODE_ONEPASS = 0x400,
		DECODE_NOSMOOTH = 0x800
	}; 

	int m_nPredictor;
	int m_nPointTransform;
	int m_nSmoothing;
	int m_nQuantize;
	J_DITHER_MODE m_nDither;

};

#endif

#endif
