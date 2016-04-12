/*****************************************************************************
 *   "Gif-Lib" - Yet another gif library.                     
 *                                         
 * Written by:  Gershon Elber            IBM PC Ver 0.1,    Jun. 1989    
 *****************************************************************************
 * Handle error reporting for the GIF library.                     
 *****************************************************************************
 * History:                                     
 * 17 Jun 89 - Version 1.0 by Gershon Elber.                     
 ****************************************************************************/

#ifdef HAVE_CONFIG_H
#include <config.h>
#endif

#include <stdio.h>
#include "gif_lib.h"

/*****************************************************************************
 * Print the last GIF error to stderr.                         
 ****************************************************************************/
const char * GifGetErrorMessage(int _GifError)
{
    const char * Err;

    switch (_GifError) {
      case E_GIF_ERR_OPEN_FAILED:
        Err = "GIF-LIB error: Failed to open given file";
        break;
      case E_GIF_ERR_WRITE_FAILED:
        Err = "GIF-LIB error: Failed to Write to given file";
        break;
      case E_GIF_ERR_HAS_SCRN_DSCR:
        Err = "GIF-LIB error: Screen Descriptor already been set";
        break;
      case E_GIF_ERR_HAS_IMAG_DSCR:
        Err = "GIF-LIB error: Image Descriptor is still active";
        break;
      case E_GIF_ERR_NO_COLOR_MAP:
        Err = "GIF-LIB error: Neither Global Nor Local color map";
        break;
      case E_GIF_ERR_DATA_TOO_BIG:
        Err = "GIF-LIB error: #Pixels bigger than Width * Height";
        break;
      case E_GIF_ERR_NOT_ENOUGH_MEM:
        Err = "GIF-LIB error: Fail to allocate required memory";
        break;
      case E_GIF_ERR_DISK_IS_FULL:
        Err = "GIF-LIB error: Write failed (disk full?)";
        break;
      case E_GIF_ERR_CLOSE_FAILED:
        Err = "GIF-LIB error: Failed to close given file";
        break;
      case E_GIF_ERR_NOT_WRITEABLE:
        Err = "GIF-LIB error: Given file was not opened for write";
        break;
      case D_GIF_ERR_OPEN_FAILED:
        Err = "GIF-LIB error: Failed to open given file";
        break;
      case D_GIF_ERR_READ_FAILED:
        Err = "GIF-LIB error: Failed to Read from given file";
        break;
      case D_GIF_ERR_NOT_GIF_FILE:
        Err = "GIF-LIB error: Given file is NOT GIF file";
        break;
      case D_GIF_ERR_NO_SCRN_DSCR:
        Err = "GIF-LIB error: No Screen Descriptor detected";
        break;
      case D_GIF_ERR_NO_IMAG_DSCR:
        Err = "GIF-LIB error: No Image Descriptor detected";
        break;
      case D_GIF_ERR_NO_COLOR_MAP:
        Err = "GIF-LIB error: Neither Global Nor Local color map";
        break;
      case D_GIF_ERR_WRONG_RECORD:
        Err = "GIF-LIB error: Wrong record type detected";
        break;
      case D_GIF_ERR_DATA_TOO_BIG:
        Err = "GIF-LIB error: #Pixels bigger than Width * Height";
        break;
      case D_GIF_ERR_NOT_ENOUGH_MEM:
        Err = "GIF-LIB error: Fail to allocate required memory";
        break;
      case D_GIF_ERR_CLOSE_FAILED:
        Err = "GIF-LIB error: Failed to close given file";
        break;
      case D_GIF_ERR_NOT_READABLE:
        Err = "GIF-LIB error: Given file was not opened for read";
        break;
      case D_GIF_ERR_IMAGE_DEFECT:
        Err = "GIF-LIB error: Image is defective, decoding aborted";
        break;
      case D_GIF_ERR_EOF_TOO_SOON:
        Err = "GIF-LIB error: Image EOF detected, before image complete";
        break;
      default:
        Err = "GIF-LIB error: undefined error";
        break;
    }

	return Err;
}
