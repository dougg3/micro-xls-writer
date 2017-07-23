/*
 * MicroXLSWriter - a minimal XLS file writer suitable for use with microcontrollers
 *
 * Copyright (c) 2017, Doug Brown
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions are met:
 *
 * 1. Redistributions of source code must retain the above copyright notice, this
 *    list of conditions and the following disclaimer.
 * 2. Redistributions in binary form must reproduce the above copyright notice,
 *    this list of conditions and the following disclaimer in the documentation
 *    and/or other materials provided with the distribution.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
 * ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
 * WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
 * DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT OWNER OR CONTRIBUTORS BE LIABLE FOR
 * ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
 * (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
 * LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
 * ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
 * (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
 * SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
 */

#ifndef MICROXLSWRITER_H_
#define MICROXLSWRITER_H_

#include <stdint.h>
#include <stdlib.h>

/// Result code returned from MicroXLSWriter functions
typedef enum MicroXLSWriterResult
{
	MXW_SUCCESS = 0,        //!< Returned if an operation completes successfully
	MXW_IO_ERROR = -1,      //!< Returned if an I/O error occurred while writing the spreadsheet
	MXW_INVALID_PARAM = -2, //!< An invalid parameter was supplied to the function
} MicroXLSWriterResult;

/// Structure used for keeping track of an XLS file being written. Provides a way for
/// this library to interface with your code through callbacks.
typedef struct MicroXLSWriter
{
	/** Function pointer that will be called whenever data needs to be written to an output file.
	 *
	 * @param ptr Filled in with the "ptr" value of this struct
	 * @param data Pointer to the data to write to the output file
	 * @param len The number of bytes in "data" that should be written to the output file
	 * @return Result code (your function should return MXW_SUCCESS if the write succeeded)
	 */
	MicroXLSWriterResult (*writeFunc)(void *ptr, const uint8_t *data, size_t len);
	/// A pointer that you can use for distinguishing between multiple open files (if needed)
	void *ptr;
} MicroXLSWriter;

/** Begins writing a new XLS file
 *
 * @param w The writer
 * @return Result code. MXW_SUCCESS on success, otherwise the result code indicates the failure reason.
 */
MicroXLSWriterResult MicroXLS_Begin(MicroXLSWriter *w);

/** Sets the width of the column at the specified index.
 *
 * @param w The writer
 * @param col The index of the column (0 to 255 represent columns A through IV)
 * @param width The width of the column, in 1/256 of the width of the '0' character in the default font
 * @return Result code. MXW_SUCCESS on success, otherwise the result code indicates the failure reason.
 *
 * Note that the BIFF2 format used by this library is limited to setting custom column widths for the first
 * 256 columns. It's not possible to set individual custom widths past column 255. It's also not guaranteed
 * that columns above 255 are even supported in the file format; Excel 2.1 didn't seem to support them.
 */
MicroXLSWriterResult MicroXLS_SetColumnWidth(MicroXLSWriter *w, uint8_t col, uint16_t width);

/** Adds a number cell to the spreadsheet.
 *
 * @param w The writer
 * @param row The index of the row (0 represents the first row)
 * @param col The index of the column (0 represents column A, 1 represents column B, and so on)
 * @param val The value to put into the cell
 * @return Result code. MXW_SUCCESS on success, otherwise the result code indicates the failure reason.
 */
MicroXLSWriterResult MicroXLS_AddNumberCell(MicroXLSWriter *w, uint16_t row, uint16_t col, double val);

/** Adds a label (string value) cell to the spreadsheet.
 *
 * @param w The writer
 * @param row The index of the row (0 represents the first row)
 * @param col The index of the column (0 represents column A, 1 represents column B, and so on)
 * @param label Pointer to the label data, which can be a maximum of 255 bytes in length.
 * @param len The length of the label.
 * @return Result code. MXW_SUCCESS on success, otherwise the result code indicates the failure reason.
 */
MicroXLSWriterResult MicroXLS_AddLabelCell(MicroXLSWriter *w, uint16_t row, uint16_t col, const char *label, uint8_t len);

/** Finishes writing the XLS file.
 *
 * @param w The writer
 * @return Result code. MXW_SUCCESS on success, otherwise the result code indicates the failure reason.
 */
MicroXLSWriterResult MicroXLS_Finish(MicroXLSWriter *w);

#endif /* MICROXLSWRITER_H_ */
