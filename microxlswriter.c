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

#include "microxlswriter.h"
#include <string.h>

/// Macro used for ensuring the supplied writer is valid before calling it
#define MXW_CHECKWRITER(w) \
	do { \
		if (!w || !w->writeFunc) { \
			return MXW_INVALID_PARAM; \
		} \
	} while (0)

/// Macro used for calling the write callback and returning an error code if necessary
#define MXW_WRITE(w, rec, recLen) \
	do { \
		MicroXLSWriterResult r = w->writeFunc(w->ptr, rec, recLen); \
		if (r != MXW_SUCCESS) { \
			return r; \
		} \
	} while (0)

MicroXLSWriterResult MicroXLS_Begin(MicroXLSWriter *w)
{
	// Check to make sure a valid writer was passed in
	MXW_CHECKWRITER(w);

	// Write out a BIFF2 BOF record
	const uint8_t bofRec[] = {
		0x09, 0x00,	// Record type = 0x0009 = BOF
		0x04, 0x00,	// Length = 4
		0x02, 0x00,	// BIFF version = 0x0002 = BIFF2
		0x10, 0x00	// Data type = 0x0010 = Sheet
	};

	MXW_WRITE(w, bofRec, sizeof(bofRec));
	return MXW_SUCCESS;
}

MicroXLSWriterResult MicroXLS_AddNumberCell(MicroXLSWriter *w, uint16_t row, uint16_t col, double val)
{
	// Check to make sure a valid writer was passed in
	MXW_CHECKWRITER(w);

	// Write out a BIFF2 NUMBER record
	uint8_t numRec[] = {
		0x03, 0x00,			// Record type = 0x0003 = NUMBER
		0x0F, 0x00,			// Length = 15
		0x00, 0x00,			// Row (to be filled in)
		0x00, 0x00,			// Column (to be filled in)
		0x00, 0x00, 0x00,	// Cell attributes (leave at 0, is probably invalid)
		0x00, 0x00,			// 64-bit double floating point value
		0x00, 0x00,
		0x00, 0x00,
		0x00, 0x00
	};

	memcpy(&numRec[4], &row, sizeof(row));
	memcpy(&numRec[6], &col, sizeof(col));
	memcpy(&numRec[11], &val, sizeof(val));
	MXW_WRITE(w, numRec, sizeof(numRec));
	return MXW_SUCCESS;
}

MicroXLSWriterResult MicroXLS_AddLabelCell(MicroXLSWriter *w, uint16_t row, uint16_t col, const char *label, uint8_t len)
{
	// Check to make sure a valid writer was passed in
	MXW_CHECKWRITER(w);

	// Write out a BIFF2 LABEL record
	uint8_t labelRec[] = {
		0x04, 0x00,			// Record type = 0x0004 = LABEL
		0x00, 0x00,			// Length (to be filled in)
		0x00, 0x00,			// Row (to be filled in)
		0x00, 0x00,			// Column (to be filled in)
		0x00, 0x00, 0x00,	// Cell attributes (leave at 0, is probably invalid)
		0x00				// Label length (0-255)
		// The label data itself will come after this
	};

	// Total record length is the label length plus the BIFF record length,
	// but not counting the first four common bytes of the BIFF record.
	uint16_t recLen = len + sizeof(labelRec) - 4;
	memcpy(&labelRec[2], &recLen, sizeof(recLen));
	memcpy(&labelRec[4], &row, sizeof(row));
	memcpy(&labelRec[6], &col, sizeof(col));
	memcpy(&labelRec[11], &len, sizeof(len));
	// Write the beginning of the record...
	MXW_WRITE(w, labelRec, sizeof(labelRec));
	// ...followed by the rest of the record which is the label data itself.
	MXW_WRITE(w, (const uint8_t *)label, len);
	return MXW_SUCCESS;
}

MicroXLSWriterResult MicroXLS_SetColumnWidth(MicroXLSWriter *w, uint8_t col, uint16_t width)
{
	// Check to make sure a valid writer was passed in
	MXW_CHECKWRITER(w);

	// Write out a BIFF2 COLWIDTH record
	uint8_t colWidthRec[] = {
		0x24, 0x00,			// Record type = 0x0024 = COLWIDTH
		0x04, 0x00,			// Length = 4
		0x00,				// Index to first column (to be filled in)
		0x00,				// Index to last column (to be filled in)
		0x00, 0x00			// Width in 1/256 width of the zero character, using default font
	};

	colWidthRec[4] = col;
	colWidthRec[5] = col;
	memcpy(&colWidthRec[6], &width, sizeof(width));
	MXW_WRITE(w, colWidthRec, sizeof(colWidthRec));
	return MXW_SUCCESS;
}

MicroXLSWriterResult MicroXLS_Finish(MicroXLSWriter *w)
{
	// Check to make sure a valid writer was passed in
	MXW_CHECKWRITER(w);

	// Write out a BIFF2 EOF record
	const uint8_t eofRec[] = {
		0x0A, 0x00,	// Record type = 0x000A = EOF
		0x00, 0x00	// Length = 0
	};

	MXW_WRITE(w, eofRec, sizeof(eofRec));
	return MXW_SUCCESS;
}
