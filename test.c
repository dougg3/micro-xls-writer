/*
 * MicroXLSWriter test program using stdio
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

#include <stdio.h>
#include "microxlswriter.h"

/** Write callback used with MicroXLSWriter to write to a stdio FILE *
 *
 * @param ptr Data pointer for the callback. Is actually a FILE *
 * @param data The data to write
 * @param len The number of bytes to write
 * @return Result code indicating success or failure
 */
static MicroXLSWriterResult WriteToFile(void *ptr, const uint8_t *data, size_t len)
{
	FILE *f = (FILE *)ptr;
	size_t result = fwrite(data, len, 1, f);
	if (result == 1) {
		return MXW_SUCCESS;
	} else {
		return MXW_IO_ERROR;
	}
}

/** Main program
 *
 * @param argc Number of arguments, including the program itself (unused)
 * @param argv Array of arguments (unused)
 * @return 0 on success, nonzero on failure
 */
int main(int argc, char *argv[])
{
	// Open a couple of output files and set up writers for each of them
	FILE *f1 = fopen("test1.xls", "wb");
	MicroXLSWriter w1;
	w1.writeFunc = WriteToFile;
	w1.ptr = f1;

	FILE *f2 = fopen("test2.xls", "wb");
	MicroXLSWriter w2;
	w2.writeFunc = WriteToFile;
	w2.ptr = f2;

	// Begin both spreadsheets
	MicroXLS_Begin(&w1);
	MicroXLS_Begin(&w2);

	// Do some things with the first spreadsheet
	MicroXLS_SetColumnWidth(&w1, 2, 256*100);
	MicroXLS_AddNumberCell(&w1, 0, 0, 12345.6);
	MicroXLS_AddLabelCell(&w1, 0, 1, "Testing", 7);

	// Do something with the second spreadsheet (it won't interfere with the first!)
	MicroXLS_AddNumberCell(&w2, 5, 5, 555.5);

	// Do more stuff with the first spreadsheet
	MicroXLS_AddLabelCell(&w1, 0, 2, "Testing much longer cell content", 32);
	MicroXLS_AddNumberCell(&w1, 0, 255, 3.141592);

	// Finish both spreadsheets
	MicroXLS_Finish(&w1);
	MicroXLS_Finish(&w2);

	// Close the files to complete the process
	fclose(f1);
	fclose(f2);

	// Note that I ignored result codes of MicroXLS_* functions. For your production code it would probably be wise
	// to check the return result and handle errors appropriately.

	return 0;
}
