CREATE OR REPLACE PACKAGE KOCEL.BOOK
/*  KOCEL BOOK package
 create 08.02.2010
 by Nikolay Krasilnikov
 This file is distributed under Apache License 2.0 (01.2004).
   http://www.apache.org/licenses/
 update 17.01.2024
*****************************************************************************
*/

as
/*  Releasing package variables.
*/
procedure ClearData;
/*  Uploading and preparing data.
*/
procedure PrepareData(UBOOK in VARCHAR2,  MAX_ROWS in NUMBER default 1000);

/*  Provides the parameters of the sheets of the book "UBOOK" for the request.
 If the sheet number is 0 and the name is null, then this is a global parameter of the book.
*/
function Pars return KOCEL.TSHEET_PARS pipelined;

/*  Provides the formats used in the UBOOK book for the request.
 The "Fmt" field contains the format identifier, and the "Tag" field contains the sequence number.
 The first entry contains the number of formats in the "Tag" field.
*/
function Formats return KOCEL.TFORMATS pipelined;

/*  Provides the "UBOOK" book sheet numbers for the request.
 The first entry does not contain a number, but a number of sheets.
 Information about the book sheets and their names is stored in the package.
*/
function Sheets return KOCEL.TSHEET_NUMS pipelined;

/*  Provides formats and row heights of the "USHEET" sheet from the "UBOOK" book for the query
.
 The "Fmt" field does not contain the parameter ID, but the number,
 corresponding to the value of the "Tag" field in the result of the request to
the "SHEET_FORMATS" function for this book.
*/
function Sheet_Row_s(USHEET in VARCHAR2) return KOCEL.TROW_DATAS pipelined;

/*  Provides the formats and heights of the rows of sheets of the book "UBOOK" for the query.
 The "Fmt" field does not contain the parameter ID, but the number,
 corresponding to the value of the "Tag" field as a result of the request to
the "SHEET_FORMATS" function for this book.
*/
function Row_s return KOCEL.TROW_DATAS pipelined;

/*  Provides the formats and column widths of the "USHEET" sheet from the "UBOOK" book for the query
.
 The "Fmt" field does not contain the parameter ID, but the number,
 corresponding to the value of the "Tag" field as a result of the request to
the "SHEET_FORMATS" function for this book.
*/
function Sheet_Column_s(USHEET in VARCHAR2) return KOCEL.TCOLUMN_DATAS
pipelined;

/*  Provides the formats and column widths of the UBOOK book sheets for the query.
 The "Fmt" field does not contain the parameter ID, but the number,
 corresponding to the value of the "Tag" field as a result of the request to
the "SHEET_FORMATS" function for this book.
*/
function Column_s return KOCEL.TCOLUMN_DATAS pipelined;

/*  Provides the data of the "USHEET" sheet from the "UBOOK" book for the request.
 The "Fmt" field does not contain the parameter ID, but the number,
 corresponding to the value of the "Tag" field as a result of the request to
the "SHEET_FORMATS" function for this book.
*/
function Sheet_Cells(USHEET in VARCHAR2) return KOCEL.TCELL_DATAS pipelined;

/*  Provides the data of the book "UBOOK" for the request.
 The "Fmt" field does not contain the parameter ID, but the number,
 corresponding to the value of the "Tag" field as a result of the request to
the "SHEET_FORMATS" function for this book.
*/
function Cells return KOCEL.TCELL_DATAS pipelined;

/*  Provides the combined cells of the "USHEET" sheet from the "UBOOK" workbook for the query
.
*/
function Sheet_MCells(USHEET in VARCHAR2) return KOCEL.TMCELL_DATAS pipelined;

/*  Provides the combined cells of the sheets of the book "UBOOK" for the query.
*/
function MCells return KOCEL.TMCELL_DATAS pipelined;

END;
/
