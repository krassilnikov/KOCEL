/*  KOCEL Types
 create 17.04.2009
 by Nikolay Krasilnikov
 This file is distributed under Apache License 2.0 (01.2004).
   http://www.apache.org/licenses/
 update 11.02.2010 17.08.2011 28.08.2014 08.07.2015 05.07.2016 15.11.2022
        23.11.2022 23.12.2023
 
----------------------------------------------------------------------------
*/
BEGIN
  EXECUTE IMMEDIATE('
CREATE OR REPLACE TYPE KOCEL.TDATES 
/* The date table.*/
/* KOCEL.TYPES.sql*/
IS TABLE OF DATE
                   ');
END;
/
GRANT execute on KOCEL.TDATES to public;

/* ----------------------------------------------------------------------------
*/
BEGIN
  EXECUTE IMMEDIATE('
CREATE OR REPLACE TYPE KOCEL.TNUMS 
/* A table of numbers.*/
/* KOCEL.TYPES.sql*/
IS TABLE OF NUMBER
                   ');
END;
/
GRANT execute on KOCEL.TNUMS to public;

/* ----------------------------------------------------------------------------
*/
BEGIN
  EXECUTE IMMEDIATE('
CREATE OR REPLACE TYPE KOCEL.TSTRINGS 
/* A table of rows.*/
/* KOCEL.TYPES.sql*/
IS TABLE OF VARCHAR2(32000)
                   ');
END;
/
GRANT execute on KOCEL.TSTRINGS to public;

/* ----------------------------------------------------------------------------
*/
CREATE OR REPLACE TYPE KOCEL.TVALUE as OBJECT (
/* The universal value of the data cell.*/
/* KOCEL.TYPES.sql*/
/* Row number.*/
R NUMBER(6),
/* The number of the column.*/
C NUMBER(3),
/* A field for a numeric value.*/
N FLOAT(49),
/* A field for a date type value.*/
D DATE,
/* A field for a string type value.*/
S VARCHAR2(32000),
/* The field for the formula.*/
F VARCHAR2(4000),

MEMBER FUNCTION T return VARCHAR2,
MEMBER PROCEDURE ClearData(self in out KOCEL.TVALUE),
MEMBER FUNCTION as_Str return VARCHAR2,
CONSTRUCTOR FUNCTION TValue
return SELF AS RESULT,
CONSTRUCTOR FUNCTION TValue(CellValue in NUMBER)
return SELF AS RESULT,
CONSTRUCTOR FUNCTION TValue(CellValue in DATE)
return SELF AS RESULT,
CONSTRUCTOR FUNCTION TValue(CellValue in VARCHAR2)
return SELF AS RESULT,
CONSTRUCTOR FUNCTION TValue(N in NUMBER, D in DATE,S in VARCHAR2,
                            F in VARCHAR2 default null)
return SELF AS RESULT,
MAP MEMBER FUNCTION map_values(self in out KOCEL.TVALUE) return VARCHAR2
);
/
GRANT execute on KOCEL.TVALUE to public;

/* ----------------------------------------------------------------------------
*/
CREATE OR REPLACE TYPE KOCEL.TROW as OBJECT (
/* A series of data. This object is designed to work with tables consisting
of no more than 25 columns.*/
/* KOCEL.TYPES.sql*/
A KOCEL.TVALUE,
B KOCEL.TVALUE,
C KOCEL.TVALUE,
D KOCEL.TVALUE,
E KOCEL.TVALUE,
F KOCEL.TVALUE,
G KOCEL.TVALUE,
H KOCEL.TVALUE,
I KOCEL.TVALUE,
J KOCEL.TVALUE,
K KOCEL.TVALUE,
L KOCEL.TVALUE,
M KOCEL.TVALUE,
N KOCEL.TVALUE,
O KOCEL.TVALUE,
P KOCEL.TVALUE,
Q KOCEL.TVALUE,
R KOCEL.TVALUE,
S KOCEL.TVALUE,
T KOCEL.TVALUE,
U KOCEL.TVALUE,
V KOCEL.TVALUE,
W KOCEL.TVALUE,
X KOCEL.TVALUE,
Y KOCEL.TVALUE,
Z KOCEL.TVALUE,
CONSTRUCTOR FUNCTION TRow
return SELF AS RESULT,
MEMBER PROCEDURE ClearData(self in out KOCEL.TROW)
);
/
GRANT execute on KOCEL.TROW to public;

/* ----------------------------------------------------------------------------
*/
CREATE OR REPLACE TYPE KOCEL.TSROW as OBJECT (
/* A series of lines. This object is designed to work with tables consisting
of no more than 25 columns.*/
/* KOCEL.TYPES.sql*/
A VARCHAR2(32000),
B VARCHAR2(32000),
C VARCHAR2(32000),
D VARCHAR2(32000),
E VARCHAR2(32000),
F VARCHAR2(32000),
G VARCHAR2(32000),
H VARCHAR2(32000),
I VARCHAR2(32000),
J VARCHAR2(32000),
K VARCHAR2(32000),
L VARCHAR2(32000),
M VARCHAR2(32000),
N VARCHAR2(32000),
O VARCHAR2(32000),
P VARCHAR2(32000),
Q VARCHAR2(32000),
R VARCHAR2(32000),
S VARCHAR2(32000),
T VARCHAR2(32000),
U VARCHAR2(32000),
V VARCHAR2(32000),
W VARCHAR2(32000),
X VARCHAR2(32000),
Y VARCHAR2(32000),
Z VARCHAR2(32000),
CONSTRUCTOR FUNCTION TSROW
return SELF AS RESULT,
MEMBER PROCEDURE ClearData(self in out KOCEL.TSROW)
);
/
GRANT execute on KOCEL.TSROW to public;

/* ----------------------------------------------------------------------------
*/
BEGIN
  EXECUTE IMMEDIATE('
CREATE OR REPLACE TYPE KOCEL.TCELLS 
/* A table of rows.*/
/* KOCEL.TYPES.sql*/
IS TABLE OF KOCEL.TROW
                   ');
END;
/
GRANT execute on KOCEL.TCELLS to public;

/* ----------------------------------------------------------------------------
*/
BEGIN
  EXECUTE IMMEDIATE('
CREATE OR REPLACE TYPE KOCEL.TSCELLS 
/* A table of rows of rows.*/
/* KOCEL.TYPES.sql*/
IS TABLE OF KOCEL.TSROW
                   ');
END;
/
GRANT execute on KOCEL.TSCELLS to public;

/* ----------------------------------------------------------------------------
*/
CREATE or REPLACE TYPE KOCEL.TTableDef 
AS OBJECT (
/* An object describing a field in a database table.*/
/* KOCEL.TYPES.sql*/
/* The name of the field.*/
ColName VARCHAR2(40),
/* The name of the data type for this field.*/
ColTypeName VARCHAR2(30),
/* The data type for this field.*/
ColType NUMBER(9),
/* The size of the data for this field.*/
ColLength NUMBER(9),
/* Accuracy.*/
ColPrecision NUMBER(9),
/* Scale.*/
ColScale NUMBER(9)
);
/
GRANT execute on KOCEL.TTableDef to public;

/* ----------------------------------------------------------------------------
*/
BEGIN
  EXECUTE IMMEDIATE('
CREATE OR REPLACE TYPE KOCEL.TTableDefs 
/* A table containing the structure of a database table.*/
/* KOCEL.TYPES.sql*/
IS TABLE OF KOCEL.TTableDef
                   ');
END;
/
GRANT execute on KOCEL.TTableDefs to public;
 
/* ----------------------------------------------------------------------------
*/
BEGIN
  EXECUTE IMMEDIATE('
CREATE OR REPLACE TYPE KOCEL.TFORMATS 
/* A table containing descriptions of cell formats.*/
/* KOCEL.TYPES.sql*/
IS TABLE OF KOCEL.TFORMAT
                   ');
END;
/
GRANT execute on KOCEL.TFormats to public;

/* ----------------------------------------------------------------------------
*/
CREATE or REPLACE TYPE KOCEL.TCELL_DATA 
AS OBJECT (
/* An object containing the complete data of a workbook cell.*/
/* KOCEL.TYPES.sql*/
/* ROWID as a string.*/
  RID VARCHAR2(36),
/* Data type. "N" is a number, "D" is a date, "S" is a string.*/
  T CHAR(1),
/* Row number.*/
  R NUMBER(6),
/* The line number.*/
	C NUMBER(3),
/* A numeric value.*/
	N FLOAT(49),
/* The value of the date type.*/
  D DATE,
/* The value of the string type.*/
	S VARCHAR2(32000),
/* Formula.*/
	F VARCHAR2(4000),
/* The serial number of the sheet.*/
  SHEET_NUM NUMBER(3),
/* The name of the sheet.*/
  SHEET VARCHAR2(255),
/* The name of the book.*/
  BOOK VARCHAR2(255),
/* The format number.*/
	Fmt NUMBER,
  CONSTRUCTOR FUNCTION TCELL_DATA return SELF AS RESULT
);
/
GRANT execute on KOCEL.TCELL_DATA to public;

/* ----------------------------------------------------------------------------
*/
BEGIN
  EXECUTE IMMEDIATE('
CREATE OR REPLACE TYPE KOCEL.TCELL_DATAS 
/* A table containing the complete data of the book cells.*/
/* KOCEL.TYPES.sql*/
IS TABLE OF KOCEL.TCELL_DATA
                   ');
END;
/
GRANT execute on KOCEL.TCELL_DATAS to public;

/* ----------------------------------------------------------------------------
*/
CREATE or REPLACE TYPE KOCEL.TCOLUMN_DATA 
AS OBJECT (
/* An object containing column description data for a book sheet.*/
/* KOCEL.TYPES.sql*/
/* ROWID as a string.*/
  RID VARCHAR2(36),
/* The number of the column.*/
	C NUMBER(3),
/* Column width.*/
	W FLOAT(49),
/* The sequential number of the sheet in the book.*/
  SHEET_NUM NUMBER(3),
/* The name of the sheet.*/
  SHEET VARCHAR2(255),
/* The name of the book.*/
  BOOK VARCHAR2(255),
/* The format number.*/
	Fmt NUMBER,
  CONSTRUCTOR FUNCTION TCOLUMN_DATA return SELF AS RESULT
);
/
GRANT execute on KOCEL.TCOLUMN_DATA to public;

/* ----------------------------------------------------------------------------
*/
BEGIN
  EXECUTE IMMEDIATE('
CREATE OR REPLACE TYPE KOCEL.TCOLUMN_DATAS 
/* A table containing complete column description data for a book sheet.*/
/* KOCEL.TYPES.sql*/
IS TABLE OF KOCEL.TCOLUMN_DATA
                   ');
END;
/
GRANT execute on KOCEL.TCOLUMN_DATAS to public;

/* ----------------------------------------------------------------------------
*/
CREATE or REPLACE TYPE KOCEL.TROW_DATA 
AS OBJECT (
/* An object containing row description data for a book sheet.*/
/* KOCEL.TYPES.sql*/
/* ROWID as a string.*/
  RID VARCHAR2(36),
/* Row number.*/
  R NUMBER(6),
/* The height of the row.*/
	H FLOAT(49),
/* The sequential number of the sheet in the book.*/
  SHEET_NUM NUMBER(3),
/* The name of the sheet in the book.*/
  SHEET VARCHAR2(255),
/* The name of the book.*/
  BOOK VARCHAR2(255),
/* The format number.*/
	Fmt NUMBER,
  CONSTRUCTOR FUNCTION TROW_DATA return SELF AS RESULT
);
/
GRANT execute on KOCEL.TROW_DATA to public;

/* ----------------------------------------------------------------------------
*/
BEGIN
  EXECUTE IMMEDIATE('
CREATE OR REPLACE TYPE KOCEL.TROW_DATAS 
/* A table containing complete row description data for a book sheet.*/
/* KOCEL.TYPES.sql*/
IS TABLE OF KOCEL.TROW_DATA
                   ');
END;
/
GRANT execute on KOCEL.TROW_DATAS to public;

/* ----------------------------------------------------------------------------
*/
CREATE or REPLACE TYPE KOCEL.TMCELL_DATA 
AS OBJECT (
/* An object containing data describing a combined cell on a book sheet.*/
/* KOCEL.TYPES.sql*/
/* The left column from which the combined cell begins.*/
  L NUMBER(3),
/* The top row from which the combined cell starts.*/
  T NUMBER(6),
/* The right column where the combined cell ends.*/
  R NUMBER(3),
/* The bottom row where the combined cell ends.*/
  B NUMBER(6),
/* The sequential number of the sheet in the book.*/
  SHEET_NUM NUMBER(3),
/* The name of the sheet.*/
  SHEET VARCHAR2(255),
/* The name of the book.*/
  BOOK VARCHAR2(255),
  CONSTRUCTOR FUNCTION TMCELL_DATA return SELF AS RESULT
);
/
GRANT execute on KOCEL.TMCELL_DATA to public;

/* ----------------------------------------------------------------------------
*/
BEGIN
  EXECUTE IMMEDIATE('
CREATE OR REPLACE TYPE KOCEL.TMCELL_DATAS 
/* A table containing data describing the combined cells on a workbook sheet.*/
/* KOCEL.TYPES.sql*/
IS TABLE OF KOCEL.TMCELL_DATA
                   ');
END;
/
GRANT execute on KOCEL.TMCELL_DATAS to public;
/* ----------------------------------------------------------------------------
*/
CREATE or REPLACE TYPE KOCEL.TSHEET_PAR 
AS OBJECT (
/* An object containing the parameters of a workbook sheet.*/
/* KOCEL.TYPES.sql*/
/* The name of the parameter.*/
  PAR_NAME VARCHAR2(90),
/* Data type. "N" is a number, "D" is a date, "S" is a string.*/
  T CHAR(1),
/* A numeric value.*/
	N FLOAT(49),
/* The value of the date type.*/
  D DATE,
/* The value of the string type.*/
	S VARCHAR2(32000),
/* The sequential number of the sheet in the book.*/
  SHEET_NUM NUMBER(3),
/* The name of the sheet.*/
  SHEET VARCHAR2(255),
/* The name of the book.*/
  BOOK VARCHAR2(255),
  CONSTRUCTOR FUNCTION TSHEET_PAR return SELF AS RESULT
);
/
GRANT execute on KOCEL.TSHEET_PAR to public;

/* ----------------------------------------------------------------------------
*/
BEGIN
  EXECUTE IMMEDIATE('
CREATE OR REPLACE TYPE KOCEL.TSHEET_PARS 
/* A table containing the parameters of the workbook sheet.*/
/* KOCEL.TYPES.sql*/
IS TABLE OF KOCEL.TSHEET_PAR
                   ');
END;
/
GRANT execute on KOCEL.TSHEET_PARS to public;

/* ----------------------------------------------------------------------------
*/
CREATE or REPLACE TYPE BODY KOCEL.TMCELL_DATA 
AS
CONSTRUCTOR FUNCTION TMCELL_DATA return SELF AS RESULT
is
begin 
  self.L:=null;
  self.T:=null;
	self.R:=null;
  self.B:=null;
  self.SHEET_NUM:=null;
  self.SHEET:=null;
  self.BOOK:=null;
  return;
end;  
END;
/

/* ----------------------------------------------------------------------------
*/
CREATE or REPLACE TYPE BODY KOCEL.TSHEET_PAR 
AS
CONSTRUCTOR FUNCTION TSHEET_PAR return SELF AS RESULT
is
begin 
  self.PAR_NAME:=null;
  self.T:=null;
	self.N:=null;
  self.D:=null;
	self.S:=null;
  self.SHEET_NUM:=null;
  self.SHEET:=null;
  self.BOOK:=null;
  return;
end;  
END;
/

/* ----------------------------------------------------------------------------
*/
CREATE or REPLACE TYPE BODY KOCEL.TCELL_DATA 
AS
CONSTRUCTOR FUNCTION TCELL_DATA return SELF AS RESULT
is
begin 
  self.RID:=null;
  self.T:=null;
  self.R:=null;
	self.C:=null;
	self.N:=null;
  self.D:=null;
	self.S:=null;
	self.F:=null;
  self.SHEET_NUM:=null;
  self.SHEET:=null;
  self.BOOK:=null;
	self.Fmt:=null;
  return;
end;  
END;
/

/* ----------------------------------------------------------------------------
*/
CREATE or REPLACE TYPE BODY KOCEL.TCOLUMN_DATA 
AS
CONSTRUCTOR FUNCTION TCOLUMN_DATA return SELF AS RESULT
is
begin
  self.RID:=null;
	self.C:=null;
	self.W:=null;
  self.SHEET_NUM:=null;
  self.SHEET:=null;
  self.BOOK:=null;
	self.Fmt:=null;
  return;
end;  
END;
/

/* ----------------------------------------------------------------------------
*/
CREATE or REPLACE TYPE BODY KOCEL.TROW_DATA 
AS
CONSTRUCTOR FUNCTION TROW_DATA return SELF AS RESULT
is
begin
  self.RID:=null;
	self.R:=null;
	self.H:=null;
  self.SHEET_NUM:=null;
  self.SHEET:=null;
  self.BOOK:=null;
	self.Fmt:=null;
  return;
end;  
END;
/

/* ----------------------------------------------------------------------------
*/
CREATE or REPLACE TYPE KOCEL.TSHEET_NUM 
AS OBJECT (
/* An object containing the sequential number of the sheet in the book.*/
/* KOCEL.TYPES.sql*/
/* The serial number of the sheet.*/
  SHEET_NUM NUMBER(3),
/* The name of the sheet.*/
  SHEET VARCHAR2(255),
  CONSTRUCTOR FUNCTION TSHEET_NUM return SELF AS RESULT
);
/
GRANT execute on KOCEL.TSHEET_NUM to public;

/* ----------------------------------------------------------------------------
*/
BEGIN
  EXECUTE IMMEDIATE('
CREATE OR REPLACE TYPE KOCEL.TSHEET_NUMS
/* A table containing the order of the sheets in the book.*/
/* KOCEL.TYPES.sql*/
IS TABLE OF KOCEL.TSHEET_NUM
                   ');
END;
/
GRANT execute on KOCEL.TSHEET_NUMS to public;

/* ----------------------------------------------------------------------------
*/
CREATE or REPLACE TYPE BODY KOCEL.TSHEET_NUM
AS
CONSTRUCTOR FUNCTION TSHEET_NUM return SELF AS RESULT
is
begin
  self.SHEET_NUM:=null;
  self.SHEET:=null;
  return;
end;  
END;
/

/* ----------------------------------------------------------------------------
*/
BEGIN
  EXECUTE IMMEDIATE('
CREATE OR REPLACE TYPE KOCEL.TLSTRINGS 
/* Extended string is a string.*/
/* KOCEL-TYPES.tps*/
IS TABLE OF VARCHAR2(32000)
                   ');
END;

