CREATE OR REPLACE PACKAGE KOCEL.CELL
/*  KOCEL  CELL package
 by Nikolai Krasilnikov
 create 17.04.2009
 update 18.05.2009 15.12.2009 28.01.2010 01.02.2010 13.02.2010 03.03.2010
        17.08.2011 11.11.2014 20.02.2020 07.10.2022 23.11.2022

*/

AS
/*  The flag of the KOCEL.SHEETS_ad trigger operation.
*/
ClearBOOKS BOOLEAN; 

type TCellValue is record(
R NUMBER(6),
C NUMBER(3),
N FLOAT(49),
D DATE,
S VARCHAR2(32000),
F VARCHAR2(4000)
);

type TColumn is Table of TCellValue index by BINARY_INTEGER;
type TRow is Table of TCellValue index by BINARY_INTEGER;
type TRange is Table of TColumn index by BINARY_INTEGER;

inBOOK VARCHAR2(255);
inSHEET VARCHAR2(255);
outBOOK VARCHAR2(255);
outSHEET VARCHAR2(255);

tmpRange TRange;

overwrite constant BOOLEAN:= true;

/*  Creating a new book and/or sheet.
 If the sheet already exists, then there is an error.
 A book exists if it contains at least one sheet.
 A book sheet exists if at least one cell is entered in the table.
*/
PROCEDURE New_SHEET(newSHEET in VARCHAR2,newBOOK in VARCHAR2);

/*  Creating a new sheet.
*/
PROCEDURE New_outSHEET;
/*  Create a new sheet and set it as the default receiver.
*/
PROCEDURE New_outSHEET(newSHEET in VARCHAR2);
PROCEDURE New_outSHEET(newSHEET in VARCHAR2,newBOOK in VARCHAR2);

/*  Checking that the sheet exists.
*/
FUNCTION is_inSHEET_exist(eSHEET in VARCHAR2)
return BOOLEAN;
FUNCTION is_inSHEET_exist return BOOLEAN;
FUNCTION is_outSHEET_exist return BOOLEAN;
FUNCTION is_outSHEET_exist(eSHEET in VARCHAR2)
return BOOLEAN;
FUNCTION isSHEET_exist(eSHEET in VARCHAR2,eBOOK in VARCHAR2)
return BOOLEAN;

/*  Getting the value from the current workbook and/or worksheet.
*/
FUNCTION Val(CellRow in NUMBER,CellColumn in NUMBER)
return TCellValue;

FUNCTION Val(anySHEET in VARCHAR2,
             CellRow in NUMBER,CellColumn in NUMBER)
return TCellValue;

/*  Getting the value from any book, sheet.
*/
FUNCTION Val(anySHEET in VARCHAR2,anyBook in VARCHAR2,
             CellRow in NUMBER,CellColumn in NUMBER)
return TCellValue;

FUNCTION Val(CellRow in NUMBER,CellColumn in VARCHAR2)
return TCellValue;

FUNCTION Val(anySHEET in VARCHAR2,
             CellRow in NUMBER,CellColumn in VARCHAR2)
return TCellValue;

/*  Getting the value from any book, sheet.
*/
FUNCTION Val(anySHEET in VARCHAR2,anyBook in VARCHAR2,
             CellRow in NUMBER,CellColumn in VARCHAR2)
return TCellValue;

/*  Checking that the value is a date.
*/
FUNCTION isDate(CellValue in TCellValue)
return BOOLEAN;

/*  Checking that the value is a number.
*/
FUNCTION isNumber(CellValue in TCellValue)
return BOOLEAN;

/*  Checking that the value is null.
*/
FUNCTION isNull(CellValue in TCellValue)
return BOOLEAN;

/*  Checking that the value is a string.
*/
FUNCTION isString(CellValue in TCellValue)
return BOOLEAN;

/*  Checking that the formula is present.
*/
FUNCTION hasFormula(CellValue in TCellValue)
return BOOLEAN;

/*  Getting the value as a date. If it is not a date, then null.
*/
FUNCTION asDate(CellRow in NUMBER,CellColumn in NUMBER)
return DATE;

FUNCTION asDate(anySHEET in VARCHAR2,
                CellRow in NUMBER,CellColumn in NUMBER)
return DATE;

FUNCTION asDate(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                CellRow in NUMBER,CellColumn in NUMBER)
return DATE;

FUNCTION asDate(CellRow in NUMBER,CellColumn in VARCHAR2)
return DATE;

FUNCTION asDate(anySHEET in VARCHAR2,
                CellRow in NUMBER,CellColumn in VARCHAR2)
return DATE;

FUNCTION asDate(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                CellRow in NUMBER,CellColumn in VARCHAR2)
return DATE;

/*  Getting the value as a number. If it's not a number, then it's null.
*/
FUNCTION asNumber(CellRow in NUMBER,CellColumn in NUMBER)
return NUMBER;

FUNCTION asNumber(anySHEET in VARCHAR2,
                  CellRow in NUMBER,CellColumn in NUMBER)
return NUMBER;

FUNCTION asNumber(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                  CellRow in NUMBER,CellColumn in NUMBER)
return NUMBER;

FUNCTION asNumber(CellRow in NUMBER,CellColumn in VARCHAR2)
return NUMBER;

FUNCTION asNumber(anySHEET in VARCHAR2,
                  CellRow in NUMBER,CellColumn in VARCHAR2)
return NUMBER;

FUNCTION asNumber(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                  CellRow in NUMBER,CellColumn in VARCHAR2)
return NUMBER;

/*  Getting the value as a string.
*/
FUNCTION asString(CellRow in NUMBER,CellColumn in NUMBER)
return VARCHAR2;

FUNCTION asString(anySHEET in VARCHAR2,
                  CellRow in NUMBER,CellColumn in NUMBER)
return VARCHAR2;

FUNCTION asString(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                  CellRow in NUMBER,CellColumn in NUMBER)
return VARCHAR2;

FUNCTION asString(CellRow in NUMBER,CellColumn in VARCHAR2)
return VARCHAR2;

FUNCTION asString(anySHEET in VARCHAR2,
                  CellRow in NUMBER,CellColumn in VARCHAR2)
return VARCHAR2;

FUNCTION asString(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                  CellRow in NUMBER,CellColumn in VARCHAR2)
return VARCHAR2;

/*  Getting the formula.
*/
FUNCTION Formula(CellRow in NUMBER,CellColumn in NUMBER)
return VARCHAR2;

FUNCTION Formula(anySHEET in VARCHAR2,
                 CellRow in NUMBER,CellColumn in NUMBER)
return VARCHAR2;

FUNCTION Formula(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                 CellRow in NUMBER,CellColumn in NUMBER)
return VARCHAR2;

FUNCTION Formula(CellRow in NUMBER,CellColumn in VARCHAR2)
return VARCHAR2;

FUNCTION Formula(anySHEET in VARCHAR2,
                 CellRow in NUMBER,CellColumn in VARCHAR2)
return VARCHAR2;

FUNCTION Formula(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                 CellRow in NUMBER,CellColumn in VARCHAR2)
return VARCHAR2;

/*  Getting a column.
 The result is a sparse array if the value is missing from the workbook
 it will not be among the array elements either.
 So the index null containing the column width is present only if the width
 The column has been modified by the user.
 In general, to organize loops, you need to use array attributes:
first,next,last.
 The index of the array corresponds to the number of the row in the book. 
*/
FUNCTION GetColumn(StartRow in NUMBER,StartColumn in NUMBER,
                   EndRow in NUMBER)
return TColumn;

FUNCTION GetColumn(entColumn in NUMBER)
return TColumn;

FUNCTION GetColumn(anySHEET in VARCHAR2,
                   StartRow in NUMBER,StartColumn in NUMBER,
                   EndRow in NUMBER)
return TColumn;

FUNCTION GetColumn(anySHEET in VARCHAR2,
                   entColumn in NUMBER)
return TColumn;

FUNCTION GetColumn(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                   StartRow in NUMBER,StartColumn in NUMBER,
                   EndRow in NUMBER)
return TColumn;

FUNCTION GetColumn(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                   entColumn in NUMBER)
return TColumn;

FUNCTION GetColumn(StartRow in NUMBER,StartColumn in VARCHAR2,
                   EndRow in NUMBER)
return TColumn;

FUNCTION GetColumn(entColumn in VARCHAR2)
return TColumn;

FUNCTION GetColumn(anySHEET in VARCHAR2,
                   StartRow in NUMBER,StartColumn in VARCHAR2,
                   EndRow in NUMBER)
return TColumn;

FUNCTION GetColumn(anySHEET in VARCHAR2,
                   entColumn in VARCHAR2)
return TColumn;

FUNCTION GetColumn(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                   StartRow in NUMBER,StartColumn in VARCHAR2,
                   EndRow in NUMBER)
return TColumn;

FUNCTION GetColumn(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                   entColumn in VARCHAR2)
return TColumn;

/*  Getting a row.
 The result is a sparse array if the value is missing from the workbook
 it will not be among the array elements either.
 So an element with index null containing the height of the row will be present
only if its height has been changed by the user.
 In general, to organize loops, you need to use array attributes:
first,next,last.
 The index of the array corresponds to the column number in the book.
*/
FUNCTION GetRow(StartRow in NUMBER,StartColumn in NUMBER,
                EndColumn in NUMBER)
return TRow;

FUNCTION GetRow(entRow in NUMBER)
return TRow;

FUNCTION GetRow(anySHEET in VARCHAR2,
                StartRow in NUMBER,StartColumn in NUMBER,
                EndColumn in NUMBER)
return TRow;

FUNCTION GetRow(anySHEET in VARCHAR2,
                entRow in NUMBER)
return TRow;

FUNCTION GetRow(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                StartRow in NUMBER,StartColumn in NUMBER,
                EndColumn in NUMBER)
return TRow;

FUNCTION GetRow(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                entRow in NUMBER)
return TRow;

FUNCTION GetRow(StartRow in NUMBER,StartColumn in VARCHAR2,
                EndColumn in VARCHAR2)
return TRow;

FUNCTION GetRow(anySHEET in VARCHAR2,
                StartRow in NUMBER,StartColumn in VARCHAR2,
                EndColumn in VARCHAR2)
return TRow;

FUNCTION GetRow(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                StartRow in NUMBER,StartColumn in VARCHAR2,
                EndColumn in VARCHAR2)
return TRow;

/*  Getting the next existing row. 
 The array of rows in the sheet is also discharged.
 The result is a sparse array if the value is missing from the workbook
 it will not be among the array elements either.
 So an element with index null containing the height of the row will be present
only if its height has been changed by the user.
 In general, to organize loops, you need to use array attributes:
first,next,last.
 The index of the array corresponds to the column number in the book.
*/
FUNCTION GetNextRow(StartRow in NUMBER,StartColumn in NUMBER,
                                       EndColumn in NUMBER)
return TRow;

FUNCTION GetNextRow(entRow in NUMBER)
return TRow;

FUNCTION GetNextRow(anySHEET in VARCHAR2,
                    StartRow in NUMBER,StartColumn in NUMBER,
                                       EndColumn in NUMBER)
return TRow;

FUNCTION GetNextRow(anySHEET in VARCHAR2,
                    entRow in NUMBER)
return TRow;

FUNCTION GetNextRow(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                    StartRow in NUMBER,StartColumn in NUMBER,
                                       EndColumn in NUMBER)
return TRow;

FUNCTION GetNextRow(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                    entRow in NUMBER)
return TRow;

FUNCTION GetNextRow(StartRow in NUMBER,StartColumn in VARCHAR2,
                                       EndColumn in VARCHAR2)
return TRow;

FUNCTION GetNextRow(anySHEET in VARCHAR2,
                    StartRow in NUMBER,StartColumn in VARCHAR2,
                                       EndColumn in VARCHAR2)
return TRow;

FUNCTION GetNextRow(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                    StartRow in NUMBER,StartColumn in VARCHAR2,
                                       EndColumn in VARCHAR2)
return TRow;

/*  Getting the range.
 The result is a sparse array if the value is missing from the workbook
 it will not be among the array elements either.
 In general, to organize loops, you need to use array attributes:
first,next,last.
 The indexes of the array correspond to: the first is the number of the row, the second is the number of the column.
*/
FUNCTION GetRange(StartRow in NUMBER,StartColumn in NUMBER,
                  EndRow in NUMBER,EndColumn in NUMBER)
return TRange;

FUNCTION GetRange(anySHEET in VARCHAR2,
                  StartRow in NUMBER,StartColumn in NUMBER,
                  EndRow in NUMBER,EndColumn in NUMBER)
return TRange;

FUNCTION GetRange(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                  StartRow in NUMBER,StartColumn in NUMBER,
                  EndRow in NUMBER,EndColumn in NUMBER)
return TRange;

FUNCTION GetRange(StartRow in NUMBER,StartColumn in VARCHAR2,
                  EndRow in NUMBER,EndColumn in VARCHAR2)
return TRange;

FUNCTION GetRange(anySHEET in VARCHAR2,
                  StartRow in NUMBER,StartColumn in VARCHAR2,
                  EndRow in NUMBER,EndColumn in VARCHAR2)
return TRange;

FUNCTION GetRange(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                  StartRow in NUMBER,StartColumn in VARCHAR2,
                  EndRow in NUMBER,EndColumn in VARCHAR2)
return TRange;


/*  Saving the value.
 The value is always stored according to the fields R,C.
 If the fields (N,D,S) are ambiguously defined, then an error occurs.
*/
PROCEDURE SetVal(val in TCellValue);

PROCEDURE SetVal(anySHEET in VARCHAR2,val in TCellValue);

PROCEDURE SetVal(anySHEET in VARCHAR2,anyBook in VARCHAR2, val in TCellValue);

/*  Saving the value if the argument is KOCEL.TVALUE.
*/
PROCEDURE SetVal(val in KOCEL.TValue);

PROCEDURE SetVal(anySHEET in VARCHAR2,val in KOCEL.TValue);

PROCEDURE SetVal(anySHEET in VARCHAR2,anyBook in VARCHAR2,val in KOCEL.TValue);


/*  Saving the value as a date.
*/
PROCEDURE SetValasDate(CellRow in NUMBER,CellColumn in NUMBER,
                       val in DATE);

PROCEDURE SetValasDate(anySHEET in VARCHAR2,
                       CellRow in NUMBER,CellColumn in NUMBER,
											 val in DATE);

PROCEDURE SetValasDate(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                       CellRow in NUMBER,CellColumn in NUMBER,
											 val in DATE);

PROCEDURE SetValasDate(CellRow in NUMBER,CellColumn in VARCHAR2,
                       val in DATE);

PROCEDURE SetValasDate(anySHEET in VARCHAR2,
                       CellRow in NUMBER,CellColumn in VARCHAR2,
											 val in DATE);

PROCEDURE SetValasDate(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                       CellRow in NUMBER,CellColumn in VARCHAR2,
											 val in DATE);

/*  Saving the value as a number.
*/
PROCEDURE SetValasNUMBER(CellRow in NUMBER,CellColumn in NUMBER,
                         val in NUMBER);

PROCEDURE SetValasNUMBER(anySHEET in VARCHAR2,
                         CellRow in NUMBER,CellColumn in NUMBER,
											   val in NUMBER);

PROCEDURE SetValasNUMBER(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                         CellRow in NUMBER,CellColumn in NUMBER,
											   val in NUMBER);

PROCEDURE SetValasNUMBER(CellRow in NUMBER,CellColumn in VARCHAR2,
                         val in NUMBER);

PROCEDURE SetValasNUMBER(anySHEET in VARCHAR2,
                         CellRow in NUMBER,CellColumn in VARCHAR2,
											   val in NUMBER);

PROCEDURE SetValasNUMBER(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                         CellRow in NUMBER,CellColumn in VARCHAR2,
											   val in NUMBER);

/*  Saving the value as a string.
*/
PROCEDURE SetValasString(CellRow in NUMBER,CellColumn in NUMBER,
                         val in VARCHAR2);

PROCEDURE SetValasString(anySHEET in VARCHAR2,
                         CellRow in NUMBER,CellColumn in NUMBER,
											   val in VARCHAR2);

PROCEDURE SetValasString(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                         CellRow in NUMBER,CellColumn in NUMBER,
											   val in VARCHAR2);

PROCEDURE SetValasString(CellRow in NUMBER,CellColumn in VARCHAR2,
                         val in VARCHAR2);

PROCEDURE SetValasString(anySHEET in VARCHAR2,
                         CellRow in NUMBER,CellColumn in VARCHAR2,
											   val in VARCHAR2);

PROCEDURE SetValasString(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                         CellRow in NUMBER,CellColumn in VARCHAR2,
											   val in VARCHAR2);

/*  Saving the formula.
 The value does not change.
*/
PROCEDURE SetFormula(CellRow in NUMBER,CellColumn in NUMBER,
                         val in VARCHAR2);

PROCEDURE SetFormula(anySHEET in VARCHAR2,
                         CellRow in NUMBER,CellColumn in NUMBER,
											   val in VARCHAR2);

PROCEDURE SetFormula(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                         CellRow in NUMBER,CellColumn in NUMBER,
											   val in VARCHAR2);

PROCEDURE SetFormula(CellRow in NUMBER,CellColumn in VARCHAR2,
                         val in VARCHAR2);

PROCEDURE SetFormula(anySHEET in VARCHAR2,
                         CellRow in NUMBER,CellColumn in VARCHAR2,
											   val in VARCHAR2);

PROCEDURE SetFormula(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                         CellRow in NUMBER,CellColumn in VARCHAR2,
											   val in VARCHAR2);

/*  Saving the column.
 Saving takes place according to the indexes of the array,
 and not the (R,C) value fields.
 Initially, the indexes match the value of the fields, but you should take into account
 that the correspondence is violated when the order of the elements is changed.
 To save according to the values of the field indexes, you need to save the values
using the procedure (setVal) in a loop through the array.
*/
PROCEDURE SetColumn(StartRow in NUMBER,StartColumn in NUMBER,
                    ColumnVals in TColumn);

PROCEDURE SetColumn(asColumn in NUMBER,ColumnVals in TColumn);

PROCEDURE SetColumn(anySHEET in VARCHAR2,
                    StartRow in NUMBER,StartColumn in NUMBER,
                    ColumnVals in TColumn);

PROCEDURE SetColumn(anySHEET in VARCHAR2,
                    asColumn in NUMBER,ColumnVals in TColumn);

PROCEDURE SetColumn(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                   StartRow in NUMBER,StartColumn in NUMBER,
                   ColumnVals in TColumn);

PROCEDURE SetColumn(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                    asColumn in NUMBER,ColumnVals in TColumn);

PROCEDURE SetColumn(StartRow in NUMBER,StartColumn in VARCHAR2,
                    ColumnVals in TColumn);

PROCEDURE SetColumn(asColumn in VARCHAR2,ColumnVals in TColumn);

PROCEDURE SetColumn(anySHEET in VARCHAR2,
                   StartRow in NUMBER,StartColumn in VARCHAR2,
                   ColumnVals in TColumn);

PROCEDURE SetColumn(anySHEET in VARCHAR2,
                    asColumn in VARCHAR2,ColumnVals in TColumn);

PROCEDURE SetColumn(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                   StartRow in NUMBER,StartColumn in VARCHAR2,
                   ColumnVals in TColumn);

PROCEDURE SetColumn(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                    asColumn in VARCHAR2,ColumnVals in TColumn);

/*  Saving a row.
 The saving takes place according to the indexes of the array,
 and not the (R,C) value fields.
 Initially, the indexes match the value of the fields, but you should take into account
 that the correspondence is violated when the order of the elements is changed.
 To save according to the values of the field indexes, you need to save the values
using the procedure (setVal) in a loop through the array.
*/
PROCEDURE SetRow(StartRow in NUMBER,StartColumn in NUMBER,
                 RowVals in TRow);

PROCEDURE SetRow(asRow in NUMBER,RowVals in TRow);

PROCEDURE SetRow(anySHEET in VARCHAR2,
                 StartRow in NUMBER,StartColumn in NUMBER,
                 RowVals in TRow);

PROCEDURE SetRow(anySHEET in VARCHAR2,
                 asRow in NUMBER,RowVals in TRow);

PROCEDURE SetRow(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                 StartRow in NUMBER,StartColumn in NUMBER,
                 RowVals in TRow);

PROCEDURE SetRow(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                 asRow in NUMBER,RowVals in TRow);

PROCEDURE SetRow(StartRow in NUMBER,StartColumn in VARCHAR2,
                 RowVals in TRow);

PROCEDURE SetRow(anySHEET in VARCHAR2,
                 StartRow in NUMBER,StartColumn in VARCHAR2,
                 RowVals in TRow);

PROCEDURE SetRow(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                 StartRow in NUMBER,StartColumn in VARCHAR2,
                 RowVals in TRow);

/*  Saving the range.
 The saving takes place according to the indexes of the array,
 and not the (R,C) value fields.
 Initially, the indexes match the value of the fields, but you should take into account
 that the correspondence is violated when the order of the elements is changed.
 To save according to the values of the field indexes, you need to save the values
using the procedure (setVal) in a loop through the array.
*/
PROCEDURE SetRange(StartRow in NUMBER,StartColumn in NUMBER,
									RangeVals in TRange);

PROCEDURE SetRange(anySHEET in VARCHAR2,
                  StartRow in NUMBER,StartColumn in NUMBER,
									RangeVals in TRange);

PROCEDURE SetRange(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                  StartRow in NUMBER,StartColumn in NUMBER,
									RangeVals in TRange);

PROCEDURE SetRange(StartRow in NUMBER,StartColumn in VARCHAR2,
									RangeVals in TRange);

PROCEDURE SetRange(anySHEET in VARCHAR2,
                  StartRow in NUMBER,StartColumn in VARCHAR2,
									RangeVals in TRange);

PROCEDURE SetRange(anySHEET in VARCHAR2,anyBook in VARCHAR2,
                  StartRow in NUMBER,StartColumn in VARCHAR2,
									RangeVals in TRange);

/*  Copying a sheet.
*/
PROCEDURE CopySHEET(fromSHEET in VARCHAR2,fromBook in VARCHAR2,
                    toSHEET in VARCHAR2,toBook in VARCHAR2,
										doOverwrire in BOOLEAN default false);

/*  Copying a book.
*/
PROCEDURE CopyBook(fromBook in VARCHAR2,toBook in VARCHAR2,
                   doOverwrire in BOOLEAN default false);
                   
/*  Renaming the book.
*/
PROCEDURE RenameBook(fromBook in VARCHAR2,toBook in VARCHAR2);

/*  Deleting a book.
*/
PROCEDURE DeleteBook(DBook in VARCHAR2);

/*  Deleting a sheet.
*/
PROCEDURE DeleteSheet(DBook in VARCHAR2, DSheet in VARCHAR2);

/*  Clearing the data of all the sheets of the book.
 The order of the sheets and the formatting of the columns remains.
 The merged cells are being disconnected.
*/
PROCEDURE ClearBook(DBook in VARCHAR2);
PROCEDURE ClearOutBook;

/*  Clearing the data of the workbook sheet.
 The order of the sheets and the formatting of the columns remains.
 The merged cells are being disconnected.
*/
PROCEDURE ClearSheet(DBook in VARCHAR2, DSheet in VARCHAR2);
PROCEDURE ClearOutSheet(DSheet in VARCHAR2);
PROCEDURE ClearOutSheet;

/*  Providing a column or row consisting of dates for the query.
 Missing values are ignored.
 Only values with a date that is not null are pre-assigned.
*/
FUNCTION Dates(RowVals in TRow)
return KOCEL.TDATES pipelined;

FUNCTION Dates(ColumnVals in TColumn)
return KOCEL.TDATES pipelined;

/*  Providing a column or row consisting of numbers for the query.
 Missing values are ignored.
 Only values with a number that is not null are excluded.
*/
FUNCTION Nums(RowVals in TRow)
return KOCEL.TNUMS pipelined;

FUNCTION Nums(ColumnVals in TColumn)
return KOCEL.TNUMS pipelined;

/*  Providing a column or row consisting of rows for the query.
 Missing values are ignored.
 Only values are changed if the string is not empty.
*/
FUNCTION Strings(RowVals in TRow)
return KOCEL.TSTRINGS pipelined;

FUNCTION Strings(ColumnVals in TColumn)
return KOCEL.TSTRINGS pipelined;

/*  Converting a column letter to a number.
*/
FUNCTION Col2Num(columnChar in VARCHAR2) return NUMBER;

/*  Converting a column number to a string.
*/
FUNCTION Col2Char(columnNum in NUMBER) return VARCHAR2;


/*  Providing a range for the request.
 Column numbers should be from 1 to 26 (A..Z).
Rows missing in the range are ignored.
*/
FUNCTION RANGE(OutRange in TRange)
return KOCEL.TCELLS pipelined;

/*  Providing a sheet for the request.
 Column numbers should be from 1 to 26 (A..Z).
Rows missing from the sheet are ignored.
*/
FUNCTION Vals(anySHEET in VARCHAR2,anyBOOK in VARCHAR2)
return KOCEL.TCELLS pipelined;

/*  Providing an incoming sheet for the request.
 Column numbers should be from 1 to 26 (A..Z).
Rows missing from the sheet are ignored.
*/
FUNCTION inVals
return KOCEL.TCELLS pipelined;

/*  Providing a worksheet for the query in the form of values converted to strings.
 Column numbers should be from 1 to 26 (A..Z).
Rows missing from the sheet are ignored.
*/
FUNCTION Vals_as_Str(anySHEET in VARCHAR2,anyBOOK in VARCHAR2)
return KOCEL.TSCELLS pipelined;

end CELL;
/
