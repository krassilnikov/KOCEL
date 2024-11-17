CREATE OR REPLACE PACKAGE KOCEL.IMPORT
/*  IMPORT package
 by Nikolay Krasilnikov
 This file is distributed under Apache License 2.0 (01.2004).
   http://www.apache.org/licenses/
 create 18.06.2013
 update 23.12.2023-24.12.2023
*/
AS

/*  The constant defines the maximum size of the pls block. If the size of the script
exceeds this size, the script will be split into several blocks.
*/
block_size constant NUMBER:=200;

/*  The procedure generates and executes a script for high-speed data loading
into the temporary table KOCEL.TEMP_SHEET.
 The incoming Array contains calls to the procedures described below.
*/
PROCEDURE LOAD_DATA( sheet in VARCHAR2, sheet_num in NUMBER, book in VARCHAR2,
                     sheet_data in KOCEL.TLSTRINGS);

/*  Loading a number
*/
PROCEDURE I_F(ur in NUMBER, uc in NUMBER, un in NUMBER,
              uf in VARCHAR2, uXF in NUMBER default 0);

/*  Loading the date
*/
PROCEDURE I_D(ur in NUMBER, uc in NUMBER, ud in DATE,
              uf in VARCHAR2, uXF in NUMBER default 0);

/*  Loading a line
*/
PROCEDURE I_S(ur in NUMBER, uc in NUMBER, us in VARCHAR2,
              uf in VARCHAR2, uXF in NUMBER default 0);

/*  Loading merged cells
*/
PROCEDURE I_M( 
/* --- The number of the first column of the union (cell number).
*/
   fCol IN NUMBER,
/* --- The number of the first row of the union (row number).
*/
   fRow IN NUMBER,
/* --- The number of the last column of the union (cell number).
*/
   lCol IN NUMBER,
/* --- The number of the last row of the union (row number).
*/
   lRow IN NUMBER);
    
end IMPORT;
/
