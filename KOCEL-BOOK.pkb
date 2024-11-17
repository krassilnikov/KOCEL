CREATE OR REPLACE PACKAGE BODY KOCEL.BOOK
-- KOCEL BOOK package body
-- create 08.02.2010
-- by Nikolay Krasilnikov
-- This file is distributed under Apache License 2.0 (01.2004).
--   http://www.apache.org/licenses/
-- update 15.02.2010 08.07.2015 26.09.2016 23.11.2023 24.11.2023 17.01.2024
--*****************************************************************************

as

CBook VARCHAR2(255);
SCount NUMBER(3);
MaxR NUMBER;

Fmts KOCEL.TNUMS;
type TFmtNum is table of NUMBER index by VARCHAR2(40);
FmtNum TFmtNum;
type TSheetNum is table of NUMBER index by VARCHAR2(255);
SheetNum TSheetNum;
type TSheetRec is record (SHEET_NUM NUMBER, SHEET VARCHAR2(255));
type TSheetRecs is table of TSheetRec index by PLS_INTEGER;
SheetRec TSheetRecs;

-------------------------------------------------------------------------------
procedure ClearData
is
begin
  FmtNum.delete;
  SheetNum.delete;
  SheetRec.delete;
  Fmts:=null;
  CBook:=null;
  SCount:=null;
end ClearData;

-------------------------------------------------------------------------------
procedure PrepareData(UBOOK in VARCHAR2, MAX_ROWS in NUMBER default 1000)
is
  tmpVar PLS_INTEGER;
begin
  ClearData;
  MaxR := MAX_ROWS;
  if UBook is null then return; end if;
  -- ������� ���������� ������ � �����, �� ������� ���������� ����� ������� �
  -- �����.
  tmpVar:=0;
  CBook:=UBook;
  for rec in (select NUM, SHEET from KOCEL.SHEETS_ORDER
                where BOOK=CBook
                order by NUM,SHEET)
  loop
    -- ������������� ������ ������ � ������� ����� ���������� ��������.
    -- �� ���������� � ���������� �������� ���������.
    tmpVar:=tmpVar+1;
    SheetRec(tmpVar).SHEET_NUM:=tmpVar;
    SheetRec(tmpVar).SHEET:=rec.SHEET;
    SheetNum(rec.SHEET):=tmpVar;
  end loop;              
  for rec in (select distinct SHEET from KOCEL.SHEETS s
	              where s.BOOK=CBook
		              and s.Sheet not in (select SHEET 
                                               from KOCEL.SHEETS_ORDER
	                                             where BOOK=CBook
                                             )  
              order by SHEET)
  loop
    tmpVar:=tmpVar+1;
    SheetRec(tmpVar).SHEET_NUM:=tmpVar;
    SheetRec(tmpVar).SHEET:=rec.SHEET;
    SheetNum(rec.SHEET):=tmpVar;
  end loop;
  if tmpVar=0 then
    raise_application_Error(-20443,	'Book '||UBook||' is missing!');
  end if;
  SCount:=tmpVar;
  -- ��������� ������ �������� ��������, �������������� � ����� �
  -- ������ ������������ ������� � ������� ��������. 
  FmtNum('0'):=0;
  -- ��������� ������ ������ ������� �������� �� �������.
  select distinct Fmt bulk collect into Fmts from KOCEL.SHEETS
    where Fmt is not null
      and Fmt > 0
      and BOOK=CBOOK;
  if Fmts.count=0 then return; end if;    
  for i in 1..Fmts.count    
  loop
    FmtNum(to_char(Fmts(i))):=i;
  end loop;
end PrepareData;

-------------------------------------------------------------------------------
function Pars return KOCEL.TSHEET_PARS pipelined
is
  orec KOCEL.TSHEET_PAR;
begin
  if CBook is null then return; end if;
  orec:=KOCEL.TSHEET_PAR;
  orec.BOOK:=CBOOK;
  for rec in(select PAR_NAME, T, N, D, S, SHEET, BOOK from KOCEL.V_SHEET_PARS
               where BOOK=CBOOK) 
  loop
    orec.PAR_NAME:=rec.PAR_NAME;
	  orec.T:=rec.T;
    orec.N:=rec.N;
	  orec.D:=rec.D;
	  orec.S:=rec.S;
    if rec.SHEET is null then
      orec.SHEET_NUM:=0;
    else  
      orec.SHEET_NUM:=SheetNum(rec.SHEET);
    end if;
    orec.SHEET:=rec.SHEET;
    pipe row(orec);
  end loop
  return;
exception
  when no_data_needed then 
    null;
  when others then
    d(SQLERRM, 'ERROR in KOCEL.BOOK.Pars');
    raise;
end Pars;

-------------------------------------------------------------------------------
function Formats return KOCEL.TFORMATS pipelined
is
  orec KOCEL.TFORMAT;
begin
  if CBook is null then return; end if;
  orec:=KOCEL.TFORMAT(0);
  orec.Tag:=Fmts.Count;
  pipe row(orec);
  if Fmts.count=0 then return; end if;   
  for i in 1..Fmts.count    
  loop
    orec:=KOCEL.TFORMAT(Fmts(i));
    orec.Tag:=i;
    pipe row(orec);
  end loop;
  return;
exception
  when no_data_needed then 
    null;
  when others then
    d(SQLERRM, 'ERROR in KOCEL.BOOK.Formats');
    raise;
end Formats;

-------------------------------------------------------------------------------
function Sheets return KOCEL.TSHEET_NUMS pipelined
is
  orec KOCEL.TSHEET_NUM;
begin
  if CBook is null then return; end if;
  orec:=KOCEL.TSHEET_NUM(SCount,null);
  pipe row(orec);
  for i in SheetRec.first..SheetRec.last
  loop
    orec.SHEET_NUM:=SheetRec(i).SHEET_NUM;
    orec.SHEET:=SheetRec(i).SHEET;
    pipe row(orec);
  end loop;
  return;
exception
  when no_data_needed then 
    null;
  when others then
    d(SQLERRM, 'ERROR in KOCEL.BOOK.Sheets');
    raise;
end Sheets;

-------------------------------------------------------------------------------
function Sheet_Row_s(USHEET in VARCHAR2) return KOCEL.TROW_DATAS pipelined
is
  orec KOCEL.TROW_DATA;
begin
  if CBook is null then return; end if;
  orec:=KOCEL.TROW_DATA;
  orec.SHEET_NUM:=SheetNum(upper(USHEET));
  orec.SHEET:=USHEET;
  orec.BOOK:=CBOOK;
  -- ���������� ��� ������, ������� ������������� ������� �� ��� �����.
  for rec in(select RID,R,H, nvl(Fmt,0) Fmt 
               from KOCEL.V_ROWS
               where BOOK=CBOOK
                 and R <= MaxR
                 and SHEET=USHEET) 
  loop
    orec.RID:=rec.RID;
    orec.R:=rec.R;
	  orec.H:=rec.H;
    begin
  	  orec.Fmt:=FmtNum(rec.Fmt);
    exception
      -- ���� ������ �� ����� �������, �� ����������� 0
      when no_data_found then orec.Fmt:=0;  
    end;  
    pipe row(orec);
  end loop;
  return;
exception
  when no_data_needed then 
    null;
  when others then
    d(SQLERRM, 'ERROR in KOCEL.BOOK.Sheet_Row_s');
    raise;
end Sheet_Row_s;

-------------------------------------------------------------------------------
function Row_s return KOCEL.TROW_DATAS pipelined
is
  orec KOCEL.TROW_DATA;
begin
  if CBook is null then return; end if;
  orec:=KOCEL.TROW_DATA;
  orec.BOOK:=CBOOK;
  -- ���������� ��� ������, ������� ������������� ������� �� ��� �����.
  for rec in(select RID, R, H, SHEET, nvl(Fmt,0) Fmt 
               from KOCEL.V_ROWS
               where BOOK=CBOOK
                 and R <= MaxR
            ) 
  loop
    orec.RID:=rec.RID;
    orec.R:=rec.R;
	  orec.H:=rec.H;
    orec.SHEET_NUM:=SheetNum(rec.SHEET);
    orec.SHEET:=rec.SHEET;
    begin
  	  orec.Fmt:=FmtNum(rec.Fmt);
    exception
      -- ���� ������ �� ����� �������, �� ����������� 0
      when no_data_found then orec.Fmt:=0;  
    end;  
    pipe row(orec);
  end loop
  return;
exception
  when no_data_needed then 
    null;
  when others then
    d(SQLERRM, 'ERROR in KOCEL.BOOK.Row_s');
    raise;
end Row_s;

-------------------------------------------------------------------------------
function Sheet_Column_s(USHEET in VARCHAR2) return KOCEL.TCOLUMN_DATAS
pipelined
is
  orec KOCEL.TCOLUMN_DATA;
begin
  if CBook is null then return; end if;
  orec:=KOCEL.TCOLUMN_DATA;
  orec.BOOK:=CBOOK;
  orec.SHEET_NUM:=SheetNum(USHEET);
  orec.SHEET:=USHEET;
  -- ���������� ��� ������, ������� ������������� ������� �� ��� �����.
  for rec in(select RID, C, W, SHEET, nvl(Fmt,0) Fmt 
               from KOCEL.V_COLUMNS
               where BOOK=CBOOK
                 and SHEET=USHEET) 
  loop
    orec.RID:=rec.RID;
    orec.C:=rec.C;
	  orec.W:=rec.W;
    begin
  	  orec.Fmt:=FmtNum(rec.Fmt);
    exception
      -- ���� ������ �� ����� �������, �� ����������� 0
      when no_data_found then orec.Fmt:=0;  
    end;  
    pipe row(orec);
  end loop
  return;
exception
  when no_data_needed then 
    null;
  when others then
    d(SQLERRM, 'ERROR in KOCEL.BOOK.Sheet_Column_s');
    raise;
end Sheet_Column_s;

-------------------------------------------------------------------------------
function Column_s return KOCEL.TCOLUMN_DATAS pipelined
is
  orec KOCEL.TCOLUMN_DATA;
begin
  if CBook is null then return; end if;
  orec:=KOCEL.TCOLUMN_DATA;
  orec.BOOK:=CBOOK;
  -- ���������� ��� ������, ������� ������������� ������� �� ��� �����.
  for rec in(select RID, C, W, SHEET, nvl(Fmt,0) Fmt 
               from KOCEL.V_COLUMNS
               where BOOK=CBOOK) 
  loop
    orec.RID:=rec.RID;
    orec.C:=rec.C;
	  orec.W:=rec.W;
    orec.SHEET_NUM:=SheetNum(rec.SHEET);
    orec.SHEET:=rec.SHEET;
    begin
  	  orec.Fmt:=FmtNum(rec.Fmt);
    exception
      -- ���� ������ �� ����� �������, �� ����������� 0
      when no_data_found then orec.Fmt:=0;  
    end;  
    pipe row(orec);
  end loop
  return;
exception
  when no_data_needed then 
    null;
  when others then
    d(SQLERRM, 'ERROR in KOCEL.BOOK.Column_s');
    raise;
end Column_s;

-------------------------------------------------------------------------------
function Sheet_Cells(USHEET in VARCHAR2) return KOCEL.TCELL_DATAS pipelined
is
  orec KOCEL.TCELL_DATA;
begin
  if CBook is null then return; end if;
  orec:=KOCEL.TCELL_DATA;
  orec.SHEET_NUM:=SheetNum(upper(USHEET));
  orec.SHEET:=USHEET;
  orec.BOOK:=CBOOK;
  -- ���������� ��� ������, ������� ������������� ������� �� ��� �����.
  for rec in
  (
--  select RID,T,R,C,N,D,S,F, nvl(Fmt,0) Fmt 
--               from KOCEL.V_SHEETS
--               where upper(BOOK)=upper(CBOOK)
--                 and upper(SHEET)=upper(USHEET)
  select rowidtochar(s.rowid) RID,
    case 
      when s.D is not null then 'D'
      when (s.D is null) and (s.N is not null) then 'N'
    else 'S'
    end T,
    s.R, s.C, s.N, s.D, nvl(s.LS,s.S) S, F,
    nvl(Fmt,0) Fmt 
  from KOCEL.SHEETS s
  where s.R between 1 and MaxR
    and s.C > 0          
    and BOOK = CBOOK
    and SHEET=USHEET
  ) 
  loop
    orec.RID:=rec.RID;
    orec.T:=rec.T;
    orec.R:=rec.R;
	  orec.C:=rec.C;
	  orec.N:=rec.N;
    orec.D:=rec.D;
	  orec.S:=rec.S;
	  orec.F:=rec.F;
    begin
  	  orec.Fmt:=FmtNum(rec.Fmt);
    exception
      -- ���� ������ �� ����� �������, �� ����������� 0
      when no_data_found then orec.Fmt:=0;  
    end;  
    pipe row(orec);
  end loop
  return;
exception
  when no_data_needed then 
    null;
  when others then
    d(SQLERRM, 'ERROR in KOCEL.BOOK.Sheet_Cells');
    raise;
end Sheet_Cells;

-------------------------------------------------------------------------------
function Cells return KOCEL.TCELL_DATAS pipelined
is
  orec KOCEL.TCELL_DATA;
begin
  if CBook is null then return; end if;
  orec:=KOCEL.TCELL_DATA;
  orec.BOOK:=CBOOK;
  -- ���������� ��� ������, ������� ������������� ������� �� ��� �����.
  for rec in
  (
--    select RID, T, R, C, N, D, S, F, SHEET, nvl(Fmt,0) Fmt 
--                 from KOCEL.V_SHEETS
--                 where upper(BOOK)=upper(CBOOK)
--               order by SHEET
  select rowidtochar(s.rowid) RID,
    case 
      when s.D is not null then 'D'
      when (s.D is null) and (s.N is not null) then 'N'
    else 'S'
	  end T,
    s.R, s.C, s.N, s.D, nvl(s.LS,s.S) S, F, SHEET, nvl(Fmt,0) Fmt 
  from KOCEL.SHEETS s
  where s.R between 1 and MaxR
    and s.C > 0          
    and BOOK = CBOOK
  ) 
  loop
    orec.RID:=rec.RID;
    orec.T:=rec.T;
    orec.R:=rec.R;
	  orec.C:=rec.C;
	  orec.N:=rec.N;
    orec.D:=rec.D;
	  orec.S:=rec.S;
	  orec.F:=rec.F;
    orec.SHEET_NUM:=SheetNum(rec.SHEET);
    orec.SHEET:=rec.SHEET;
    begin
  	  orec.Fmt:=FmtNum(rec.Fmt);
    exception
      -- ���� ������ �� ����� �������, �� ����������� 0
      when no_data_found then orec.Fmt:=0;  
    end;  
    pipe row(orec);
  end loop
  return;
exception
  when no_data_needed then 
    null;
  when others then
    d(SQLERRM, 'ERROR in KOCEL.BOOK.Cells');
    raise;
end Cells;

-------------------------------------------------------------------------------
function Sheet_MCells(USHEET in VARCHAR2) return KOCEL.TMCELL_DATAS pipelined
is
  orec KOCEL.TMCELL_DATA;
begin
  if CBook is null then return; end if;
  orec:=KOCEL.TMCELL_DATA;
  orec.SHEET_NUM:=SheetNum(USHEET);
  orec.SHEET:=USHEET;
  orec.BOOK:=CBOOK;
  -- ���������� ��� ������.
  for rec in(select L,T,R,B,SHEET,BOOK from KOCEL.MERGED_CELLS
               where BOOK=CBOOK
                 and SHEET=USHEET
                 and L between 1 and MaxR
                 and T between 1 and MaxR
                 and R between 1 and MaxR
                 and B between 1 and MaxR
            ) 
  loop
    orec.L:=rec.L;
	  orec.T:=rec.T;
    orec.R:=rec.R;
	  orec.B:=rec.B;
    pipe row(orec);
  end loop;
  return;
exception
  when no_data_needed then 
    null;
  when others then
    d(SQLERRM, 'ERROR in KOCEL.BOOK.Sheet_MCells');
    raise;
end Sheet_MCells;

-------------------------------------------------------------------------------
function MCells return KOCEL.TMCELL_DATAS pipelined
is
  orec KOCEL.TMCELL_DATA;
begin
  if CBook is null then return; end if;
  orec:=KOCEL.TMCELL_DATA;
  orec.BOOK:=CBOOK;
  -- ���������� ��� ������.
  for rec in(select L,T,R,B,SHEET,BOOK from KOCEL.MERGED_CELLS
               where BOOK=CBOOK
                 and L between 1 and MaxR
                 and T between 1 and MaxR
                 and R between 1 and MaxR
                 and B between 1 and MaxR
            ) 
  loop
    orec.L:=rec.L;
	  orec.T:=rec.T;
    orec.R:=rec.R;
	  orec.B:=rec.B;
    orec.SHEET_NUM:=SheetNum(rec.SHEET);
    orec.SHEET:=rec.SHEET;
    pipe row(orec);
  end loop
  return;
exception
  when no_data_needed then 
    null;
  when others then
    d(SQLERRM, 'ERROR in KOCEL.BOOK.MCells');
    raise;
end MCells;

END;
/