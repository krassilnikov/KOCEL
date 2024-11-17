CREATE OR REPLACE PACKAGE BODY KOCEL.IMPORT
as
-- IMPORT package body
-- by Nikolay Krasilnikov
-- This file is distributed under Apache License 2.0 (01.2004).
--   http://www.apache.org/licenses/
-- create 18.06.2013
-- update 23.12.2023-24.12.2023

USHEET VARCHAR2(255);
UNUM NUMBER;
UBOOK VARCHAR(255);

PROCEDURE LOAD_DATA(sheet in VARCHAR2, sheet_num in NUMBER, book in VARCHAR2,
                    sheet_data in KOCEL.TLSTRINGS)
is
  -- �����, �������������� ��� ����������.
  execScript DBMS_SQL.VARCHAR2A;
  -- ������������� �������. 
  c NUMBER;
  -- ��������� �� ������.
  EM SP.COMMANDS.COMMENTS%type;
  -- ����������� ������ �������.
  execLine PLS_INTEGER;
  -- ���������� ����� ������ ������ ����� �������.
  partLine PLS_INTEGER;
  -- ��������� ����������.
  tmpVar  NUMBER;
  k PLS_INTEGER;
  partNUM PLS_INTEGER;
  -- ������� ��������� ������, �������������� � ������� execScript.
  function exec_part return BOOLEAN
  is
  begin
    -- ��������� �������������� ������.
    partNUM := partNUM +1;
    if partNUM = 100 then
      partNUM := 0;  
      -- ��������� ������, ��������� ���������� ��������� �� ������� ����������
      -- ������ �������.
      execScript(execScript.last+1):=
        'd('''||to_char(execLine)||' OK! of '||to_char(sheet_data.count)
        ||''',''KOCEL.IMPORT'');';
	  end if;  
    -- ��������� �������� ����� END.  
	  execScript(execScript.last+1):=' END;';
	  -- ���������� � �������� ������������ ������.
--	  for i in 1..execScript.last
--	  loop
--      -- ���� ������� ������ ����� 3500 ��������, �� ��������� � �� ���.
--      if length(execScript(i)) > 3500 then
--	      d(to_char(partLine+i)||'_1 '|| substr(execScript(i),1,3500),
--          'KOCEL.IMPORT');
--	      d(to_char(partLine+i)||'_2 '|| substr(execScript(i),3501),
--          'KOCEL.IMPORT');
--      else
--	      d(to_char(partLine+i)||' '|| execScript(i),'KOCEL.IMPORT');
--      end if;  
--	  end loop;
    -- ����������� ������.
	  begin
	    dbms_sql.parse(c, execScript, 1, execScript.last, true, dbms_sql.native);
	  exception
		  when others then
		    if dbms_sql.is_open(c) then
		      dbms_sql.close_cursor(c);
		    end if;
	      EM:=SQLERRM;
	      d(EM,'ERROR parse KOCEL.IMPORT');
	      EM:='ERROR parse KOCEL.IMPORT  '||EM;
	      return false;
	  end;
	  -- ��������� ������.
	  begin
	   tmpVar:=dbms_sql.execute(c);
	  exception
		  when others then
		    if dbms_sql.is_open(c) then
		      dbms_sql.close_cursor(c);
		    end if;
	      EM:=SQLERRM;
	      d(EM,'ERROR execute KOCEL.IMPORT');
	      EM:='ERROR execute KOCEL.IMPORT  '||EM;
	      return false;
	  end;
    return true;
  end exec_part;
--
--  
begin
  USHEET := sheet;
  UNUM := sheet_num;
  UBOOK := book;
  partNUM := 0;
  if (sheet_data is null) or (sheet_data.first  is null) then
    d(' is empty?!!!','KOCEL.IMPORT');
    return;
  end if;
	d('START '||UBOOK||':'||USHEET,'KOCEL.IMPORT!');
  -- ��������� ������ ������.
  execScript(1):='BEGIN ';
  execLine:=1;
  partLine:=0;
  k:=1;
	c:=dbms_sql.open_cursor;
  -- ��������� ����� ������� ������ � block_size �����,
  -- ���������� ������ ����� ������ INPUT � ���� ������.
  for i in sheet_data.first..sheet_data.last
  loop
    -- �������� ����� ������ ��������� �������,
    -- ����  ������ �������� ������� ���������� � ������ ���������.
    if instr(sheet_data(i),'I_') = 1 then
	    -- ���� ����� ����� � ������� ����� block_size,
	    -- �� ��������� ����� �������.
	    if k >= block_size then
		    -- ���� �������� ������, �� ���������� ������.
		    if not exec_part then 
		      raise_application_Error(-20443,	EM);
		    end if;
        execScript.delete;
        k:=1;
		    -- ��������� BEGIN � ��������� ����� �������.
		    execScript(k):='BEGIN ';
        partLine:=execLine;
		    execLine:=execLine+1;
	    end if;  
--      if k > 1 then
--        execScript(execScript.last+1):=
--          'd('''||to_char(execLine)||' OK! '',''KOCEL.IMPORT'');';
--        k:=k+1;
--        execLine:=execLine+1;
--      end if;    
      k:=k+1;
      execLine:=execLine+1;
      execScript(k) := 'KIMP.' || sheet_data(i);
    else
      execScript(k):=execScript(k)||sheet_data(i);
    end if;
  end loop;
  -- ��������� ��������� ����� �������, .
  if not exec_part then 
    raise_application_Error(-20443,	EM);
  end if;
  dbms_sql.close_cursor(c);
  d('FINISH '||UBOOK||':'||USHEET,'KOCEL.IMPORT!');
  return;
exception
  when others then
    if dbms_sql.is_open(c) then
      dbms_sql.close_cursor(c);
    end if;
    EM:=SQLERRM || '  BACKTRACE '|| DBMS_UTILITY.FORMAT_ERROR_BACKTRACE();
    d(EM,'ERROR KOCEL.IMPORT');
    raise;
end LOAD_DATA;


PROCEDURE I_F(ur in NUMBER, uc in NUMBER, un in NUMBER,
              uf in VARCHAR2, uXF in NUMBER default 0)
is
begin
  insert into KOCEL.TEMP_SHEET
    (R, C, N, F, SHEET, NUM, BOOK, FMT_NUM)
    values(uR,uC,uN,uF,USHEET,UNUM,UBOOK,uXF);         
end;

PROCEDURE I_D(ur in NUMBER, uc in NUMBER, ud in DATE,
              uf in VARCHAR2, uXF in NUMBER default 0)
is
begin
  insert into KOCEL.TEMP_SHEET
    (R, C, D, F, SHEET, NUM, BOOK, FMT_NUM)
    values(uR,uC,uD,uF,USHEET,UNUM,UBOOK,uXF);         

end I_D;

PROCEDURE I_S(ur in NUMBER, uc in NUMBER, us in VARCHAR2,
              uf in VARCHAR2, uXF in NUMBER default 0)
is
begin
  insert into KOCEL.TEMP_SHEET
    (R, C, S, F, SHEET, NUM, BOOK, FMT_NUM)
    values(uR,uC,uS,uF,USHEET,UNUM,UBOOK,uXF);         

end I_S;

PROCEDURE I_M( fCol IN NUMBER, fRow IN NUMBER, lCol IN NUMBER, lRow IN NUMBER) 
is
begin
  insert into KOCEL.Merged_Cells
  values
    (fCol, fRow, lCol, lRow, USHEET, UBOOK);
exception
  when dup_val_on_index then
    raise_application_error(-20001,
                            '����������� �����: fCol="' || fCol ||
                            '"; fRow:"' || fRow ||
                            '" - ��� ���������� ��� �����: "' || UBOOK ||
                            '", � �����: "' || USHEET || '"');
end I_M;


end IMPORT;
/