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
  -- Текст, подготовленный для компиляции.
  execScript DBMS_SQL.VARCHAR2A;
  -- Идентификатор курсора. 
  c NUMBER;
  -- Сообщение об ошибке.
  EM SP.COMMANDS.COMMENTS%type;
  -- Исполняемая строка скрипта.
  execLine PLS_INTEGER;
  -- Порядковый номер первой строки части скрипта.
  partLine PLS_INTEGER;
  -- Временные переменные.
  tmpVar  NUMBER;
  k PLS_INTEGER;
  partNUM PLS_INTEGER;
  -- Функция выполняет скрипт, подготовленный в массиве execScript.
  function exec_part return BOOLEAN
  is
  begin
    -- Добавляем заключительные строки.
    partNUM := partNUM +1;
    if partNUM = 100 then
      partNUM := 0;  
      -- Добавляем строку, выводящую отладочное сообщение об удачном выполнении
      -- частей скрипта.
      execScript(execScript.last+1):=
        'd('''||to_char(execLine)||' OK! of '||to_char(sheet_data.count)
        ||''',''KOCEL.IMPORT'');';
	  end if;  
    -- Добавляем ключевое слово END.  
	  execScript(execScript.last+1):=' END;';
	  -- Записываем в отладчик получившийся скрипт.
--	  for i in 1..execScript.last
--	  loop
--      -- Если текущая строка более 3500 символов, то разбиваем её на две.
--      if length(execScript(i)) > 3500 then
--	      d(to_char(partLine+i)||'_1 '|| substr(execScript(i),1,3500),
--          'KOCEL.IMPORT');
--	      d(to_char(partLine+i)||'_2 '|| substr(execScript(i),3501),
--          'KOCEL.IMPORT');
--      else
--	      d(to_char(partLine+i)||' '|| execScript(i),'KOCEL.IMPORT');
--      end if;  
--	  end loop;
    -- Компилируем скрипт.
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
	  -- Выполняем скрипт.
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
  -- Добавляем первую строку.
  execScript(1):='BEGIN ';
  execLine:=1;
  partLine:=0;
  k:=1;
	c:=dbms_sql.open_cursor;
  -- Формируем части скрипта длиной в block_size строк,
  -- выстраивая каждый вызов пакета INPUT в одну строку.
  for i in sheet_data.first..sheet_data.last
  loop
    -- Начинаем новую строку выходного скрипта,
    -- если  строка входного скрипта начинается с вызова процедуры.
    if instr(sheet_data(i),'I_') = 1 then
	    -- Если число строк в скрипте более block_size,
	    -- то выполняем часть скрипта.
	    if k >= block_size then
		    -- Если возникла ошибка, то возбуждаем ошибку.
		    if not exec_part then 
		      raise_application_Error(-20443,	EM);
		    end if;
        execScript.delete;
        k:=1;
		    -- Добавляем BEGIN к следующей части скрипта.
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
  -- Выполняем последнюю часть скрипта, .
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
                            'Объединение ячеек: fCol="' || fCol ||
                            '"; fRow:"' || fRow ||
                            '" - уже существует для книги: "' || UBOOK ||
                            '", и листа: "' || USHEET || '"');
end I_M;


end IMPORT;
/