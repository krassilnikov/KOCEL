/*  KOCEL tables triggers 
 create 17.04.2009
 by Nikolay Krasilnikov
 This file is distributed under Apache License 2.0 (01.2004).
   http://www.apache.org/licenses/
 update 27.04.2009 01.02.2010 16.02.2010 02.03.2020 23.11.2022 18.07.2024
*****************************************************************************
*/

/*  Books.
BEFORE_INSERT_TABLE------------------------------------------------------------

BEFORE_INSERT------------------------------------------------------------------
*/
CREATE OR REPLACE TRIGGER KOCEL.SHEETS_BIR
BEFORE INSERT ON KOCEL.SHEETS
FOR EACH ROW
-- (KOCEL_TRIGGERS.trg)
BEGIN
  -- Если поле даты не нулл, то строку и число делаем нулами.
	if :NEW.D is not null then
	  :NEW.N:=null;
		:NEW.S:=null;
		:NEW.LS:=null;
	else 
	-- Если поле даты нулл, то если число не нулл, то строка нулл.
	  if :NEW.N is not null then 
		  :NEW.S:=null;
		  :NEW.LS:=null;
		end if;		
	end if;
	if :NEW.LS is not null then 
		:NEW.S:=null;
	  :NEW.N:=null;
		:NEW.D:=null;
	end if;		
	if :NEW.S is not null then 
		:NEW.LS:=null;
	  :NEW.N:=null;
		:NEW.D:=null;
	end if;		
--	  raise_application_Error(-20443,'!!!! ');
  if :NEW.Fmt is null then :NEW.Fmt:=0; end if;
END;
/
/* AFTER_INSERT----------------------------------------------------------------

AFTER_INSERT_TABLE-------------------------------------------------------------

BEFORE_UPDATE_TABLE------------------------------------------------------------

BEFORE_UPDATE------------------------------------------------------------------
*/
CREATE OR REPLACE TRIGGER KOCEL.SHEETS_BUR
BEFORE UPDATE ON KOCEL.SHEETS
FOR EACH ROW
-- (KOCEL_TRIGGERS.trg)
BEGIN
  -- Ряд, колонку, книгу и лист изменять нельзя - можно удалить и вставить.
	:NEW.R:=:OLD.R;
	:NEW.C:=:OLD.C;
--	:NEW.BOOK:=:OLD.BOOK;
--	:NEW.SHEET:=:OLD.SHEET;
--	  raise_application_Error(-20443,'!!!! ');
  if :NEW.Fmt is null then :NEW.Fmt:=0; end if;
  -- Если поле даты было нулл, а стало не нулл,
	-- то строку и число делаем нулами.
	if (:NEW.D is not null) and (:OLD.D is NULL)  then
	  :NEW.N:=null;
		:NEW.S:=null;
    :NEW.LS:=null;
    return;
	end if;
  -- Если поле числа было нулл, а стало не нулл,
	-- то строку и дату делаем нулами.
	if (:NEW.N is not null) and (:OLD.N is NULL)  then
	  :NEW.D:=null;
		:NEW.S:=null;
    :NEW.LS:=null;
		return;
	end if;
  -- Если поле строки было нулл, а стало не нулл,
	-- то дату и число делаем нулами.
	if (:NEW.S is not null) and (:OLD.S is NULL)  then
	  :NEW.N:=null;
		:NEW.D:=null;
    :NEW.LS:=null;
		return;
	end if;
  -- Если поле строки было нулл, а стало не нулл,
	-- то дату и число делаем нулами.
	if (:NEW.LS is not null) and (:OLD.LS is NULL)  then
	  :NEW.N:=null;
		:NEW.D:=null;
    :NEW.S:=null;
		return;
	end if;
END;
/
/* 
AFTER_UPDATE------------------------------------------------------------
*/
CREATE OR REPLACE TRIGGER KOCEL.SHEETS_AUR
AFTER UPDATE ON KOCEL.SHEETS
FOR EACH ROW
-- (KOCEL_TRIGGERS.trg)
BEGIN
  if :OLD.R=0 and :OLD.C=0 then return; end if;
  -- Если изменён формат колонки, то меняем формат у всех ячеек колонки.
  if :NEW.Fmt != :OLD.Fmt and :OLD.R=0 then
    update KOCEL.SHEETS set Fmt=:NEW.Fmt
      where R>0
        and C=:OLD.C
        and upper(SHEET)=upper(:NEW.SHEET)
        and upper(BOOK)=upper(:NEW.BOOK);
    return;    
  end if;
  -- Если изменён формат ряда, то меняем формат у всех ячеек ряда.
  if :NEW.Fmt != :OLD.Fmt and :OLD.R=0 then
    update KOCEL.SHEETS set Fmt=:NEW.Fmt
      where C>0
        and R=:OLD.R
        and upper(SHEET)=upper(:NEW.SHEET)
        and upper(BOOK)=upper(:NEW.BOOK);
  end if;
END;
/
/* AFTER_UPDATE_TABLE----------------------------------------------------------

BEFORE_DELETE_TABLE------------------------------------------------------------

AFTER_DELETE_TABLE-------------------------------------------------------------

*****************************************************************************
*/

/*  Formats.
BEFORE_INSERT_TABLE------------------------------------------------------------

BEFORE_INSERT------------------------------------------------------------------
*/
CREATE OR REPLACE TRIGGER KOCEL.FORMATS_bir
BEFORE INSERT ON KOCEL.FORMATS
FOR EACH ROW
/*  (KOCEL_TRIGGERS.trg)
*/
BEGIN
  SELECT KOCEL.FORMAT_SEQ.NEXTVAL INTO :NEW.Fmt FROM DUAL;
  IF :NEW.FORMAT_STRING is null THEN :NEW.FORMAT_STRING:='GENERAL';END IF;
END;
/
/* AFTER_INSERT----------------------------------------------------------------

AFTER_INSERT_TABLE-------------------------------------------------------------

BEFORE_UPDATE_TABLE------------------------------------------------------------

BEFORE_UPDATE------------------------------------------------------------------
*/
CREATE OR REPLACE TRIGGER KOCEL.FORMATS_bur
BEFORE UPDATE ON KOCEL.FORMATS
FOR EACH ROW
/*  (KOCEL_TRIGGERS.trg)
*/
BEGIN
  raise_application_Error(-20443,
    'KOCEL.FORMATS_bur. Editing is prohibited! You can delete and paste it.');
END;
/
/* AFTER_UPDATE----------------------------------------------------------------

AFTER_UPDATE_TABLE-------------------------------------------------------------

BEFORE_DELETE_TABLE------------------------------------------------------------
*/
CREATE OR REPLACE TRIGGER KOCEL.FORMATS_bdr
BEFORE DELETE ON KOCEL.FORMATS
FOR EACH ROW
/*  (KOCEL_TRIGGERS.trg)
*/
BEGIN
  if :OLD.Fmt<100 then
    raise_application_Error(-20443,
      'KOCEL.FORMATS_bur. It is forbidden to delete embedded formats!');
  end if;
END;
/
/* AFTER_DELETE_TABLE-----------------------------------------------------------

*****************************************************************************
*/
/*  R1C1.
BEFORE_INSERT_TABLE------------------------------------------------------------

BEFORE_INSERT------------------------------------------------------------------
*/
CREATE OR REPLACE TRIGGER KOCEL.BOOKS_R1C1_BIR
BEFORE INSERT ON KOCEL.BOOKS_R1C1
FOR EACH ROW
-- (KOCEL_TRIGGERS.trg)
BEGIN
	if :NEW.LOAD_DATE is null then
	  :NEW.LOAD_DATE := sysdate;
	end if;
END;
/
/* AFTER_INSERT----------------------------------------------------------------
AFTER_INSERT_TABLE-------------------------------------------------------------
BEFORE_UPDATE_TABLE------------------------------------------------------------
BEFORE_UPDATE------------------------------------------------------------------
AFTER_UPDATE-------------------------------------------------------------------
AFTER_UPDATE_TABLE-------------------------------------------------------------
BEFORE_DELETE_TABLE------------------------------------------------------------
*/

/*Views
******************************************************************************/

/*  Representation of the combined cells.
INSTEAD_OF_INSERT--------------------------------------------------------------
*/
CREATE OR REPLACE TRIGGER KOCEL.V_MERGED_CELLS_ii
INSTEAD OF INSERT ON KOCEL.V_MERGED_CELLS
/*  (KOCEL_TRIGGERS.trg)
*/
DECLARE
  type TMS is record(RID ROWID, L NUMBER,T NUMBER,R NUMBER,B NUMBER);
  type TMSS is TABLE of TMS;
  MSS TMSS;
  i PLS_INTEGER;
BEGIN
  -- Проверяем, что у добавляемого диапазона ячеек нет пересечения с уже,
  -- существующими диапазономи.
  -- Если пересечения нет, то добавляем новый диапазон.
  -- Если пересечение существует, то измеяем последний из уже сущестующих,
  -- объединяя в нём все ячейки и удаляем остальные.
  select ROWID RID, L,T,R,B bulk collect into MSS from KOCEL.MERGED_CELLS
    where upper(SHEET)=upper(:NEW.SHEET)
      and upper(BOOK)=upper(:NEW.BOOK)
      and L between :NEW.L and :NEW.R 
      and T between :NEW.T and :NEW.B 
      and R between :NEW.L and :NEW.R
      and B between :NEW.T and :NEW.B;
  if MSS.count=0 then
    insert into KOCEL.MERGED_CELLS values (
      :NEW.L, :NEW.T, :NEW.R, :NEW.B, :NEW.SHEET, :NEW.BOOK);
    return;    
  end if;
  i:=MSS.last;      
  for j in MSS.first..i   
  loop
    MSS(i).L:=least(MSS(i).L, MSS(j).L);
    MSS(i).T:=least(MSS(i).T, MSS(j).T);
    MSS(i).R:=greatest(MSS(i).R,MSS(j).R);
    MSS(i).L:=greatest(MSS(i).B,MSS(j).B);
    if i!=j then
      delete from KOCEL.MERGED_CELLS where ROWID=MSS(j).RID;
    end if;  
  end loop;    
    MSS(i).L:=least(MSS(i).L,:NEW.L);
    MSS(i).T:=least(MSS(i).T,:NEW.T);
    MSS(i).R:=greatest(MSS(i).R,:NEW.R);
    MSS(i).L:=greatest(MSS(i).B,:NEW.B);
  update KOCEL.MERGED_CELLS set 
    L=MSS(i).L,
    T=MSS(i).T,
    R=MSS(i).R,
    B=MSS(i).B
    where ROWID=MSS(i).RID;
END;
/
/* 
INSTEAD_OF_UPDATE--------------------------------------------------------------
*/
CREATE OR REPLACE TRIGGER KOCEL.V_MERGED_CELLS_iu
INSTEAD OF UPDATE ON KOCEL.V_MERGED_CELLS
/*  (KOCEL_TRIGGERS.trg)
*/
DECLARE
  type TMS is record(RID ROWID, L NUMBER,T NUMBER,R NUMBER,B NUMBER);
  type TMSS is TABLE of TMS;
  MSS TMSS;
  i PLS_INTEGER;
BEGIN
  -- Находим все существующие диапазоны, с которыми появилось пересечение у
  -- изменённого диапазона.
  -- Изменяем диапазон, объединяя все пересекающиеся диапазоны.
  -- Удаляем найденные диапазоны.
  select ROWID RID, L,T,R,B bulk collect into MSS from KOCEL.MERGED_CELLS
    where upper(SHEET)=upper(:OLD.SHEET)
      and upper(BOOK)=upper(:OLD.BOOK)
      and L between :NEW.L and :NEW.R 
      and T between :NEW.T and :NEW.B 
      and R between :NEW.L and :NEW.R
      and B between :NEW.T and :NEW.B;
  if MSS.count=0 then
  update KOCEL.MERGED_CELLS set 
    L=:NEW.L,
    T=:NEW.T,
    R=:NEW.R,
    B=:NEW.B
    where L =:OLD.L
      and T =:OLD.T 
      and R =:OLD.R 
      and B =:OLD.B 
      and  upper(SHEET)=upper(:OLD.SHEET)
      and upper(BOOK)=upper(:OLD.BOOK);
    return;    
  end if;      
  i:=MSS.last;      
  for j in MSS.first..i   
  loop
    MSS(i).L:=least(MSS(i).L, MSS(j).L);
    MSS(i).T:=least(MSS(i).T, MSS(j).T);
    MSS(i).R:=greatest(MSS(i).R,MSS(j).R);
    MSS(i).L:=greatest(MSS(i).B,MSS(j).B);
    delete from KOCEL.MERGED_CELLS where ROWID=MSS(j).RID;
  end loop;    
  MSS(i).L:=least(MSS(i).L,:NEW.L);
  MSS(i).T:=least(MSS(i).T,:NEW.T);
  MSS(i).R:=greatest(MSS(i).R,:NEW.R);
  MSS(i).L:=greatest(MSS(i).B,:NEW.B);
  update KOCEL.MERGED_CELLS set 
    L=MSS(i).L,
    T=MSS(i).T,
    R=MSS(i).R,
    B=MSS(i).B
    where L =:OLD.L
      and T =:OLD.T 
      and R =:OLD.R 
      and B =:OLD.B 
      and  upper(SHEET)=upper(:OLD.SHEET)
      and upper(BOOK)=upper(:OLD.BOOK);
END;
/
/* 
INSTEAD_OF_DELETE--------------------------------------------------------------
*/
CREATE OR REPLACE TRIGGER KOCEL.V_MERGED_CELLS_id
INSTEAD OF DELETE ON KOCEL.V_MERGED_CELLS
/*  (KOCEL_TRIGGERS.trg)
*/
BEGIN
  delete from KOCEL.V_MERGED_CELLS 
    where L=:OLD.L 
      and T=:OLD.T 
      and upper(SHEET)=upper(:OLD.SHEET) 
      and upper(BOOK)=upper(:OLD.BOOK);
END;
/

/* *****************************************************************************
*/

/*  Presentation of sheet parameters.
INSTEAD_OF_INSERT--------------------------------------------------------------
*/
CREATE OR REPLACE TRIGGER KOCEL.V_SHEET_PARS_II
INSTEAD OF INSERT ON KOCEL.V_SHEET_PARS
-- (KOCEL_TRIGGERS.trg)
DECLARE
  NewN NUMBER;
  NewD DATE;
  NewS KOCEL.SHEET_PARS.S%type;
BEGIN
  NewN:=null;
  NewD:=null;
  NewS:=null;
  -- Если значение типа параметра не нулл, то поступаем в соответствии
  -- с типом.
  case 
    when upper(:NEW.T)='N' then 
      NewN:=:NEW.N;
      NewD:=null;
      NewS:=null;
    when upper(:NEW.T)='D' then 
      NewN:=null;
      NewD:=:NEW.D;
      NewS:=null;
    when upper(:NEW.T)='S' then 
      NewN:=null;
      NewD:=null;
      NewS:=:NEW.S;
    -- Если тип нулл, то порядок следующий.
    when :NEW.T is null then 
      case 
        -- Если поле "N" не нулл, то заносим в базу только поле "N".
        when :NEW.N is not null then NewN:=:NEW.N; 
        -- Если поле "D" не нулл, то заносим в базу только поле "D".
        when :NEW.D is not null then NewD:=:NEW.D; 
      -- Если поля "N" и "D" нулл, то заносим в базу только поле "S".
      else
        NewS:=:NEW.S;
      end case;
  else
    -- Если тип не правильно определён, то ошибка.
    raise_application_Error(-20443,
      'KOCEL.V_SHEET_PARS_ii. Неверный тип параметра!');
  end case;
  -- Если параметр уже существует, то редактируем значение параметра.
  -- Если параметр отсутствует, то добавляем новый параметр.
  if :NEW.SHEET is null then
    update KOCEL.SHEET_PARS
      set N=NewN,D=NewD,S=NewS
      where SHEET ='*'
        and BOOK=:NEW.BOOK
        and PAR_NAME=:NEW.PAR_NAME;
  else
    update KOCEL.SHEET_PARS
      set N=NewN,D=NewD,S=NewS
      where SHEET=:NEW.SHEET
        and BOOK=:NEW.BOOK
        and PAR_NAME=:NEW.PAR_NAME;
  end if;
  if SQL%rowcount = 0 then
    if :NEW.SHEET is null then
	    insert into KOCEL.SHEET_PARS 
	      values(:NEW.PAR_NAME,NewN,NewD,NewS,'*',:NEW.BOOK);
    else  
	    insert into KOCEL.SHEET_PARS 
	      values(:NEW.PAR_NAME,NewN,NewD,NewS,:NEW.SHEET,:NEW.BOOK);
    end if;  
  end if;
END;
/
/* 
INSTEAD_OF_UPDATE--------------------------------------------------------------
*/
CREATE OR REPLACE TRIGGER KOCEL.V_SHEET_PARS_IU
INSTEAD OF UPDATE ON KOCEL.V_SHEET_PARS
-- (KOCEL_TRIGGERS.trg)
DECLARE
  NewN NUMBER;
  NewD DATE;
  NewS KOCEL.SHEET_PARS.S%type;
BEGIN
  NewN:=null;
  NewD:=null;
  NewS:=null;
  -- Если значение типа параметра не нулл, то поступаем в соответствии
  -- с типом.
  case 
    when upper(:NEW.T)='N' then 
      NewN:=:NEW.N;
      NewD:=null;
      NewS:=null;
    when upper(:NEW.T)='D' then
      NewN:=null;
      NewD:=:NEW.D;
      NewS:=null;
    when upper(:NEW.T)='S' then 
      NewN:=null;
      NewD:=null;
      NewS:=:NEW.S;
    -- Если тип нулл, то порядок следующий.
    when :NEW.T is null then 
      case 
        -- Если поле "N" не нулл, то заносим в базу только поле "N".
        when :NEW.N is not null then NewN:=:NEW.N; 
        -- Если поле "D" не нулл, то заносим в базу только поле "D".
        when :NEW.D is not null then NewD:=:NEW.D; 
      -- Если поля "N" и "D" нулл, то заносим в базу только поле "S".
      else
        NewS:=:NEW.S;
      end case;
  else
    -- Если тип не правильно определён, то ошибка.
    raise_application_Error(-20443,
      'KOCEL.V_SHEET_PARS_ii. Неверный тип параметра!');
  end case;
  -- Если изменено имя параметра, книга или лист, то ошибка.
  if not(     :OLD.PAR_NAME=:NEW.PAR_NAME 
         and (   (:NEW.SHEET=:OLD.SHEET)
              or (:NEW.SHEET is null and :OLD.SHEET is null) )
         and :OLD.BOOK=:NEW.BOOK)      
  then
    raise_application_Error(-20443,
      'KOCEL.V_SHEET_PARS_iu. Изменить можно только значение параметра!');
  end if;
  -- Если параметр уже существует, то редактируем значение параметра.
  -- Если параметр отсутствует, то добавляем новый параметр.
  if :OLD.SHEET is null then
    update KOCEL.SHEET_PARS
      set N=NewN,D=NewD,S=NewS
      where SHEET ='*'
        and BOOK=:OLD.BOOK
        and PAR_NAME=:OLD.PAR_NAME;
  else
    update KOCEL.SHEET_PARS
      set N=NewN,D=NewD,S=NewS
      where SHEET=:OLD.SHEET
        and BOOK=:OLD.BOOK
        and PAR_NAME=:OLD.PAR_NAME;
  end if;
  if SQL%rowcount = 0 then
    if :OLD.SHEET is null then
	    insert into KOCEL.SHEET_PARS 
	      values(:OLD.PAR_NAME,NewN,NewD,NewS,'*',:OLD.BOOK);
    else    
	    insert into KOCEL.SHEET_PARS 
	      values(:OLD.PAR_NAME,NewN,NewD,NewS,:OLD.SHEET,:OLD.BOOK);
    end if;    
  end if;
END;
/
/* 
INSTEAD_OF_DELETE--------------------------------------------------------------
*/
CREATE OR REPLACE TRIGGER KOCEL.V_SHEET_PARS_id
INSTEAD OF DELETE ON KOCEL.V_SHEET_PARS
/*  (KOCEL_TRIGGERS.trg)
*/
BEGIN
  return;
END;
/

/* *****************************************************************************
end of triggers 
*/

