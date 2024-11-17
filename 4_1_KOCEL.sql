/*  The script for creating the KOCEL scheme 
 create 25.04.2009
 by Nikolay Krasilnikov
 This file is distributed under Apache License 2.0 (01.2004).
   http://www.apache.org/licenses/
 update 28.01.2010 05.11.2010 26.03.2011 28.08.2014 17.08.2021 15.11.2022
        23.12.2023 

*****************************************************************************
 Let's make users of "KOCEL" and "KOCELSYS".
*/
declare
  tmpVar NUMBER;
begin
  select count(*) into tmpVar from ALL_USERS where USERNAME='KOCEL';
	if tmpVar>0 then
	  execute immediate('DROP USER "KOCEL" Cascade');
  end if;  
  select count(*) into tmpVar from ALL_USERS where USERNAME='KOCELSYS';
	if tmpVar>0 then
	  execute immediate('DROP USER "KOCELSYS" Cascade');
  end if;  
end;
/
declare
  tmpVar NUMBER;
begin
  select count(*)into tmpVar from dual where exists 
  ( select TABLESPACE_NAME from dba_tablespaces 
    where TABLESPACE_NAME = 'KOCEL_FILES');
  if tmpVar = 0 then 
    execute immediate('
      CREATE TABLESPACE KOCEL_FILES DATAFILE 
      ''KOCEL_FILES'' SIZE 10M REUSE AUTOEXTEND ON NEXT 10M MAXSIZE UNLIMITED
    ');
  end if;
end;
/

CREATE USER "KOCEL" IDENTIFIED BY "K"
  DEFAULT TABLESPACE KOCEL_FILES
  TEMPORARY TABLESPACE TEMP
  QUOTA UNLIMITED ON KOCEL_FILES;
CREATE USER "KOCELSYS" IDENTIFIED BY "K"
  DEFAULT TABLESPACE KOCEL_FILES
  TEMPORARY TABLESPACE TEMP
  QUOTA UNLIMITED ON KOCEL_FILES;
GRANT CREATE SESSION TO "KOCEL" ;
GRANT CREATE PROCEDURE  TO "KOCEL" ;
GRANT CREATE TABLE  TO "KOCEL" ;
GRANT CREATE VIEW  TO "KOCEL" ;
GRANT SELECT ON SYS.V_$SESSION TO "KOCELSYS";
GRANT ALTER SYSTEM TO "KOCELSYS";
GRANT execute on SYS.dbms_LOCK to "KOCELSYS";
CREATE SEQUENCE KOCELSYS.SCRIPT_SEQ
  START WITH 1
  INCREMENT BY 1
  MINVALUE 1
  NOCYCLE 
  NOORDER
  NOCACHE;
/
CREATE SEQUENCE KOCEL.SEQ
  START WITH 100
  INCREMENT BY 100
  MINVALUE 100
  NOCYCLE 
  NOORDER
  CACHE 10;
/
CREATE SEQUENCE KOCEL.FORMAT_SEQ
  START WITH 100
  INCREMENT BY 100
  MINVALUE 100
  NOCYCLE 
  NOORDER
  NOCACHE;
/
/*  Types.
*/
@"KOCEL_TFORMAT.tps"
GRANT execute on KOCEL.TFormat to public;
@"KOCEL_TYPES.sql"
/*  Tables.
*/
@"KOCELSYS_TABLES.sql"
@"KOCEL_TABLES.sql"
insert into KOCEL.FORMATS (Fmt) values (0);
/*  The trigger.
*/
@"KOCEL-TRIGGERS.trg"
/*  FLEXCEL API Procedures
*/
@"KOCEL-PROCEDURES.prc"
/*  Test lines.
*/
begin
  KOCEL.UPDATE_CELL(1,1,
  null,null,'A1',null,
	'Q','Q');
  KOCEL.UPDATE_CELL(1,2,
  null,null,'B1',null,
	'Q','Q');
  KOCEL.UPDATE_CELL(2,1,
  null,null,'A2',null,
	'Q','Q');
  KOCEL.UPDATE_CELL(2,2,
  null,null,'B2',null,
	'Q','Q');
	commit;
end;
/
/*  The system package.
*/
@"KOCELSYS-S.pks"
/*  The formatting package.
*/
@"KOCEL-FORMAT.pks"
/*  System procedures.
*/
@"KOCELSYS-PROCEDURES.prc"

/*  System triggers.
*/
GRANT ADMINISTER DATABASE TRIGGER to KOCELSYS;
@"KOCELSYS-TRIGGERS.trg"
/*  A package for use in data processing scripts.
*/
@"KOCEL-CELL.pks"
/*  A package of requests for book data.
*/
@"KOCEL-BOOK.pks"
/*  Cell Import Package
*/
@"KOCEL-IMPORT.pks"
@"KOCEL-IMPORT.pkb"
/*  Bodies of types and packages.
*/
@"KOCEL-CELL.pkb"
@"KOCEL-TROW.tpb"
@"KOCEL-TSROW.tpb"
@"KOCEL-TVALUE.tpb"
@"KOCEL_TFORMAT.tpb"
@"KOCEL-FORMAT.pkb"
@"KOCEL-BOOK.pkb"
GRANT EXECUTE on KOCEL.IMPORT to public;
CREATE OR REPLACE public synonym KIMP for KOCEL.IMPORT;
GRANT EXECUTE on KOCEL.CELL to public;
CREATE OR REPLACE public synonym CELL for KOCEL.CELL;
GRANT EXECUTE on KOCEL.BOOK to public;
CREATE OR REPLACE public synonym BOOK for KOCEL.CELL;
GRANT EXECUTE on KOCEL.FORMAT to public;
CREATE OR REPLACE public synonym KFM for KOCEL.FORMAT;
GRANT EXECUTE on KOCEL.TFORMAT to public;
CREATE OR REPLACE public synonym TFORMAT for KOCEL.TFORMAT;
/*  Performances.
*/
@"KOCEL-Views.vw"
/*  Recompile the objects with errors and output the result.
*/
declare
  tmpVar VARCHAR2(38);
begin
  select to_char(count(*)) into tmpVar from DBA_OBJECTS o 
     where  (o.STATUS='INVALID') and (o.OWNER='KOCEL');
	O('KOCEL Invalid Obj before: '||tmpVar);	 
   sys.utl_recomp.recomp_serial('KOCEL');
  select to_char(count(*)) into tmpVar from DBA_OBJECTS o 
     where  (o.STATUS='INVALID') and (o.OWNER='KOCEL');
	O('after: '||tmpVar);	 
  select to_char(count(*)) into tmpVar from DBA_OBJECTS o 
     where  (o.STATUS='INVALID') and (o.OWNER='KOCELSYS');
	O('KOCELSYS Invalid Obj before: '||tmpVar);	 
   sys.utl_recomp.recomp_serial('KOCELSYS');
  select to_char(count(*)) into tmpVar from DBA_OBJECTS o 
     where  (o.STATUS='INVALID') and (o.OWNER='KOCELSYS');
	O('after: '||tmpVar);
end;
/

