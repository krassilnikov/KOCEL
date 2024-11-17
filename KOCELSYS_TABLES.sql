/*  KOCELSYS  tables
 create 14.05.2010
 by Nikolay Krasilnikov
 This file is distributed under Apache License 2.0 (01.2004).
   http://www.apache.org/licenses/
 update 
*/
CREATE TABLE KOCELSYS.SCRIPTS( 
  SCRIPT NUMBER not null,
	ScriptSession NUMBER not null,
	
	CONSTRAINT PK_SHEETS PRIMARY KEY (SCRIPT)
);

COMMENT ON TABLE KOCELSYS.SCRIPTS
  IS 'A list of running scripts.';
  
COMMENT ON COLUMN KOCELSYS.SCRIPTS.SCRIPT
  IS 'The script ID.';  
COMMENT ON COLUMN KOCELSYS.SCRIPTS.ScriptSession
  IS 'The script session.';  

