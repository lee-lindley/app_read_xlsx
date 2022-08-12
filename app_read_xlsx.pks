CREATE OR REPLACE PACKAGE app_read_xlsx
AUTHID CURRENT_USER
AS
-- will fail if you do an alter session set current_schema before calling. Jacks up private temp tables
    PROCEDURE xlsx_to_ptt(
        p_xlsx      BLOB
        ,p_sheets   VARCHAR2 := NULL
        ,p_cell     VARCHAR2 := NULL
        ,p_ptt_name VARCHAR2 := 'ora$ptt_XLSX' -- must start with ora$ptt unless changed by dba
    );
END app_read_xlsx ;
/
show errors
