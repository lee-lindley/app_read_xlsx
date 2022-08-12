CREATE OR REPLACE PACKAGE app_read_xlsx
AUTHID CURRENT_USER
AS
    PROCEDURE xlsx_to_ptt(
        p_xlsx      BLOB
        ,p_sheets   VARCHAR2 := NULL
        ,p_cell     VARCHAR2 := NULL
        ,p_ptt_name VARCHAR2 := 'XLSX_PTT'
    );
END app_read_xlsx ;
/
show errors
