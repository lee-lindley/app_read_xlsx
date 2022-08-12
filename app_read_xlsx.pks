CREATE OR REPLACE PACKAGE app_read_xlsx
AUTHID CURRENT_USER
AS
-- will fail if you do an alter session set current_schema before calling. Jacks up private temp tables
    PROCEDURE xlsx_to_ptt(
        -- create a session local private temporary table (survives commit) containing spreadsheet data.
        -- column names will match first row of spreadsheet case preserved with " " (beware illegal names).
        -- Data will be all VARCHAR2.
        -- Additional column added to the start -- data_row_nr -- with row after the header being 1.
        -- Uses session nls settings for date and number conversions
        -- BEWARE procedure does DDL which will result in a commit in your session.
        p_xlsx      BLOB
        ,p_sheets   VARCHAR2 := NULL
        ,p_cell     VARCHAR2 := NULL
        ,p_ptt_name VARCHAR2 := 'ora$ptt_XLSX' -- must start with ora$ptt unless changed by dba
    );
END app_read_xlsx ;
/
show errors
