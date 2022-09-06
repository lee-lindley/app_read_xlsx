CREATE OR REPLACE TYPE app_read_xlsx_udt FORCE AS OBJECT (
    -- example:
--declare
--    v_o app_read_xlsx_udt;
--    v_sql CLOB;
--    v_ctxId         ExcelGen.ctxHandle;
--    v_sheetHandle   BINARY_INTEGER;
--begin
--    v_o := app_read_xlsx_udt(to_blob(bfilename('TMP_DIR' ,'Book1.xlsx')), '1');
--    v_sql := 'WITH a AS (
--'||v_o.get_sql||'
--) SELECT * FROM a';
-- dbms_output.put_line(v_sql);
--    /**/
--        v_ctxId := ExcelGen.createContext();
--        v_sheetHandle := ExcelGen.addSheetFromQuery(v_ctxId, 'app_read_xlsx demo', v_sql, p_sheetIndex => 1);
--        -- freeze the top row with the column headers
--        ExcelGen.setHeader(v_ctxId, v_sheetHandle, p_frozen => TRUE);
--        -- style with alternating colors on each row. 
--        ExcelGen.setTableFormat(v_ctxId, v_sheetHandle, 'TableStyleLight2');
--        ExcelGen.createFile(v_ctxId, 'TMP_DIR', 'app_read_xlsx_demo.xlsx');
--        ExcelGen.closeContext(v_ctxId);
--    /**/
--    --v_o.destructor;
--end;
--/
    ctx         NUMBER(10)
    ,sheet_name VARCHAR2(128)
    ,col_names  &&d_arr_varchar2_udt.
    --
    -- see as_read_xlsx for parameters
    ,CONSTRUCTOR FUNCTION app_read_xlsx_udt(
        p_xlsx      BLOB
        ,p_sheets   VARCHAR2 := NULL
        ,p_cell     VARCHAR2 := NULL
    ) RETURN SELF AS RESULT
    ,MEMBER FUNCTION get_sql(p_oname VARCHAR2 := 'X.R') RETURN CLOB
    -- these two functions do not include data_row_nr in list or count
    ,MEMBER FUNCTION get_col_names RETURN &&d_arr_varchar2_udt.
    ,MEMBER FUNCTION get_col_count RETURN NUMBER 
    -- utility to reuse one already built in the session. Unlikely you will ever need it
    ,CONSTRUCTOR FUNCTION app_read_xlsx_udt(
        p_ctx   NUMBER
    ) RETURN SELF AS RESULT
    ,MEMBER PROCEDURE destructor -- clears context and global temporary table records for this spreadsheet instance
    ,MEMBER FUNCTION get_ctx RETURN NUMBER
    ,MEMBER FUNCTION get_sheet_name RETURN VARCHAR2
    ,MEMBER FUNCTION get_col_sql(p_oname VARCHAR2 := 'X.R') RETURN CLOB
    -- used from the dynamic SQL returned to you by get_sql()
    ,STATIC FUNCTION get_data_rows(
         p_ctx      NUMBER
        ,p_col_cnt  NUMBER
    ) RETURN arr_app_read_xlsx_row_udt PIPELINED
    -- internal use
    ,MEMBER PROCEDURE app_read_xlsx_constructor(
        SELF IN OUT NOCOPY  app_read_xlsx_udt
        ,p_xlsx             BLOB DEFAULT NULL
        ,p_sheets           VARCHAR2 := NULL
        ,p_cell             VARCHAR2 := NULL
        ,p_ctx              NUMBER := NULL
    )

);
/
show errors
