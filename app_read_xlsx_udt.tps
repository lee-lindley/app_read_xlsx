CREATE OR REPLACE TYPE app_read_xlsx_udt FORCE AS OBJECT (
    -- example:
    --   variable c REFCURSOR
    -- declare
    --    v_xlsx app_read_xlsx_udt;
    --    v_sql CLOB;
    --    v_src SYS_REFCURSOR
    -- begin
    --    v_xlsx := app_read_xlsx_udt(to_blob(bfilename('TMP_DIR','Book1.xlsx')), '1');
    --    v_sql := 'SELECT' ||v_xlsx.get_col_sql('X')||' FROM TABLE( app_read_xlsx_pkg.get_rows(:ctx, :cnt) ) X';
    --    OPEN v_src FOR v_sql USING v_xlsx.ctx, v_xlsx.get_col_count;
    --    :c := v_src;
    -- end;
    -- 
    ctx         NUMBER(10)
    ,sheet_name VARCHAR2(128)
    ,col_names  &&d_arr_varchar2_udt.
    ,CONSTRUCTOR FUNCTION app_read_xlsx_udt(
        p_xlsx      BLOB
        ,p_sheets   VARCHAR2 := NULL
        ,p_cell     VARCHAR2 := NULL
    ) RETURN SELF AS RESULT
    ,CONSTRUCTOR FUNCTION app_read_xlsx_udt(
        p_ctx   NUMBER
    ) RETURN SELF AS RESULT
    ,MEMBER PROCEDURE destructor -- clears context and global temporary table records for this spreadsheet
    ,MEMBER FUNCTION get_col_names RETURN &&d_arr_varchar2_udt.
    ,MEMBER FUNCTION get_col_count RETURN NUMBER
    ,MEMBER FUNCTION get_sql RETURN CLOB
    ,MEMBER FUNCTION get_col_sql(p_oname VARCHAR2 := 'R') RETURN CLOB

);
/
show errors
