CREATE OR REPLACE TYPE BODY app_read_xlsx_udt
AS
    MEMBER PROCEDURE destructor -- clears context and global temporary table records for this spreadsheet
    IS
    BEGIN
        app_read_xlsx_pkg.close_ctx(ctx);
    END destructor
    ;
    CONSTRUCTOR FUNCTION app_read_xlsx_udt(
        p_xlsx      BLOB
        ,p_sheets   VARCHAR2 := NULL
        ,p_cell     VARCHAR2 := NULL
    ) RETURN SELF AS RESULT
    IS
    BEGIN
        ctx := app_read_xlsx_pkg.create_ctx;
        app_read_xlsx_pkg.parse_blob(
            p_ctx   => ctx
            ,p_xlsx => p_xlsx
            ,p_sheets   => p_sheets
            ,p_cell     => p_cell
        );
        SELECT string_val BULK COLLECT INTO col_names
        FROM as_read_xlsx_gtt t
        WHERE t.ctx = SELF.ctx AND t.row_nr = 1
        ORDER BY col_nr
        ;

        RETURN;
    END app_read_xlsx_udt
    ;

    MEMBER FUNCTION get_col_names RETURN &&d_arr_varchar2_udt.
    IS
    BEGIN
        RETURN col_names;
    END get_col_names
    ;

    MEMBER FUNCTION get_col_count RETURN NUMBER
    IS
    BEGIN
        RETURN col_names.COUNT;
    END get_col_count
    ;

    MEMBER FUNCTION get_sql 
    RETURN CLOB
    IS
        v_col_count     NUMBER := SELF.get_col_count();
        v_sql           CLOB;
    BEGIN
        v_sql := q'{
WITH app_read_xlsx_cols AS (
    SELECT level AS c FROM dual CONNECT BY level <= }'||TO_CHAR(v_col_count)||q'{
), app_read_xlsx_t AS (
    SELECT row_nr - 1 AS data_row_nr, col_nr, cell_type, string_val, date_val, number_val
    FROM as_read_xlsx_gtt t
    WHERE t.ctx = }'||TO_CHAR(SELF.ctx)||q'{ AND t.row_nr > 1 AND t.col_nr <= }'||TO_CHAR(v_col_count)||q'{
), app_read_xlsx_b AS (
    SELECT
        t.data_row_nr
        ,cols.c AS col_nr
        ,anydata_shell_udt(
            CASE t.cell_type
                WHEN 'S' THEN SYS.ANYDATA.convertVarchar2(t.string_val)
                WHEN 'D' THEN SYS.ANYDATA.convertDate(t.date_val)
                WHEN 'N' THEN SYS.ANYDATA.convertNumber(t.number_val)
            END 
        ) AS ad
    FROM app_read_xlsx_t t
    PARTITION BY (t.data_row_nr)    -- fill in gaps for empty cells
    RIGHT OUTER JOIN app_read_xlsx_cols cols 
        ON cols.c = t.col_nr
), app_read_xlsx_c AS (
    SELECT data_row_nr, CAST( COLLECT(ad ORDER BY col_nr) AS arr_anydata_shell_udt) AS arr_ad
    FROM app_read_xlsx_b b
    GROUP BY data_row_nr
), app_read_xlsx_d AS (
    SELECT data_row_nr
        ,app_read_xlsx_row_udt(arr_ad) AS ad
    FROM app_read_xlsx_c c
), app_read_xlsx_sql AS ( 
  SELECT
}'|| SELF.get_col_sql||q'{
  FROM app_read_xlsx_d R
)}';
        RETURN v_sql;
    END get_sql
    ;

    MEMBER FUNCTION get_col_sql(p_oname VARCHAR2 := 'R')
    RETURN CLOB
    IS
        v_sql   CLOB;
        v_comma CONSTANT VARCHAR2(8) := '
    ,';
    BEGIN
        v_sql := '   '||p_oname||'.ad.get(1) AS "'||col_names(1)||'"';
        FOR i IN 2..col_names.COUNT
        LOOP
            v_sql := v_sql||v_comma||p_oname||'.ad.get('||TO_CHAR(i)||') AS "'||col_names(i)||'"';
        END LOOP;
        RETURN v_sql;
    END get_col_sql
    ;


END;
/
show errors
