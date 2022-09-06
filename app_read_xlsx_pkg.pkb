CREATE OR REPLACE PACKAGE BODY app_read_xlsx_pkg 
AS
    TYPE t_ctx_cache IS TABLE OF NUMBER(10) INDEX BY BINARY_INTEGER;
    g_ctx_cache t_ctx_cache;

    FUNCTION create_ctx RETURN NUMBER
    IS
        v_ctx   NUMBER(10);
        v_cnt   BINARY_INTEGER;
    BEGIN
        -- we do not try to reuse gaps
        IF g_ctx_cache.LAST IS NULL THEN
            v_ctx := 1;
        ELSE
            v_ctx := g_ctx_cache(g_ctx_cache.LAST) + 1;
        END IF;
        g_ctx_cache(v_ctx) := v_ctx;
        RETURN v_ctx;
    END create_ctx
    ;

    PROCEDURE close_ctx(p_ctx NUMBER)
    IS
    BEGIN
        DELETE FROM as_read_xlsx_gtt WHERE ctx = p_ctx;
        COMMIT; -- required for autonomous_transaction
        g_ctx_cache.DELETE(p_ctx);
    END close_ctx;

    PROCEDURE parse_blob(
        p_ctx       NUMBER
        ,p_xlsx     BLOB
        ,p_sheets   VARCHAR2 := NULL
        ,p_cell     VARCHAR2 := NULL
    ) IS
        PRAGMA AUTONOMOUS_TRANSACTION;
    BEGIN
        -- prepare the GTT defined in package schema for all data from the spreadsheet
        DELETE FROM as_read_xlsx_gtt WHERE ctx = p_ctx;
        INSERT INTO as_read_xlsx_gtt(
                ctx, sheet_nr, sheet_name, row_nr, col_nr, cell, cell_type, string_val, number_val, date_val, formula
            )
        SELECT p_ctx, 
                sheet_nr, sheet_name, row_nr, col_nr, cell, cell_type, string_val, number_val, date_val, formula
        FROM TABLE( AS_READ_XLSX.read(p_xlsx, p_sheets, p_cell) ) t
        ;
        COMMIT; -- required for autonomous_transaction
        

    END parse_blob;


    
END app_read_xlsx_pkg
;
/
show errors

