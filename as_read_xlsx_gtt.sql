whenever sqlerror continue
-- may be necessary in order to drop it
TRUNCATE TABLE as_read_xlsx_gtt; 
DROP TABLE as_read_xlsx_gtt;
prompt ok for trunc and drop to fail for table does not exist
CREATE GLOBAL TEMPORARY TABLE as_read_xlsx_gtt(
    -- these must match type tp_one_cell in package as_read_xlsx
    ctx             NUMBER(10)
    ,sheet_nr       NUMBER(2)
    , sheet_name    VARCHAR(4000)
    , row_nr        NUMBER(10)
    , col_nr        NUMBER(10)
    , cell          VARCHAR2(100)
    , cell_type     VARCHAR2(1)
    , string_val    VARCHAR2(4000)
    , number_val    NUMBER
    , date_val      DATE
    , formula       VARCHAR2(4000)
) ON COMMIT PRESERVE ROWS
;
prompt if table failed to create because already exists, only matters if it changed. Probably not
whenever sqlerror exit failure
