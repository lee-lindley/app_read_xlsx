whenever sqlerror continue
DROP TABLE as_read_xlsx_gtt;
prompt ok for drop to fail for table does not exist
whenever sqlerror exit failure
CREATE GLOBAL TEMPORARY TABLE as_read_xlsx_gtt(
    -- these must match type tp_one_cell in package as_read_xlsx
    sheet_nr number(2)
    , sheet_name varchar(4000)
    , row_nr number(10)
    , col_nr number(10)
    , cell varchar2(100)
    , cell_type varchar2(1)
    , string_val varchar2(4000)
    , number_val number
    , date_val date
    , formula varchar2(4000)
) ON COMMIT PRESERVE ROWS
;
