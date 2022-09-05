CREATE OR REPLACE TYPE app_read_xlsx_row_udt FORCE AS OBJECT (
    data_row_nr NUMBER 
    ,aa         arr_anydata_udt
    ,MEMBER FUNCTION get(p_i NUMBER) RETURN SYS.anydata
);
/
show errors
