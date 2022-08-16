CREATE OR REPLACE TYPE app_read_xlsx_row_udt FORCE AS OBJECT (
    -- purpose is to be able to extract an anydata object from an array in SQL
    -- which cannot use collection element syntax
    aad     arr_anydata_shell_udt
    ,MEMBER FUNCTION get(p_i NUMBER) RETURN SYS.anydata
);
/
show errors
