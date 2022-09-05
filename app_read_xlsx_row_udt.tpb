CREATE OR REPLACE TYPE BODY app_read_xlsx_row_udt AS
    MEMBER FUNCTION get(p_i NUMBER)
    RETURN SYS.anydata
    AS
    BEGIN
        RETURN aa(p_i);
    END get;
END;
/
show errors
