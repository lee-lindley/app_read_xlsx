CREATE OR REPLACE TYPE BODY app_read_xlsx_row_udt AS
    MEMBER FUNCTION get(p_i NUMBER)
    RETURN anydata
    AS
    BEGIN
        RETURN aad(p_i).ad;
    END get;
END;
/
show errors