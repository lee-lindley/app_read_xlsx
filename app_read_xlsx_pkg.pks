CREATE OR REPLACE PACKAGE app_read_xlsx_pkg ACCESSIBLE BY (app_read_xlsx_udt)
AS
    FUNCTION create_ctx RETURN NUMBER;
    PROCEDURE close_ctx(p_ctx NUMBER);
    PROCEDURE parse_blob(
        p_ctx       NUMBER
        ,p_xlsx     BLOB
        ,p_sheets   VARCHAR2 := NULL
        ,p_cell     VARCHAR2 := NULL
    );

END app_read_xlsx_pkg ;
/
show errors
