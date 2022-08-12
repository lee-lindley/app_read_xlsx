CREATE OR REPLACE PACKAGE BODY app_read_xlsx
AS
    PROCEDURE xlsx_to_ptt(
        p_xlsx      BLOB
        ,p_sheets   VARCHAR2 := NULL
        ,p_cell     VARCHAR2 := NULL
        ,p_ptt_name VARCHAR2 := 'XLSX_PTT'
    ) IS
        v_cnt       BINARY_INTEGER;
        v_tbl_sql   CLOB;
        v_sql       CLOB;
        v_comma     VARCHAR2(2) := '
,';
    BEGIN
        DELETE FROM as_read_xlsx_gtt;
        INSERT INTO as_read_xlsx_gtt
        SELECT * 
        FROM TABLE( AS_READ_XLSX.read(p_xlsx, p_sheets, p_cell) )
        ;
        SELECT COUNT(*) INTO v_cnt
        FROM as_read_xlsx_gtt
        WHERE row_nr = 1
        ;
        IF v_cnt = 0 THEN
            raise_application_error(-20222, 'app_read_xlsx.xlsx_to_ptt found no data in input blob for sheets='||NVL(p_sheets,'NULL')||', cell='||NVL(p_cell,'NULL'));
        END IF;

        v_sql := 'DROP TABLE ora$ptt_csv';
        BEGIN
            EXECUTE IMMEDIATE v_sql;
        EXCEPTION WHEN OTHERS THEN NULL;
        END;

        v_tbl_sql := 'CREATE PRIVATE TEMPORARY TABLE ora$ptt_xlsx(
row_nr      NUMBER(10)';

        v_sql := 'INSERT /*+ APPEND */ INTO ora$ptt_xlsx
WITH a AS (
    SELECT row_nr, col_nr, cell_type, string_val, number_val, date_val
    FROM as_read_xlsx_gtt
    WHERE row_nr > 1 AND col_nr <= '||TO_CHAR(v_cnt)||'
) SELECT row_nr';

        FOR r IN (
            SELECT string_val, col_nr
            FROM as_read_xlsx_gtt
            WHERE row_nr = 1
            ORDER BY col_nr
        ) LOOP
            v_tbl_sql := v_tbl_sql||v_comma||'"'||r.string_val||'" SYS.ANYDATA';
            v_sql := v_sql||v_comma||'MAX(CASE WHEN col_nr = '||TO_CHAR(r.col_nr)||q'! THEN
    CASE cell_type
        WHEN 'S' THEN SYS.ANYDATA.convertVarchar2(string_val)
        WHEN 'D' THEN SYS.ANYDATA.convertDate(date_val)
        WHEN 'N' THEN SYS.ANYDATA.convertNumber(number_val)
    END
END) AS "!'||r.string_val||'"';
        END LOOP;
        v_tbl_sql := v_tbl_sql||'
)';
        DBMS_OUTPUT.put_line('v_tbl_sql='||v_tbl_sql);
        EXECUTE IMMEDIATE v_tbl_sql;

        v_sql := v_sql||'
FROM a
GROUP BY row_nr';
        DBMS_OUTPUT.put_line('v_sql='||v_sql);
        EXECUTE IMMEDIATE v_sql;


    END xlsx_to_ptt
    ;
END app_read_xlsx;
/
show errors
