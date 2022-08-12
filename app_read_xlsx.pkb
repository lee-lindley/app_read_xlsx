CREATE OR REPLACE PACKAGE BODY app_read_xlsx
AS
    PROCEDURE xlsx_to_ptt(
        p_xlsx          BLOB
        ,p_sheets       VARCHAR2 := NULL
        ,p_cell         VARCHAR2 := NULL
        ,p_ptt_name     VARCHAR2 := 'ora$ptt_XLSX' -- must start with ora$ptt unless changed by dba
    ) IS
        TYPE t_colname_rec IS RECORD(
            string_val  VARCHAR2(4000)
            ,col_nr     NUMBER(10)
        );
        TYPE t_arr_colname IS TABLE OF t_colname_rec;
        v_arr_colname   t_arr_colname;
        v_cnt           BINARY_INTEGER;
        v_tbl_sql       CLOB;
        v_sql           CLOB;
        v_comma         VARCHAR2(2) := '
,';
    BEGIN
        EXECUTE IMMEDIATE 'DELETE FROM '||$$PLSQL_UNIT_OWNER||'.as_read_xlsx_gtt';
        EXECUTE IMMEDIATE 'INSERT INTO '||$$PLSQL_UNIT_OWNER||'.as_read_xlsx_gtt
        SELECT * 
        FROM TABLE( AS_READ_XLSX.read(p_xlsx, p_sheets, p_cell) )'
        ;
        EXECUTE IMMEDIATE 'SELECT COUNT(*) 
        FROM '||$$PLSQL_UNIT_OWNER||'.as_read_xlsx_gtt
        WHERE row_nr = 1' INTO v_cnt
        ;
        IF v_cnt = 0 THEN
            raise_application_error(-20222, 'app_read_xlsx.xlsx_to_ptt found no data in input blob for sheets='||NVL(p_sheets,'NULL')||', cell='||NVL(p_cell,'NULL'));
        END IF;

        v_sql := 'DROP TABLE '||p_ptt_name;
        BEGIN
            EXECUTE IMMEDIATE v_sql;
        EXCEPTION WHEN OTHERS THEN NULL;
        END;

        v_tbl_sql := 'CREATE PRIVATE TEMPORARY TABLE '||p_ptt_name||'(
row_nr      NUMBER(10)';

        v_sql := 'INSERT /*+ APPEND */ INTO '||p_ptt_name||'
WITH a AS (
    SELECT row_nr, col_nr, cell_type, string_val, number_val, date_val
    FROM '||$$PLSQL_UNIT_OWNER||'.as_read_xlsx_gtt
    WHERE row_nr > 1 AND col_nr <= '||TO_CHAR(v_cnt)||'
) SELECT row_nr';

        EXECUTE IMMEDIATE 'SELECT string_val, col_nr
            FROM '||$$PLSQL_UNIT_OWNER||'.as_read_xlsx_gtt
            WHERE row_nr = 1
            ORDER BY col_nr'
            BULK COLLECT INTO v_arr_colname
        ;
        FOR i IN 1..v_arr_colname.COUNT
        LOOP
            v_tbl_sql := v_tbl_sql||v_comma||'"'||v_arr_colname(i).string_val||'" SYS.ANYDATA';
            v_sql := v_sql||v_comma||'MAX(CASE WHEN col_nr = '||TO_CHAR(v_arr_colname(i).col_nr)||q'! THEN
    CASE cell_type
        WHEN 'S' THEN SYS.ANYDATA.convertVarchar2(string_val)
        WHEN 'D' THEN SYS.ANYDATA.convertDate(date_val)
        WHEN 'N' THEN SYS.ANYDATA.convertNumber(number_val)
    END
END) AS "!'||v_arr_colname(i).string_val||'"';
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
