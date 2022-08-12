CREATE OR REPLACE PACKAGE BODY app_read_xlsx
AS
    PROCEDURE xlsx_to_ptt(
        -- create a session local private temporary table (survives commit) containing spreadsheet data.
        -- column names will match first row of spreadsheet case preserved with " " (beware illegal names).
        -- Data will be all VARCHAR2.
        -- Additional column added to the start -- data_row_nr -- with row after the header being 1.
        -- Uses session nls settings for date and number conversions
        -- BEWARE procedure does DDL which will result in a commit in your session.
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
        -- prepare the GTT defined in package schema for all data from the spreadsheet
        EXECUTE IMMEDIATE 'TRUNCATE TABLE '||$$PLSQL_UNIT_OWNER||'.as_read_xlsx_gtt';
        EXECUTE IMMEDIATE 'INSERT INTO '||$$PLSQL_UNIT_OWNER||'.as_read_xlsx_gtt
            SELECT * 
            FROM TABLE( '||$$PLSQL_UNIT_OWNER||'.AS_READ_XLSX.read(:p_xlsx, :p_sheets, :p_cell) )'
            USING p_xlsx, p_sheets, p_cell
        ;
        -- we only care about columns that have an entry in the header row.
        EXECUTE IMMEDIATE 'SELECT COUNT(*) 
        FROM '||$$PLSQL_UNIT_OWNER||'.as_read_xlsx_gtt
        WHERE row_nr = 1' INTO v_cnt
        ;
        IF v_cnt = 0 THEN
            raise_application_error(-20222, 'app_read_xlsx.xlsx_to_ptt found no data in input blob for sheets='||NVL(p_sheets,'NULL')||', cell='||NVL(p_cell,'NULL'));
        END IF;

        -- prepare to create a private temporary table. An oddity is that if user session changed current_schema,
        -- creating an unqualifed name PTT doesn't work. Weird bug since the darn thing only exists in session memory.
        -- This safely creates it in the session user schema no matter how we are called.
        v_sql := 'DROP TABLE '||user||'.'||p_ptt_name;
        BEGIN
            EXECUTE IMMEDIATE v_sql;
        EXCEPTION WHEN OTHERS THEN NULL;
        END;

        -- start the table create and the insert sql statements 
        v_tbl_sql := 'CREATE PRIVATE TEMPORARY TABLE '||user||'.'||p_ptt_name||'(
data_row_nr      NUMBER(10)';

        v_sql := 'INSERT /*+ APPEND */ INTO '||user||'.'||p_ptt_name||'
WITH a AS (
    SELECT row_nr, col_nr, cell_type, string_val, number_val, date_val
    FROM '||$$PLSQL_UNIT_OWNER||'.as_read_xlsx_gtt
    WHERE row_nr > 1 AND col_nr <= '||TO_CHAR(v_cnt)||'
) SELECT row_nr - 1 AS data_row_nr';

        -- get the column names
        EXECUTE IMMEDIATE 'SELECT string_val, col_nr
            FROM '||$$PLSQL_UNIT_OWNER||'.as_read_xlsx_gtt
            WHERE row_nr = 1
            ORDER BY col_nr'
            BULK COLLECT INTO v_arr_colname
        ;
        FOR i IN 1..v_arr_colname.COUNT
        LOOP
            -- add this column name to the ptt definition
            v_tbl_sql := v_tbl_sql||v_comma||'"'||v_arr_colname(i).string_val||'" VARCHAR2(4000)';
            -- Will be grouped by row number. Pivot the row value into the right column and convert from
            -- whatever format the column value for this row was in excel to varchar2. Note that in excel
            -- you can have dates and text in different cells of the same column, so you cannot rely on a
            -- type for all values in a column. Thus convert everthing to text.
            v_sql := v_sql||v_comma||'MAX(CASE WHEN col_nr = '||TO_CHAR(v_arr_colname(i).col_nr)||q'! THEN
    CASE cell_type
        WHEN 'S' THEN string_val
        WHEN 'D' THEN TO_CHAR(date_val)
        WHEN 'N' THEN TO_CHAR(number_val)
    END
END) AS "!'||v_arr_colname(i).string_val||'"';

        END LOOP;

        -- finish the PTT definition. Make it last even if caller does a commit. Goes away when session ends.
        v_tbl_sql := v_tbl_sql||'
) ON COMMIT PRESERVE DEFINITION';
        DBMS_OUTPUT.put_line('v_tbl_sql='||v_tbl_sql);
        EXECUTE IMMEDIATE v_tbl_sql;

        -- finish the sql that pivots the cell data into rows and inserts into the PTT. Populate the PTT
        v_sql := v_sql||'
FROM a
GROUP BY row_nr';
        DBMS_OUTPUT.put_line('v_sql='||v_sql);
        EXECUTE IMMEDIATE v_sql;

        DBMS_OUTPUT.put_line(p_ptt_name||' is populated with data from the spreadsheet');

    END xlsx_to_ptt
    ;
END app_read_xlsx;
/
show errors
