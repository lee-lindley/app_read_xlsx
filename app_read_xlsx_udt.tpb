CREATE OR REPLACE TYPE BODY app_read_xlsx_udt
AS
-- see documentation at
-- https://github.com/lee-lindley/app_read_xlsx
/*
MIT License

Copyright (c) 2022 Lee Lindley

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
*/

    -- clears context and global temporary table records for this spreadsheet.
    -- You generally don't need to worry about it. If you are running multiple times
    -- in a session for one or more spreadsheet inputs, the global temporary table can
    -- fill up with data. Again, probably doesn't matter, but this gives you a way
    -- to clean up.
    MEMBER PROCEDURE destructor 
    IS
    BEGIN
        app_read_xlsx_pkg.close_ctx(ctx);
    END destructor
    ;
    -- for internal uses. Called from constructor functions
    MEMBER PROCEDURE app_read_xlsx_constructor(
        -- internal use
        SELF IN OUT NOCOPY  app_read_xlsx_udt
        ,p_xlsx             BLOB DEFAULT NULL
        ,p_sheets           VARCHAR2 := NULL
        ,p_cell             VARCHAR2 := NULL
        ,p_ctx              NUMBER := NULL
    ) IS
    BEGIN
        IF p_ctx IS NULL THEN
            SELF.ctx := app_read_xlsx_pkg.create_ctx;
            app_read_xlsx_pkg.parse_blob(
                p_ctx       => SELF.ctx
                ,p_xlsx     => p_xlsx
                ,p_sheets   => p_sheets
                ,p_cell     => p_cell
            );
        ELSE
            SELF.ctx := p_ctx;
        END IF;
        SELECT string_val BULK COLLECT INTO SELF.col_names
        FROM as_read_xlsx_gtt t
        WHERE t.ctx = SELF.ctx AND t.row_nr = 1
        ORDER BY col_nr
        ;
        SELECT sheet_name INTO SELF.sheet_name
        FROM as_read_xlsx_gtt t
        WHERE t.ctx = SELF.ctx AND ROWNUM = 1
        ;
    END app_read_xlsx_constructor
    ;
    CONSTRUCTOR FUNCTION app_read_xlsx_udt(
        p_xlsx      BLOB
        ,p_sheets   VARCHAR2 := NULL
        ,p_cell     VARCHAR2 := NULL
    ) RETURN SELF AS RESULT
    IS
    BEGIN
        SELF.app_read_xlsx_constructor(
            p_xlsx      => p_xlsx
            ,p_sheets   => p_sheets
            ,p_cell     => p_cell
        );
        RETURN;
    END app_read_xlsx_udt
    ;
    CONSTRUCTOR FUNCTION app_read_xlsx_udt(
        p_ctx       NUMBER
    ) RETURN SELF AS RESULT
    IS
    BEGIN
        SELF.app_read_xlsx_constructor(
            p_ctx       => p_ctx
        );
        RETURN;
    END app_read_xlsx_udt
    ;


    MEMBER FUNCTION get_col_names RETURN &&d_arr_varchar2_udt.
    IS
    BEGIN
        RETURN col_names;
    END get_col_names
    ;

    MEMBER FUNCTION get_col_count RETURN NUMBER
    IS
    BEGIN
        RETURN col_names.COUNT;
    END get_col_count
    ;

    MEMBER FUNCTION get_ctx RETURN NUMBER
    IS
    BEGIN
        RETURN ctx;
    END get_ctx
    ;

    MEMBER FUNCTION get_sheet_name RETURN VARCHAR2
    IS
    BEGIN
        RETURN sheet_name;
    END get_sheet_name
    ;

    -- you can use this in a dynamic sql statement. It can become a WITH clause
    -- that you join to as needed. 
    --
    -- The query this outputs gives you 
    -- a column named data_row_nr with the line number in the input spreadsheet
    -- starting at 1 for the line after the column headers 
    -- (so row 2 becomes data_row_nr=1).
    --
    -- The rest of the columns are named as the column headers in the input
    -- spreadsheet with case preserved. Newlines in headers will likely blow it up.
    --
    MEMBER FUNCTION get_sql
    RETURN CLOB
    IS
        v_sql           CLOB ;
    BEGIN
        -- since this is dynamic sql that will be executed by a procedure
        -- that can be in another schema, and we have not created synonyms,
        -- fully qualify the name of the function
        v_sql := q'{SELECT X.R.data_row_nr AS data_row_nr, 
}'
            --
            -- The use of VALUE() function is to get the full object from the pipelined table
            -- function. By default you get the object members as columns. We want the object
            -- that includes the anydata collection because the object has the get() method we
            -- use to extract individual elements from the collection.
            --
            || SELF.get_col_sql('X.R')||'
  FROM (
    SELECT VALUE(t) AS R -- full object, not the object members * would provide
    FROM TABLE('||$$PLSQL_UNIT_OWNER||'.app_read_xlsx_udt.get_data_rows('
            ||TO_CHAR(SELF.ctx)
            ||','||TO_CHAR(SELF.get_col_count)
            || ')) t
  ) X';
DBMS_OUTPUT.put_line(v_sql);
        RETURN v_sql;
    END get_sql
    ;


    MEMBER FUNCTION get_col_sql(p_oname VARCHAR2 := 'X.R')
    RETURN CLOB
    IS
        v_sql   CLOB;
        v_comma CONSTANT VARCHAR2(8) := '
    ,';
    BEGIN
        v_sql := '   '||p_oname||'.get(1) AS "'||col_names(1)||'"';
        FOR i IN 2..col_names.COUNT
        LOOP
            v_sql := v_sql||v_comma||p_oname||'.get('||TO_CHAR(i)||') AS "'||col_names(i)||'"';
        END LOOP;
        RETURN v_sql;
    END get_col_sql
    ;

    -- called for you in the SQL provided by get_sql
    STATIC FUNCTION get_data_rows(
         p_ctx      NUMBER
        ,p_col_cnt  NUMBER
    )
    RETURN arr_app_read_xlsx_row_udt PIPELINED
    IS
        v_arr       arr_app_read_xlsx_row_udt;

        CURSOR c_filled_gaps IS
WITH cols AS (
    SELECT level AS col_nr FROM dual CONNECT BY level <= p_col_cnt
)
, this_ctx_cols AS (
    SELECT row_nr, col_nr, cell_type, string_val, date_val, number_val
    FROM as_read_xlsx_gtt
    WHERE ctx = p_ctx AND row_nr > 1 AND col_nr <= p_col_cnt
)
, ad_cols AS (
    SELECT t.row_nr, cols.col_nr
        ,CASE t.cell_type
            WHEN 'S' THEN SYS.ANYDATA.convertVarchar2(t.string_val)
            WHEN 'D' THEN SYS.ANYDATA.convertDate(t.date_val)
            WHEN 'N' THEN SYS.ANYDATA.convertNumber(t.number_val)
            ELSE SYS.ANYDATA.convertVarchar2(NULL) -- must have a placeholder for collect
        END AS ad
    FROM this_ctx_cols t
    PARTITION BY (t.row_nr) -- fill gaps for empty cells
    RIGHT OUTER JOIN cols
        ON cols.col_nr = t.col_nr
)
, ad_arr AS (
    SELECT row_nr - 1 AS data_row_nr
        ,CAST( COLLECT(ad ORDER BY col_nr) AS arr_anydata_udt) AS vaa
    FROM ad_cols
    GROUP BY row_nr
) 
SELECT app_read_xlsx_row_udt(data_row_nr, vaa)
FROM ad_arr
--ORDER BY data_row_nr
        ;

    BEGIN
        OPEN c_filled_gaps;
        LOOP
            FETCH c_filled_gaps BULK COLLECT INTO v_arr LIMIT 100;
            EXIT WHEN v_arr.COUNT = 0;
            FOR i IN 1..v_arr.COUNT
            LOOP
                PIPE ROW(v_arr(i));
            END LOOP;
        END LOOP;
        CLOSE c_filled_gaps;
        RETURN;
    END get_data_rows
    ;

END;
/
show errors
