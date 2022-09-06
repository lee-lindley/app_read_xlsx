DECLARE
    v_o             app_read_xlsx_udt;
    v_sql           VARCHAR2(32767);
    v_ctxId         ExcelGen.ctxHandle;
    v_sheetHandle   BINARY_INTEGER;
BEGIN
    v_o := app_read_xlsx_udt(to_blob(bfilename('TMP_DIR' ,'test_salaries.xlsx')), '1');
    v_sql := 'WITH a AS (
'||v_o.get_sql||q'{
)
, b AS (
    SELECT a.data_row_nr
        ,CASE SYS.ANYDATA.getTypeName("Emp #")
            WHEN 'SYS.NUMBER' THEN SYS.ANYDATA.accessNumber("Emp #")
        END AS employee_id
    FROM a
)
, c AS (
    SELECT b.data_row_nr
        ,e.salary AS "Old Salary"
        ,e.first_name||' '||e.last_name AS "Name"
    FROM b
    INNER JOIN hr.employees e
        ON e.employee_id = b.employee_id
    WHERE b.employee_id IS NOT NULL
)
SELECT a.*, c."Old Salary", c."Name"
FROM a
LEFT OUTER JOIN c
    ON c.data_row_nr = a.data_row_nr
}';

    DBMS_OUTPUT.put_line(v_sql);
    v_ctxId := ExcelGen.createContext();
    v_sheetHandle := ExcelGen.addSheetFromQuery(v_ctxId, v_o.get_sheet_name, v_sql, p_sheetIndex => 1);
    -- freeze the top row with the column headers
    ExcelGen.setHeader(v_ctxId, v_sheetHandle, p_frozen => TRUE);
    -- style with alternating colors on each row.
    ExcelGen.setTableFormat(v_ctxId, v_sheetHandle, 'TableStyleLight2');
    -- anydata type comes in as a number. Set format for this column to date. strings will be fine
    ExcelGen.setColumnFormat(v_ctxId, v_sheetHandle, 4, 'MM/DD/YYYY');
    ExcelGen.createFile(v_ctxId, 'TMP_DIR', 'test_salaries_output.xlsx');
    ExcelGen.closeContext(v_ctxId);
    --v_o.destructor;
END;
/

