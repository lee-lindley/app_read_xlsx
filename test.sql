declare
    v_o app_read_xlsx_udt;
    --v_src SYS_REFCURSOR;
    v_sql CLOB;
    --v_blob          BLOB;
    v_ctxId         ExcelGen.ctxHandle;
    v_sheetHandle   BINARY_INTEGER;
begin
    v_o := app_read_xlsx_udt(to_blob(bfilename('TMP_DIR' ,'Book1.xlsx')), '1');
    v_sql := v_o.get_sql||' SELECT * FROM app_read_xlsx_sql';
    dbms_output.put_line(v_sql);
        v_ctxId := ExcelGen.createContext();
        v_sheetHandle := ExcelGen.addSheetFromQuery(v_ctxId, 'app_read_xlsx demo', v_sql, p_sheetIndex => 1);
        -- freeze the top row with the column headers
        ExcelGen.setHeader(v_ctxId, v_sheetHandle, p_frozen => TRUE);
        -- style with alternating colors on each row. 
        ExcelGen.setTableFormat(v_ctxId, v_sheetHandle, 'TableStyleLight2');
        ExcelGen.createFile(v_ctxId, 'TMP_DIR', 'app_read_xlsx_demo.xlsx');
        ExcelGen.closeContext(v_ctxId);
    v_o.destructor;
end;
/

select to_blob(bfilename('TMP_DIR','app_read_xlsx_demo.xlsx')) from dual;