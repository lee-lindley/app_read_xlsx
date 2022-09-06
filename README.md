# app_read_xlsx - An Oracle Object Type and Package for reading an XLSX worksheet

Anton Scheffer's *as_read_xlsx* is the goto tool for reading an Excel spreadsheet in PL/SQL.
*app_read_xlsx* extends *as_read_xlsx* providing a SQL
statement that presents the spreadsheet data as **ANYDATA** values with SQL column names
matching the column headers from the first row.

*as_read_xsx* is included in the distribution, but you can find the original source in 
a link named "here" from [this page](https://technology.amis.nl/languages/oracle-plsql/read-a-excel-xlsx-with-plsql/)
if you want to verify. The only change I made was to set invoker rights via an *AUTHID CURRENT_USER* package attribute.

A slight tweak to Marc Bleron's [ExcelGen](https://github.com/mbleron/ExcelGen) supports **ANYDATA** type columns
in the input cursor such that you can copy data, even messy data with varying data types within columns, from one spreadsheet
to another. That tweak lives in [a fork of ExcelGen](https://github.com/lee-lindley/ExcelGen/tree/anydata) at the moment.
Marc is refactoring *ExcelGen* so the pull request will not be accepted; however, we have a soft committment to consider support 
of **ANYDATA** columns in the future.

# Content

- [Installation](#installation)
- [Use Case](#use-case)
- [Examples](#examples)
- [Manual Page](#manual-page)

# Installation

Clone this repository or download it as a [zip](https://github.com/lee-lindley/app_read_xlsx/archive/refs/heads/main.zip)
archive.

The file *install.sql* has a few define values you may what to change:

- *GRANT_LIST* - The required set of grants are provided to schemas in this string. It can be multiple schemas comma separated. The default is **PUBLIC**.
- *d_arr_varchar2_udt* - The name to use for a User Defined Type defined as a TABLE OF VARCHAR2(4000). You may already have one. The default name is *arr_varchar2_udt*.
- *compile_arr_varchar2_udt* - The default of **TRUE** means we will create the type named by *d_arr_varcahr2_udt*. **FALSE** indicates you already have one and do not want to create or replace it.

Once you complete any changes to *install.sql*, run it with sqlplus:

`sqlplus YourLoginConnectionString @install.sql`

The script runners in Toad or SqlDeveloper should also work just fine, but I did not test them.

# Use Case

| ![Spreadsheet Input Use Case](images/spreadsheet_input_use_case.gif) |
|:--:|
| app_read_xlsx_udt Use Case Diagram |

Premise: Spreadsheet users do naughty things like put strings in the middle of date columns. A string like 'N/A' will invalidate
an assumption that the column contains either a valid date or NULL in all rows. You can convert everything to strings,
but when you load back into another spreadsheet, that is less than desirable.

The primary use case for creation of this tool is to ingest messy spreadsheet data, use the less messy parts
to join to other database tables to gather additional information, then faithfully reproduce the original
spreadsheet data in an output spreadsheet while adding one or more columns of supplemental data.

# Examples

The first example reads the spreadsheet data into a global temporary table, then prints out a SQL
statement that you can use during that session. In practice one would build a dynamic SQL statement
in a program using that string.

```sql
DECLARE
    v_o app_read_xlsx_udt;
BEGIN
    -- read sheet 1 of the spreadsheet found on the database server in a directory mapped to TMP_DIR
    v_o := app_read_xlsx_udt(TO_BLOB(BFILENAME('TMP_DIR', 'Book1.xlsx')), '1');
    DBMS_OUTPUT.put_line(v_o.get_sql);
END;
/
```

The second example will input a spreadsheet as pictured here. You can find the xlsx input and output files
in the test directory.

| ![test_salaries spreadsheet](images/test_salaries.gif) |
|:--:|
| test_salaries spreadsheet|

Notice the "Decision Date" column contains cells that are not dates. We read this data, treat the "Emp #" column
as a number (handle exceptions as you see fit) and join to **hr.employees** to retrieve the name and current salary.
We then output a new spreadsheet with the joined data elements together with the original spreadsheet content.


```sql
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
```

| ![test_salaries_output spreadsheet](images/test_salaries_output.gif) |
|:--:|
| test_salaries_output spreadsheet|

# Manual Page

## app_read_xlsx_udt constructor

Creates the object from the provided spreadsheet as a BLOB. The other two arguments are for *as_read_xlsx*. You will
generally provide the ordinal number for a sheet you want to read as a string. Although *as_read_xlsx* supports reading
more than one sheet, *app_read_xlsx_udt* does not.

Data from the spreadsheet is stored in a global temporary table named *as_read_xlsx_gtt* for the life of the session
(unless *destructor* method is called).

```sql
    CONSTRUCTOR FUNCTION app_read_xlsx_udt(
        p_xlsx      BLOB
        ,p_sheets   VARCHAR2 := NULL
        ,p_cell     VARCHAR2 := NULL
    ) RETURN SELF AS RESULT
```

## get_sql

Returns a string containing a SQL SELECT statement the columns of which have the names of the column headers in
the first row of the input spreadsheet. The column values are type **ANYDATA**. 

```sql
    MEMBER FUNCTION get_sql RETURN CLOB
```

## get_col_names

Returns a collection of type **arr_varchar2_udt** (unless you configured a different type name in the install.sql script)
consisting of the input spreadsheet column header values. Case is preserved. Beware that if you use these strings
to construct dynamic sql, you should protect them with double quotes. This list does not include the special column
*data_row_nr* that is added to the resultset provided by *get_sql*.

```sql
    MEMBER FUNCTION get_col_names RETURN arr_varchar2_udt
```

## get_col_count

Returns the number of columns from the input spreadsheet that have headers. Same as number of elements in the collection
returned by *get_col_names*.

```sql
    MEMBER FUNCTION get_col_count RETURN NUMBER
```

## get_ctx

Returns the context number that is a key value for selecting from the *as_read_xlsx_gtt* global temporary table. The
context number allows having multiple input spreadsheets in the same session.

```sql
    MEMBER FUNCTION get_ctx RETURN NUMBER
```

## get_sheet_name

Returns the name of the tab/sheet from the input spreadsheet.

```sql
    MEMBER FUNCTION get_sheet_name RETURN VARCHAR2
```

## get_col_sql

Returns the column select list used by *get_sql*. It is possible you could find a use for it, but it is not
part of mainstream usage.

```sql
    MEMBER FUNCTION get_col_sql(p_oname VARCHAR2 := 'X.R') RETURN CLOB
```

## get_data_rows

A pipelined table function that returns object type rows including a collection of **ANYDATA** values.
It provides the pivot of data into rows, densifies the missing pieces for empty cells, and allows
the selection of individual elements of the collection via a *get* method and an index. This is how *get_sql*
is able to provide dynamic SQL to extract individual column **ANYDATA** elements 
and provide column names from the input spreadsheet.

Alternatives were considered.
Polymorphic Table Functions do not support object types and the **ANYDATASET** design pattern is complex. We could
build an **ANYDATASET** implementation (see *ExcelTable.getRows* in [ExcelTable](https://github.com/mbleron/ExcelTable)),
but the number of people with the skill and willingness to support it is limited. I did not feel comfortable
encumbering my current employer with that liability. This method uses a resultset footprint that is known at compile
time, then object methods to build a column list in dynamic SQL at run time. It is still a bit complicated, but
should be in the wheelhouse of most journeyman Oracle developers.

```sql
    STATIC FUNCTION get_data_rows(
         p_ctx      NUMBER
        ,p_col_cnt  NUMBER
    ) RETURN arr_app_read_xlsx_row_udt PIPELINED
```

## destructor

PL/SQL objects do not automatically call a destructor method. Shame.

This method deletes rows from *as_read_xlsx_gtt* for this spreadsheet as identified by the context number. Removes
the context number from the collection maintained in a session level package global variable.

If you have a long running session that handles multiple spreadsheet inputs, this provides a way to
keep the memory and temporary table sizes from growing unbounded. You can also call it if you are a neat freak.
For most use cases a session will exist only long enough to process one or a few spreadsheets then exit, thus
automatically freeing the memory and space. You may never need this method.

```sql
    MEMBER PROCEDURE destructor
```
