spool install.log
set serveroutput on
whenever sqlerror exit failure
--
-- for conditional compilation based on sqlplus define settings.
-- When we select a column alias named "file_choice", we get a sqlplus define value for "file_choice" called "do_file"
--
COLUMN file_choice NEW_VALUE do_file NOPRINT
-- who do you want to grant execute to?
define GRANT_LIST="public"
--
-- People care about naming conventions. You must define this collection type name.
-- If you already have collection type named the way you like, define that name here
--
define d_arr_varchar2_udt="arr_varchar2_udt"
-- Set these to FALSE if you do not need to compile them
define compile_arr_varchar2_udt="TRUE"

SELECT DECODE('&&compile_arr_varchar2_udt','TRUE','arr_varchar2_udt.tps', 'do_nothing.sql arr_varchar2_udt') AS file_choice FROM dual;
prompt calling &&do_file
@@&&do_file
--
prompt compile anydata_shell_udt
@@anydata_shell_udt.tps
GRANT EXECUTE ON anydata_shell_udt TO &&GRANT_LIST ;
--
prompt compile arr_anydata_shell_udt
@@arr_anydata_shell_udt.tps
GRANT EXECUTE ON arr_anydata_shell_udt TO &&GRANT_LIST ;
--
prompt compile app_read_xlsx_row_udt
@@app_read_xlsx_row_udt.tps
@@app_read_xlsx_row_udt.tpb
GRANT EXECUTE ON app_read_xlsx_row_udt TO &&GRANT_LIST ;
--
prompt create global temporary table as_read_xlsx_gtt
@@as_read_xlsx_gtt.sql
GRANT SELECT, INSERT, DELETE, UPDATE ON as_read_xlsx_gtt TO &&GRANT_LIST ;
--
prompt compile package as_read_xlsx
@@as_read_xlsx/as_read_xlsx.pkg
-- only app_read_xlsx needs it
--GRANT EXECUTE ON as_read_xlsx TO &&GRANT_LIST ;
--
prompt compile package spec app_read_xlsx_pkg
@@app_read_xlsx_pkg.pks
prompt compile type spec app_read_xlsx_udt
@@app_read_xlsx_udt.tps
prompt compile package body app_read_xlsx_pkg
@@app_read_xlsx_pkg.pkb
prompt compile type body app_read_xlsx_udt
@@app_read_xlsx_udt.tpb
GRANT EXECUTE ON app_read_xlsx_udt TO &&GRANT_LIST ;
