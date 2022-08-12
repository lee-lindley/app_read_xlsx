spool install.log
-- set serveroutput on
whenever sqlerror exit failure
define GRANT_LIST="public"
--
prompt compile package as_read_xlsx
@@as_read_xlsx/as_read_xlsx.pkg
GRANT EXECUTE ON as_read_xlsx TO &&GRANT_LIST ;
--
prompt create global temporary table as_read_xlsx_gtt
@@as_read_xlsx_gtt.sql
GRANT SELECT, INSERT, UPDATE, DELETE ON as_read_xlsx_gtt TO &&GRANT_LIST ;
--
prompt compile package app_read_xlsx
@@app_read_xlsx.pks
@@app_read_xlsx.pkb
