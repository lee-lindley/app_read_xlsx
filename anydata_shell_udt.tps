CREATE OR REPLACE TYPE anydata_shell_udt FORCE AS OBJECT(
    -- purpose is to allow NULLS to hold a place in an array from COLLECT aggregate function
    ad  anydata
);
/
show errors
