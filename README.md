# app_read_xlsx - An Oracle PL/SQL package for reading an XLSX worksheet into a temporary table

Anton Scheffer's *as_read_xlsx* is the goto tool for reading an Excel spreadsheet in PL/SQL.


REST OF THIS IS TO BE DONE JUST A COPY PLACEHOLDER
# Content

- [Installation](#installation)
- [Use Cases](#use-cases)
    - [Generate CSV Rows](#generate-csv-rows)
    - [Create CSV FIle](#create-csv-file)

# Installation

Clone this repository or download it as a [zip](https://github.com/lee-lindley/app_csv_udt/archive/refs/heads/main.zip) archive.

Note: [plsql_utilties](https://github.com/lee-lindley/plsql_utilities) is provided as a submodule,
so use the clone command with recursive-submodules option:

`git clone --recursive-submodules https://github.com/lee-lindley/app_csv_udt.git`

or download it separately as a zip 
archive ([plsql_utilities.zip](https://github.com/lee-lindley/plsql_utilities/archive/refs/heads/main.zip)),
and extract the content of root folder into *plsql_utilities* folder.

## install.sql

If you already have a suitable TABLE type, you can update the sqlplus define variable *d_arr_varchar2_udt*
and set the define *compile_arr_varchar2_udt* to FALSE in the install file. You can change the name
of the type with *d_arr_varchar2_udt* and keep *compile_arr_varchar2_udt* as TRUE in which case
it will compile the appropriate type with your name.

Same can be done for *d_arr_integer_udt*, *d_arr_clob_udt* and *d_arr_arr_clob_udt*.

The User Defined Types *app_dbms_sql* and *app_dbms_sql_str* are required subcomponents from plsql_utilities. 
If you want to rename them you will need to edit the code and install scripts.

You could also edit *app_csv_udt.tps* and *app_csv_udt.tpb* to
change the return type of *get_rows* to TABLE of CLOB if you have a need for rows longer than 4000 chars 
in a TABLE function callable from SQL. If you are dealing exclusively in PL/SQL, the rows are already CLOB.

Once you complete any changes to *install.sql*, run it with sqlplus:

`sqlplus YourLoginConnectionString @install.sql`

# Use Cases

| ![app_csv_udt Use Case Diagram](images/app_csv_use_case.gif) |
|:--:|
| app_csv_udt Use Case Diagram |

## Generate CSV Rows

You can use a simple SQL SELECT to read CSV strings as records from the TABLE function *get_rows*, perhaps
spooling them to a text file with sqlplus. Given how frequently I've seen a cobbled together
SELECT concatenting multiple fields and separators into one value, this may
be the most common use case.

