This is the readme for XIRR_AGG, a PL/SQL function that provides similar functionality to the XIRR function in Microsoft Excel

Contents:
    1. License
    2. Minimum Requirements
    3. Installation and Usage
    
1. License
--------------------------------------------------------------------------------
This code is licensed using the Apache 2 License - a copy of this license is 
provided in the LICENSE.txt file.

2. Minimum Requirements
--------------------------------------------------------------------------------
At a minimum, to install and test XIRR_AGG you must have the following:

1.  An Oracle RDBMS, version 9i or higher.  (The code has only been tested on 11g
    and 12c)
2.  A database schema (user) with the following privileges:
    * CREATE TABLE
    * SELECT and INSERT
    * CREATE INDEX
    * CREATE TYPE
    * CREATE FUNCTION
    * ALTER TABLE
    * DROP TABLE
    * DROP INDEX
    * DROP TYPE
    
3. Installation and Usage
--------------------------------------------------------------------------------
To install, connect to the Oracle database and run the included script:
    
    xirr_t.sql
    
This script should create 2 VARRAY types and an object type specification
and body.  Finally, it will create the XIRR_AGG function.

It is recommended to run the following scripts, in the following order, to test
the installation:

    xirr_test_data.sql  <-- this will create a table and insert test data
    xirr_validate.sql   <-- this will use the test data to validate the function
                            is working correctly
                            
How to use XIRR_AGG
--------------------------

XIRR_AGG needs 2 columns, one with dates and one with the associated values (cash flows).  
Because user-defined aggregates can only accept one value, these must be concatenated 
together as a string, joined by a delimiter, which is assumed to be a tilde (~) by default.
Also by default the date is assumed to be in YYYYMMDD format.

Here is a sample query using the test table and data:

SELECT  test_name,
        XIRR_AGG(TO_CHAR(cash_flow_date, 'YYYYMMDD') || '~' || TO_CHAR(cash_flow_amount)) AS xirr
FROM    xirr_test
GROUP BY test_name;
                     

Debugging
--------------------------
The code was written with a method to turn on debugging statements without having to
recompile the object, the result is similar to passing a -v or verbose argument to a UNIX 
shell script, by passing a '1' in the column to be aggregated using a UNION ALL like so:

SET SERVEROUTPUT ON SIZE 1000000;
SELECT  test_name,
        XIRR_AGG(xirr_value) AS xirr
FROM    (SELECT 'XIRR TEST 1' AS test_name,
                '1' AS xirr_value
        FROM    DUAL
        UNION ALL
        SELECT  test_name,
                TO_CHAR(cash_flow_date, 'YYYYMMDD') || '~' || TO_CHAR(cash_flow_amount) AS xirr_value
        FROM    xirr_test)
GROUP BY test_name;

Just be sure to include values in the dummy columns that are also values in the real columns
           
    
    
   