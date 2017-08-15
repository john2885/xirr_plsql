--------------------------------------------------------------------------------
--
-- Copyright 2016 Russ Johnson <john_2885@yahoo.com>
--
-- Licensed under the Apache License, Version 2.0 (the "License");
-- you may not use this file except in compliance with the License.
-- You may obtain a copy of the License at
--
--     http://www.apache.org/licenses/LICENSE-2.0
--
-- Unless required by applicable law or agreed to in writing, software
-- distributed under the License is distributed on an "AS IS" BASIS,
-- WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
-- See the License for the specific language governing permissions and
-- limitations under the License.
--
--------------------------------------------------------------------------------

--------------------------------------------------------------------------------
-- xirr_validate.sql
--
-- A PL/SQL block to test that XIRR_AGG is working correctly based on the sample
-- set of test data provided in xirr_test_data.sql and XIRR values calculated by
-- Microsoft Excel
--------------------------------------------------------------------------------

SET SERVEROUTPUT ON SIZE 1000000;

DECLARE

    -- These are the Excel values rounded to 5 decimal places
    -- The XIRR TEST X should return a NULL because the IRR/NPV curve
    -- does not have a proper root (XIRR returns 0 in Excel for that test data)
    n_xirr_test1        NUMBER := -0.08825;
    n_xirr_test2        NUMBER := 0.02471;
    n_xirr_test3        NUMBER := 0.30006;
    n_xirr_test4        NUMBER := -0.04323;
    
    -- Assume that all tests will pass
    b_pass              BOOLEAN := TRUE;

    CURSOR c_xirr
    IS
    SELECT test_name,
           ROUND(XIRR_AGG(TO_CHAR(cash_flow_date, 'YYYYMMDD') || '~' || TO_CHAR(cash_flow_amount)), 5) AS xirr
    FROM   xirr_test
    GROUP BY test_name
    ORDER BY test_name;
    
BEGIN

    FOR r_xirr IN c_xirr
    LOOP
        DBMS_OUTPUT.PUT_LINE('Performing test: ' || r_xirr.test_name);
    
        IF UPPER(r_xirr.test_name) = 'XIRR TEST 1'
        THEN
            DBMS_OUTPUT.PUT_LINE('  Expected IRR: ' || TO_CHAR(n_xirr_test1));
            DBMS_OUTPUT.PUT_LINE('  Returned IRR: ' || TO_CHAR(r_xirr.xirr));
            
            IF n_xirr_test1 = r_xirr.xirr
            THEN
                DBMS_OUTPUT.PUT_LINE('PASSED');
                
            ELSE
                DBMS_OUTPUT.PUT_LINE('FAILED');
                b_pass := FALSE;              
            
            END IF;
            
            DBMS_OUTPUT.PUT_LINE('');
            
        ELSIF UPPER(r_xirr.test_name) = 'XIRR TEST 2'
        THEN
            DBMS_OUTPUT.PUT_LINE('  Expected IRR: ' || TO_CHAR(n_xirr_test2));
            DBMS_OUTPUT.PUT_LINE('  Returned IRR: ' || TO_CHAR(r_xirr.xirr));
            
            IF n_xirr_test2 = r_xirr.xirr
            THEN
                DBMS_OUTPUT.PUT_LINE('PASSED');
                
            ELSE
                DBMS_OUTPUT.PUT_LINE('FAILED');
                b_pass := FALSE;              
                  
            END IF;
            
            DBMS_OUTPUT.PUT_LINE('');
        
        ELSIF UPPER(r_xirr.test_name) = 'XIRR TEST 3'
        THEN
            DBMS_OUTPUT.PUT_LINE('  Expected IRR: ' || TO_CHAR(n_xirr_test3));
            DBMS_OUTPUT.PUT_LINE('  Returned IRR: ' || TO_CHAR(r_xirr.xirr));
            
            IF n_xirr_test3 = r_xirr.xirr
            THEN
                DBMS_OUTPUT.PUT_LINE('PASSED');
                
            ELSE
                DBMS_OUTPUT.PUT_LINE('FAILED');
                b_pass := FALSE;              
                  
            END IF;        
            
            DBMS_OUTPUT.PUT_LINE('');
        
        ELSIF UPPER(r_xirr.test_name) = 'XIRR TEST 4'
        THEN
            DBMS_OUTPUT.PUT_LINE('  Expected IRR: ' || TO_CHAR(n_xirr_test4));
            DBMS_OUTPUT.PUT_LINE('  Returned IRR: ' || TO_CHAR(r_xirr.xirr));
            
            IF n_xirr_test4 = r_xirr.xirr
            THEN
                DBMS_OUTPUT.PUT_LINE('PASSED');
                
            ELSE
                DBMS_OUTPUT.PUT_LINE('FAILED');
                b_pass := FALSE;              
                  
            END IF;        
            
            DBMS_OUTPUT.PUT_LINE('');
        
        ELSIF UPPER(r_xirr.test_name) = 'XIRR TEST X'
        THEN
            DBMS_OUTPUT.PUT_LINE('  Expected IRR: NULL');
            
            IF r_xirr.xirr IS NULL
            THEN
                DBMS_OUTPUT.PUT_LINE('  Returned IRR: NULL');
                DBMS_OUTPUT.PUT_LINE('PASSED');
                
            ELSE
                DBMS_OUTPUT.PUT_LINE('  Returned IRR: ' || TO_CHAR(r_xirr.xirr));
                DBMS_OUTPUT.PUT_LINE('FAILED');
                b_pass := FALSE;              
                  
            END IF;        
            
            DBMS_OUTPUT.PUT_LINE('');
        
        END IF;
        
    END LOOP;
        
    IF b_pass
    THEN
        DBMS_OUTPUT.PUT_LINE('All tests passed successfully, XIRR_AGG is installed and working properly');
        
    ELSE
        DBMS_OUTPUT.PUT_LINE('Some or all tests failed. Please review the output above and ensure the XIRR_AGG function and test data wast installed successfully.');
            
    END IF;
   
END;
/


    
    