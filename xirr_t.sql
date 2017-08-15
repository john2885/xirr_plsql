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
-- These array types are required by the XIRR_T object type
--------------------------------------------------------------------------------

CREATE OR REPLACE TYPE T_DATE_ARRAY IS VARRAY(1000000) OF DATE;
/

CREATE OR REPLACE TYPE T_AMOUNT_ARRAY IS VARRAY(1000000) OF NUMBER;
/

--------------------------------------------------------------------------------
-- The main XIRR_T object specification
--------------------------------------------------------------------------------

CREATE OR REPLACE TYPE XIRR_T AS OBJECT (

    --------------------------------------------------------------------------------
    -- Value count - used internally to track the number of values processed
    --------------------------------------------------------------------------------
    n_value_count   NUMBER,

    --------------------------------------------------------------------------------
    -- Delimiter - the character (or string) that will separate a date value
    --             and an amount value   
    --------------------------------------------------------------------------------
    v_delim         VARCHAR2(50),

    --------------------------------------------------------------------------------
    -- Internal rate of return (IRR) - this will store the IRR value that is
    --          eventually returned by the main function
    --------------------------------------------------------------------------------
    n_irr           NUMBER,

    --------------------------------------------------------------------------------
    -- Investment dates - an array of dates for which we have investment amounts
    --------------------------------------------------------------------------------
    t_dates         T_DATE_ARRAY,

    --------------------------------------------------------------------------------
    --  Amounts - an array of investment amounts that correspond to the investment dates
    --------------------------------------------------------------------------------
    t_amounts       T_AMOUNT_ARRAY,
    
    --------------------------------------------------------------------------------
    --  Debug - a flag to indicate if debug messages should be printed
    --          using DBMS_OUTPUT.
    --          0 = do not print debug messages
    --          1 = print debug messages
    --------------------------------------------------------------------------------
    n_debug         NUMBER,

    --------------------------------------------------------------------------------
    --  Function: CONSTRUCTOR
    --  
    --  Purpose: Provide the ability to instantiate this object without passing any values into the
    --           constructor.  This will set object values to their defaults and initialize any PL/SQL
    --           arrays.
    --
    --           v_delim    will be set to '~' (tilde)
    --           t_dates    will be initialized
    --           t_amounts  will be initialized
    --           n_irr      will be set to 0.0
    --           n_debug    will be set to 0 (meaning debug is OFF)    
    --
    --------------------------------------------------------------------------------    
    CONSTRUCTOR FUNCTION xirr_t RETURN SELF AS RESULT,

    ---------------------------------------------------------------------------------
    --  Procedure: DEBUG
    --
    --  Purpose: Print debugging messages to stdout using DBMS_OUTPUT
    --
    --  PARAMETER      TYPE    DESCRIPTION
    --  ------------   ------- ------------------------------------------------------
    --  SELF           IN OUT  Explicit reference to this object
    --  p_message      IN      Debugging message to print
    --
    ---------------------------------------------------------------------------------
    MEMBER PROCEDURE DEBUG(SELF         IN OUT XIRR_T,
                           p_message    IN     VARCHAR2),        

    ---------------------------------------------------------------------------------
    --  Procedure: NPV
    --
    --  Purpose: Calculate NPV using the internal dates and amounts arrays,
    --           given an IRR
    --
    --  PARAMETER      TYPE    DESCRIPTION
    --  ------------   ------- ------------------------------------------------------
    --  SELF           IN OUT  Explicit reference to this object
    --  p_rate         IN      The rate to use for NPV calculations
    --
    --  Returns: the NPV value
    --
    ---------------------------------------------------------------------------------                           
    MEMBER FUNCTION NPV(SELF            IN OUT XIRR_T,
                        p_rate          IN     NUMBER) 
        RETURN NUMBER,

    ---------------------------------------------------------------------------------
    --  Procedure: FIND_IRR_RANGE
    --
    --  Purpose: Find a "wide" upper and lower IRR bound such that between the 2 IRR
    --           values there exists an IRR (where NPV is 0).
    --
    --  PARAMETER      TYPE    DESCRIPTION
    --  ------------   ------- ------------------------------------------------------
    --  SELF           IN OUT  Explicit reference to this object
    --  p_guess_limit  IN      Limit the number of guesses for the range
    --  p_irr_low      IN OUT  Initial lower bound for IRR
    --  p_irr_high     IN OUT  Initial upper bound for IRR
    --  p_npv_low      OUT     Once range is found, this is the NPV for the lower IRR
    --  p_npv_high     OUT     Once range is found, this is the NPV for the upper IRR
    --
    ---------------------------------------------------------------------------------        
    MEMBER PROCEDURE FIND_IRR_RANGE(SELF           IN OUT XIRR_T,
                                    p_guess_limit  IN     NUMBER,
                                    p_irr_low      IN OUT NUMBER,
                                    p_irr_high     IN OUT NUMBER,
                                    p_npv_low      OUT    NUMBER,
                                    p_npv_high     OUT    NUMBER),
                                    
    ---------------------------------------------------------------------------------
    --  Function: IRR
    --
    --  Purpose: assuming we have found a "wide" upper and lower bound, narrow the 
    --           bounds until we arrive at an NPV close to 0, in which one of the 
    --           bounds will be the IRR - this should be called after FIND_IRR_RANGE
    --
    --  PARAMETER      TYPE    DESCRIPTION
    --  ------------   ------- ------------------------------------------------------
    --  SELF           IN OUT  Explicit reference to this object
    --  p_guess_limit  IN      Limit the number of guesses for the range
    --  p_irr_low      IN      Initial lower bound for IRR
    --  p_irr_high     IN      Initial upper bound for IRR
    --  p_npv_low      OUT     Initial NPV for the lower IRR
    --  p_npv_high     OUT     Initial NPV for the upper IRR
    --
    -- Returns: the IRR for the values in the internal date and amount arrays
    --
    ---------------------------------------------------------------------------------    
    MEMBER FUNCTION IRR(SELF            IN OUT XIRR_T,
                        p_guess_limit   IN     NUMBER,
                        p_irr_low       IN     NUMBER,
                        p_irr_high      IN     NUMBER,
                        p_npv_low       IN     NUMBER,
                        p_npv_high      IN     NUMBER)
        RETURN NUMBER,   
        
    --------------------------------------------------------------------------------
    --  Function: ODCIAggregateInitialize
    --  
    --  Purpose: This static function is inherited from the Oracle User Defined Aggregate API
    --           it instantiates the class by calling the constructor and passing back the instance
    --           of the class as the context, then returns a success indicator
    --
    --  Params:
    --          p_ctx - the context, datatype is the same as this object and it is passed by reference
    --          
    --
    --------------------------------------------------------------------------------    
    STATIC FUNCTION ODCIAggregateInitialize(p_ctx   IN OUT XIRR_T) 
        RETURN NUMBER,
    
    --------------------------------------------------------------------------------
    --  Function: ODCIAggregateIterate
    --  
    --  Purpose: This function is called on every member/row in the group
    --
    --  Params:
    --          SELF    - this object and it is passed by reference
    --          p_value - the column/value that this custom aggregate function operates on 
    --
    --------------------------------------------------------------------------------          
    MEMBER FUNCTION ODCIAggregateIterate(SELF       IN OUT  XIRR_T, 
                                         p_value    IN      VARCHAR2) 
        RETURN NUMBER,

    --------------------------------------------------------------------------------
    --  Function: ODCIAggregateTerminate
    --  
    --  Purpose: This function is called at the end of the aggregate function, after
    --           all rows/values have been processed.  It does any final calculations
    --           and returns the final value
    --
    --  Params:
    --          SELF         - this object and it is passed by reference
    --          p_return_val - the value calculated and returned by this aggregate function 
    --          p_flags      - A bit vector that indicates various options. A set bit of ODCI_AGGREGATE_REUSE_CTX 
    --                         indicates that the context will be reused and any external context should not be freed
    --
    --------------------------------------------------------------------------------         
    MEMBER FUNCTION ODCIAggregateTerminate(SELF         IN OUT  XIRR_T, 
                                           p_return_val OUT     NUMBER, 
                                           p_flags      IN      NUMBER) 
        RETURN  NUMBER,

    --------------------------------------------------------------------------------
    --  Function: ODCIAggregateMerge
    --  
    --  Purpose: The ODCIAggregateMerge function is invoked by Oracle to merge two aggregation contexts into a single object 
    --           instance. Two aggregation contexts may need to be merged during either serial or parallel evaluation of the 
    --           user-defined aggregate. This function takes the two aggregation contexts as input, merges them, and returns 
    --           the single, merged instance of the aggregation context.
    --
    --  Params:
    --          SELF         - this object and it is passed by reference
    --          p_ctx2       - the value of the other aggregation context
    --
    --------------------------------------------------------------------------------         
    MEMBER FUNCTION ODCIAggregateMerge(SELF         IN OUT XIRR_T, 
                                       p_ctx2       IN     XIRR_T) 
        RETURN  NUMBER
        
);
/

--------------------------------------------------------------------------------
-- The main XIRR_T object body
--------------------------------------------------------------------------------

CREATE OR REPLACE TYPE BODY XIRR_T 
IS

    --------------------------------------------------------------------------------
    --  Function: CONSTRUCTOR
    --------------------------------------------------------------------------------
    CONSTRUCTOR FUNCTION XIRR_T
        RETURN SELF AS RESULT
    IS
    
    BEGIN
        SELF.n_value_count := 0;
        SELF.v_delim    := '~';
        SELF.t_dates    := T_DATE_ARRAY();
        SELF.t_amounts  := T_AMOUNT_ARRAY();
        SELF.n_irr      := 0.0;
        SELF.n_debug    := 0;
        
        RETURN;
    END;

    --------------------------------------------------------------------------------
    --  Procedure: DEBUG
    --------------------------------------------------------------------------------    
    MEMBER PROCEDURE DEBUG(SELF      IN OUT XIRR_T,
                           p_message IN VARCHAR2)
    IS
    
    BEGIN
        IF SELF.n_debug = 1
        THEN
            DBMS_OUTPUT.PUT_LINE(p_message);
        END IF;
                
    END;    

    --------------------------------------------------------------------------------
    --  Function: NPV
    --------------------------------------------------------------------------------    
    MEMBER FUNCTION NPV(SELF    IN OUT  XIRR_T,
                        p_rate  IN      NUMBER)
        RETURN NUMBER
    IS
    
        n_npv   NUMBER := 0;
    
    BEGIN
    
        SELF.DEBUG('*** In member function NPV ***');
        SELF.DEBUG('Rate: ' || TO_CHAR(p_rate));

        FOR i IN 1..SELF.t_amounts.LAST
        LOOP
            n_npv := n_npv + (SELF.t_amounts(i) / POWER((1 + p_rate), (SELF.t_dates(i) - SELF.t_dates(1)) / 365));
            
        END LOOP;
        
        SELF.DEBUG('NPV: ' || TO_CHAR(n_npv));
        
        RETURN n_npv;
    
    EXCEPTION
    WHEN OTHERS
    THEN
        RAISE;
        
    END;
    
    --------------------------------------------------------------------------------
    --  Procedure: FIND_IRR_RANGE
    --------------------------------------------------------------------------------
    MEMBER PROCEDURE FIND_IRR_RANGE(SELF           IN OUT XIRR_T,
                                    p_guess_limit  IN     NUMBER,
                                    p_irr_low      IN OUT NUMBER,
                                    p_irr_high     IN OUT NUMBER,
                                    p_npv_low      OUT    NUMBER,
                                    p_npv_high     OUT    NUMBER)
    IS
        b_range_found   BOOLEAN := FALSE;
        b_move_right    BOOLEAN := FALSE;
        b_move_left     BOOLEAN := FALSE;
        n_guesses       NUMBER  := 1;
        n_npv_low       NUMBER;
        n_npv_high      NUMBER;
        n_guess_low     NUMBER := p_irr_low;
        n_guess_high    NUMBER := p_irr_high;
    
    BEGIN

        -----------------------------------------------------------------------------
        --
        -- Main loop
        --
        -- 1. Start with the low and high IRR guesses passed to this procedure
        -- 2. Calculate the NPV at the high and the low
        -- 3. If one is negative and the other positive then we found our range, loop ends
        -- 4. If both are positive or both are negative then test to see which is closer to 0 by using absolute value
        -- 5. If the low IRR value has an NPV further away from 0 (higher ABS value) then we need to "move" right by
        --    using higher values for our guesses
        -- 6. If the low IRR value has an NPV closer to 0 (lower ABS value) then we need to "move" left by using
        --    lower values for our guesses
        --
        -----------------------------------------------------------------------------
    
        n_npv_high := SELF.NPV(n_guess_high);
        n_npv_low  := SELF.NPV(n_guess_low);
    
        WHILE b_range_found = FALSE
        LOOP
            IF n_guesses >= p_guess_limit
            THEN
                -- Too many guesses, it appears there is no NPV of 0 for this set of values
                RAISE_APPLICATION_ERROR(-20999, 'IRR RANGE NOT FOUND');

            ELSE
                SELF.DEBUG('Guess: ' || TO_CHAR(n_guesses));
                SELF.DEBUG('Guess Low: ' || TO_CHAR(n_guess_low));
                SELF.DEBUG('Guess High: ' || TO_CHAR(n_guess_high));
                
                ---------------------------------------------------------------------
                -- 2017-02-28 (RAJ) Performance improvement
                -- Only calculate NPV for what we need.  If we moved right then we
                -- already have a low value (the previous high value) so we only
                -- need to calculate a new high NPV.  If we moved left then we 
                -- already have a high value (the previous low value) so we only
                -- need to calculate a new low NPV.
                ---------------------------------------------------------------------
                IF b_move_right
                THEN
                    n_npv_high := SELF.NPV(n_guess_high);
                    
                ELSIF b_move_left
                THEN
                    n_npv_low  := SELF.NPV(n_guess_low);
                    
                ELSE
                    NULL;
                    
                END IF;

                SELF.DEBUG('NPV Low: ' || TO_CHAR(n_npv_low));
                SELF.DEBUG('NPV High: ' || TO_CHAR(n_npv_high));
                
                ---------------------------------------------------------------------
                -- Here's the test to see if we've found the right range
                -- which will be where the "low" NPV is positive and the
                -- "high" NPV is negative so multiplying their signs
                -- will be negative
                ---------------------------------------------------------------------
                IF (SIGN(n_npv_high) * SIGN(n_npv_low)) <= 0
                THEN
                    SELF.DEBUG('IRR range found');
                    b_range_found  := TRUE;
                    
                ELSE
                    ---------------------------------------------------------------------
                    -- 2016-11-07 (RAJ) Fixing bug with moving the range the wrong way
                    -- Now the code assumes that the NPV curve will cross the x axis at some point (otherwise it will
                    -- exhaust itself and give up after n_range_guess_limit is reached).  After calculating the NPV for
                    -- both high and low, if we are in this section of code we know that they are both the same sign: 
                    -- either positive (above the x axis) or negative (below the x axis).  If the High NPV is closer to 0
                    -- (absolute value of n_npv_high is less than the absolute value of n_npv_low) 
                    -- then we need to move to the "right" with higher IRR values.  If the Low NPV is closer to 0 then
                    -- we need to move left with lower IRR values
                    ---------------------------------------------------------------------
                    IF ABS(n_npv_low) > ABS(n_npv_high) 
                    THEN
                        SELF.DEBUG(' moving right...');
                        n_guess_low     := n_guess_high;
                        n_guess_high    := n_guess_high + .5;
                        n_npv_low       := n_npv_high;
                        b_move_right    := TRUE;
                        b_move_left     := FALSE;
                    
                    ELSE
                        SELF.DEBUG(' moving left...');
                        n_guess_high    := n_guess_low;
                        n_guess_low     := n_guess_low - .5;
                        n_npv_high      := n_npv_low;
                        b_move_right    := FALSE;
                        b_move_left     := TRUE;
                        
                    END IF;
                    
                END IF;
                
                n_guesses := n_guesses + 1;
                
            END IF;
            
        END LOOP;

        -- Set our OUT parameter values
        p_irr_low  := n_guess_low;
        p_irr_high := n_guess_high;
        p_npv_low  := n_npv_low;
        p_npv_high := n_npv_high;
    
    EXCEPTION
    WHEN OTHERS
    THEN
        RAISE;
        
    END;

    --------------------------------------------------------------------------------
    --  Function: IRR
    --------------------------------------------------------------------------------    
    MEMBER FUNCTION IRR(SELF            IN OUT XIRR_T, 
                        p_guess_limit   IN NUMBER,
                        p_irr_low       IN NUMBER,
                        p_irr_high      IN NUMBER,
                        p_npv_low       IN NUMBER,
                        p_npv_high      IN NUMBER)
        RETURN NUMBER

    IS
        n_npv          NUMBER := -1;
        n_guesses      NUMBER := 1;
        n_irr          NUMBER;
        n_irr_low      NUMBER := p_irr_low;
        n_irr_high     NUMBER := p_irr_high;
        n_npv_low      NUMBER := p_npv_low;
        n_npv_high     NUMBER := p_npv_high;
        n_temp_low     NUMBER;
        n_npv_prev_low NUMBER;
      
    BEGIN

        ---------------------------------------------------------------------
        --
        -- Main loop
        --
        --  1. Check the NPVs of the "new" lower and upper bounds - it is optimized to not
        --     recalculate values that have already been calculated.  If we have one that is close
        --     to 0 (within 5 decimal places by default) then we are done - return the IRR
        --  2. Bisect the bounded range by moving the lower bound halfway to the upper bound
        --  3. If the new lower and upper bounds yield a positive and negative NPV then we
        --     have narrowed our range, start again at #1 with our new bounds
        --  4. If the new NPVs are both negative then we've moved too far, the IRR must exist
        --     between the previous lower bound and the new lower bound, set the bounds to
        --     those values and start again with #1
        --
        --  Example: start with a IRR bounds of -.5 and .5
        --
        --  STEP  IRR LOWER   IRR UPPER    NPV LOWER     NPV UPPER
        --  ----- ----------- ------------ ------------- --------------
        --  1     -.5         .5           20,000        -20,000 <-- Move halfway to the upper bound in step 2
        --  2     0           .5           -10,000       -20,000 <-- Wrong half, use the other half in step 3
        --  3     -.5         0            20,000        -10,000 <-- Move halfway to the upper bound in step 4
        --  4     -.25        0            10,000        -10,000 <-- Move halfway to the upper bound in step 5
        --  5     -.125       0            -5,000        -10,000 <-- Wrong half, use the other half in step 6
        --  6     -.25        -.125        10,000        -5,000  <-- Move halfway to the upper bound in step 7
        --  7     -.1825      -.125        5,000         -5,000  <-- Move halfway to the upper bound in step 8
        --  8     -.15375     -.125        0             -5,000  <-- IRR found, it is equal to the lower bound
        --
        --  Return this IRR: -.15375 or -15.38%
        --        
        ---------------------------------------------------------------------
        WHILE  n_npv <> 0 AND n_guesses < p_guess_limit
        LOOP
        
            SELF.DEBUG('Guess Low: ' || TO_CHAR(n_irr_low));
            SELF.DEBUG('Guess High: ' || TO_CHAR(n_irr_high));
            SELF.DEBUG('NPV Low: ' || TO_CHAR(n_npv_low));
            SELF.DEBUG('NPV High: ' || TO_CHAR(n_npv_high));
            
            
            IF (ROUND(n_npv_low * 100000) = 0)
            THEN
                -- Great, we have our IRR
                SELF.DEBUG('IRR found: ' || TO_CHAR(n_irr_low));
                n_npv := 0;
                n_irr := n_irr_low;
            
            ELSIF (ROUND(n_npv_high * 100000) = 0)
            THEN
                -- Great, we have our IRR
                SELF.DEBUG('IRR found: ' || TO_CHAR(n_irr_high));
                n_npv := 0;
                n_irr := n_irr_high;
                
            ELSE
                -- Need to bisect our range and test to see
                -- which "side" of the bisection we need to
                -- continue with
                IF (n_npv_high < 0 AND n_npv_low < 0) OR (n_npv_high > 0 AND n_npv_low > 0)
                THEN
                    -- We need to move back to the left, we've moved too far right
                    n_temp_low  := n_irr_low + (n_irr_low - n_irr_high);
                    n_irr_high  := n_irr_low;
                    n_irr_low   := n_temp_low;
                    
                    -- Since we already have NPV for the low value and the previous
                    -- low value, we can save some processing time by not recalculating
                    -- NPV for the bounds
                    n_npv_high    := n_npv_low;
                    n_npv_low     := n_npv_prev_low;
                    
                ELSE
                    -- We have moved in the correct direction, we have a narrower range
                    -- of IRR values and between them is an NPV of 0
                    -- Keep track of the current low NPV to save processing time later
                    -- and bisect again, then "move" to the right half
                    n_npv_prev_low := n_npv_low;
                    n_irr_low      := (n_irr_low + n_irr_high) / 2;
                    
                    -- Since we've "moved" we need to calculate the new "low" NPV
                    n_npv_low := SELF.NPV(n_irr_low);
                    
                END IF;
                
            END IF;

            n_guesses := n_guesses + 1;
        
        END LOOP;
        
        RETURN n_irr;

    
    EXCEPTION
    WHEN OTHERS
    THEN
        RAISE;
        
    END;

    --------------------------------------------------------------------------------
    --  Function: ODCIAggregateInitialize
    --  
    --  Purpose: This static function is inherited from the Oracle User Defined Aggregate API.
    --           It instantiates the class by calling the constructor and passing back the instance
    --           of the class as the context, then returns a success indicator
    --
    --  Params:
    --          p_ctx - the context, datatype is the same as this object and it is passed by reference
    --          
    --
    --------------------------------------------------------------------------------
    STATIC FUNCTION ODCIAggregateInitialize(p_ctx IN OUT XIRR_T) 
        RETURN NUMBER
    IS
    begin 
        p_ctx := XIRR_T;
        RETURN ODCIConst.Success;
    END;

    --------------------------------------------------------------------------------
    --  Function: ODCIAggregateIterate
    --  
    --  Purpose: This function takes in a separated date, amount pair as a string (VARCHAR2).  By default
    --           they should be passed in with the date first, in YYYYMMDD format, separated with a tilde '~'
    --           and then the amount (positive or negative).  This function will add the date and amount
    --           to the internal t_dates and t_amounts arrays as well as keep them ordered (by date).
    --
    --           2016-12-15     Added the ability to dynamically set the debug flag if a special value
    --                          is included in the query, usually by a UNION
    --
    --           For example, consider this data set:
    --
    --              '20161001~6000'
    --              '20160701~5000'
    --              '20160901~3000'
    --              '20160801~-1000'
    --              '20160601~-150000'
    --              '20161101~2500'
    --              '20161201~4500'
    --
    --           After all rows are processed the arrays will look like this: 
    --          
    --           Index      t_dates         t_amounts
    --           ---------- --------------- ---------
    --           1          2016-06-01        -150000
    --           2          2016-07-01           5000
    --           3          2016-08-01          -1000
    --           4          2016-09-01           3000
    --           5          2016-10-01           6000   
    --           6          2016-11-01           2500   
    --           7          2016-12-01           4500   
    --
    --------------------------------------------------------------------------------
    MEMBER FUNCTION ODCIAggregateIterate(SELF       IN OUT  XIRR_T, 
                                         p_value    IN      VARCHAR2) 
        RETURN NUMBER
    IS
        d_date      DATE;
        n_amount    NUMBER;
        n_log       NUMBER;
        
    BEGIN
        
        -- Check to see if we should override the debug value
        IF p_value = '1'
        THEN
            SELF.n_debug := TO_NUMBER(p_value);
            
        ELSIF INSTR(p_value, SELF.v_delim) > 0
        THEN
        
            SELF.n_value_count := SELF.n_value_count + 1;
            
            d_date      := TO_DATE(SUBSTR(p_value, 1, INSTR(p_value, SELF.v_delim)-1), 'YYYYMMDD');
            n_amount    := TO_NUMBER(SUBSTR(p_value, INSTR(p_value, SELF.v_delim)+1));
            
            IF SELF.t_dates.COUNT = 0
            THEN
                -- This is the very first record, add it to each array
                SELF.DEBUG('First record: ' || TO_CHAR(d_date, 'YYYY-MM-DD') || ' ' || TO_CHAR(n_amount));
                SELF.t_dates.EXTEND;
                SELF.t_amounts.EXTEND;        
                SELF.t_dates(1)      := d_date;
                SELF.t_amounts(1)    := n_amount;
                
            ELSIF d_date > SELF.t_dates(SELF.t_dates.LAST) 
            THEN
                -- This should be the last record, add it to the end of each array
                SELF.DEBUG('Adding as last record: ' || TO_CHAR(d_date, 'YYYY-MM-DD') || ' ' || TO_CHAR(n_amount));
                SELF.t_dates.EXTEND;
                SELF.t_amounts.EXTEND;                
                SELF.t_dates(SELF.t_dates.LAST)       := d_date;
                SELF.t_amounts(SELF.t_dates.LAST)     := n_amount;         
                
            ELSE
                FOR j IN SELF.t_dates.FIRST..SELF.t_dates.LAST
                LOOP
                    IF d_date <= SELF.t_dates(j)
                    THEN
                       ----------------------------------------------------------------------------
                        -- We are somewhere in the middle, the date value is 
                        -- greater than the previous array element but less than the current
                        -- array element, like this:
                        -- t_dates(11)  :   2016-03-01
                        -- t_dates(12)  :   2016-04-01
                        -- t_dates(13)  :   2016-05-01
                        -- t_dates(14)  :   2016-07-01  <-- 2016-06-01 belongs here
                        -- t_dates(15)  :   2016-08-01
                        -- 
                        -- So we need to put the date value at the current index
                        -- in the array and shift everything "down" like this:
                        -- t_dates(11)  :   2016-03-01
                        -- t_dates(12)  :   2016-04-01
                        -- t_dates(13)  :   2016-05-01
                        -- t_dates(14)  :   2016-06-01
                        -- t_dates(15)  :   2016-07-01  <-- this index is increased by 1 
                        -- t_dates(16)  :   2016-08-01  <-- this index is increased by 1
                        --
                        -- We also need to keep the amount in sync with the date index
                        -- so we do the same thing to the t_amounts array that we do with
                        -- the t_dates array
                        --
                        -- Update: original code kept duplicate dates like this:
                        -- t_dates(14)  :   2016-06-01      t_amounts(14) := 1000 
                        -- t_dates(15)  :   2016-06-01      t_amounts(15) := 5000
                        -- t_dates(16)  :   2016-07-01      t_amounts(15) := 900
                        --
                        -- Now, the code keeps dates unique and sums the amounts like this:
                        -- t_dates(14)  :   2016-06-01      t_amounts(14) := 6000
                        -- t_dates(15)  :   2016-07-01      t_amounts(15) := 900
                        --
                        -- This keeps the arrays small which allows the NPV and IRR calculations
                        -- to complete
                        ----------------------------------------------------------------------------
                        
                        IF d_date = SELF.t_dates(j)
                        THEN
                            SELF.DEBUG('Date ' || TO_CHAR(d_date, 'YYYY-MM-DD') || ' already has a value: ' || TO_CHAR(SELF.t_amounts(j)));
                            SELF.DEBUG('    adding this value to it: ' || TO_CHAR(n_amount));
                            
                            SELF.t_amounts(j) := SELF.t_amounts(j) + n_amount;
                            
                            SELF.DEBUG('New record value: ' || TO_CHAR(d_date, 'YYYY-MM-DD') || ' ' || TO_CHAR(n_amount));
                        ELSE
                        
                            SELF.DEBUG('Inserting this record at index '|| TO_CHAR(j) ||': ' || TO_CHAR(d_date, 'YYYY-MM-DD') || ' ' || TO_CHAR(n_amount));
                            
                            IF j > 1
                            THEN
                                SELF.DEBUG('    between ' || TO_CHAR(SELF.t_dates(j-1), 'YYYY-MM-DD') || ' at index ' || TO_CHAR(j-1) || ' and ' || TO_CHAR(SELF.t_dates(j), 'YYYY-MM-DD') || ' currently at index ' || TO_CHAR(j));
                            ELSE
                                SELF.DEBUG('    before current first record: ' || TO_CHAR(SELF.t_dates(j), 'YYYY-MM-DD'));
                            END IF;
                            
                            SELF.t_dates.EXTEND;
                            SELF.t_amounts.EXTEND;

                            FOR k IN REVERSE (j+1)..SELF.t_dates.LAST
                            LOOP
                                SELF.t_dates(k)       := SELF.t_dates(k-1);
                                SELF.t_amounts(k)     := SELF.t_amounts(k-1);
                                
                            END LOOP;
                            
                            SELF.t_dates(j)     := d_date;
                            SELF.t_amounts(j)   := n_amount;
                            
                        END IF;
                        
                        EXIT;
                        
                    END IF;
                
                END LOOP;
                
            END IF;

        END IF;
        
        RETURN ODCIConst.Success;
    END;

    --------------------------------------------------------------------------------
    --  Function: ODCIAggregateTerminate
    --  
    --  Purpose: This function calculates the IRR and returns it using the bisection method.
    --           
    --           
    --

    --
    --------------------------------------------------------------------------------
    MEMBER FUNCTION ODCIAggregateTerminate(SELF         IN OUT  XIRR_T, 
                                           p_return_val OUT     NUMBER, 
                                           p_flags      IN      NUMBER) 
        RETURN  NUMBER 
    IS
    
        n_irr_guess_limit       NUMBER  := 1000;
        n_range_guess_limit     NUMBER  := 50;
        n_irr_low               NUMBER  := -.5;
        n_irr_high              NUMBER  := .5;
        n_npv_low               NUMBER;
        n_npv_high              NUMBER;
    
    BEGIN
        
        SELF.DEBUG('Values Processed: ' || TO_CHAR(SELF.n_value_count));
        SELF.DEBUG('Date Array Size: ' || TO_CHAR(SELF.t_dates.COUNT));
        SELF.DEBUG('Amount Array Size: ' || TO_CHAR(SELF.t_amounts.COUNT));       
       
        IF SELF.n_debug > 0
        THEN
            FOR i IN SELF.t_dates.FIRST..SELF.t_dates.LAST
            LOOP
                SELF.DEBUG('Date: ' || TO_CHAR(SELF.t_dates(i), 'YYYY-MM-DD') || '    Amount: ' || TO_CHAR(SELF.t_amounts(i)));
            END LOOP;
            
        END IF;
                            

        BEGIN
        
            ----------------------------------------------------------------------------
            -- Step 1 is to find the correct range in which the one of the IRR bounds
            -- will yield a positive NPV and the other bound will yield a negative NPV
            ----------------------------------------------------------------------------
            SELF.FIND_IRR_RANGE(n_range_guess_limit,
                                n_irr_low,
                                n_irr_high,
                                n_npv_low,
                                n_npv_high);
            
            ----------------------------------------------------------------------------
            -- Step 2 is to use the bisection method to find the actual IRR
            ----------------------------------------------------------------------------
            SELF.n_irr := SELF.IRR(n_irr_guess_limit,
                                   n_irr_low,
                                   n_irr_high,
                                   n_npv_low,
                                   n_npv_high);


        EXCEPTION
        WHEN OTHERS
        THEN
            -- Something went wrong, set IRR to NULL and log an error
            SELF.DEBUG('Exception: ' || SQLCODE || ' - ' || SQLERRM);
            SELF.n_irr := NULL;
            
        END;
            
        p_return_val := SELF.n_irr;
        
        RETURN ODCIConst.Success;
    
    EXCEPTION
    WHEN OTHERS
    THEN
        RAISE;
    
    END;
    
    --------------------------------------------------------------------------------
    --  Function: ODCIAggregateMerge
    --  
    --  Purpose:  If this function executes in parallel, this will be called to merge the results.
    --            We need to take the arrays from the second XIRR_T object and merge them into 
    --            the first one.  We loop through every date value in the second XIRR_T object
    --            and apply the exact same logic as in the Iterate function to put the values in the
    --            correct spots in first object's arrays.
    --
    --------------------------------------------------------------------------------
    MEMBER FUNCTION ODCIAggregateMerge(SELF       IN OUT XIRR_T, 
                                       p_ctx2     IN     XIRR_T) 
        RETURN  NUMBER
    IS
    BEGIN
    
        IF p_ctx2.t_dates.COUNT > 0
        THEN
    
            FOR i IN p_ctx2.t_dates.FIRST..p_ctx2.t_dates.LAST
            LOOP
            
                IF SELF.t_dates.COUNT = 0
                THEN
                    SELF.t_dates.EXTEND;
                    SELF.t_amounts.EXTEND;
                    SELF.t_dates(1)      := p_ctx2.t_dates(i);
                    SELF.t_amounts(1)    := p_ctx2.t_amounts(i);
                    
                ELSIF p_ctx2.t_dates(i) > SELF.t_dates(SELF.t_dates.LAST)
                THEN
                    SELF.t_dates.EXTEND;
                    SELF.t_amounts.EXTEND;                        
                    SELF.t_dates(SELF.t_dates.LAST)       := p_ctx2.t_dates(i);
                    SELF.t_amounts(SELF.t_dates.LAST)     := p_ctx2.t_amounts(i);
                    
                    
                ELSE
                    FOR j IN SELF.t_dates.FIRST..SELF.t_dates.LAST
                    LOOP
                        IF p_ctx2.t_dates(i) <= SELF.t_dates(j)
                        THEN
                            IF p_ctx2.t_dates(i) = SELF.t_dates(j)
                            THEN
                                SELF.t_amounts(j)   := SELF.t_amounts(j) + p_ctx2.t_amounts(i);
                                
                            ELSE
                        
                                SELF.t_dates.EXTEND;
                                SELF.t_amounts.EXTEND;

                                FOR k IN REVERSE (j+1)..SELF.t_dates.LAST
                                LOOP
                                    SELF.t_dates(k)       := SELF.t_dates(k-1);
                                    SELF.t_amounts(k)     := SELF.t_amounts(k-1);
                                    
                                END LOOP;
                                
                                SELF.t_dates(j)     := p_ctx2.t_dates(i);
                                SELF.t_amounts(j)   := p_ctx2.t_amounts(i);
                            
                            END IF;
                            
                            EXIT;
                            
                        END IF;
                    
                    END LOOP;
                    
                END IF;
            
            END LOOP;
        
        END IF;
        
        RETURN ODCIConst.Success;
        
    END;
 
END;
/


CREATE OR REPLACE FUNCTION xirr_agg(p_input VARCHAR2) 
    RETURN NUMBER 
PARALLEL_ENABLE AGGREGATE USING xirr_t;
/