# VBA of Wall Street 

## Overview

The goal was to create a macro that takes data from an excel spreadsheet that contains stock information to then perform an analysis from the information. 
The macro was to caculate the total volume, starting prices and ending prices, and calculate the return percentage. This will the show which stock had a postive
return percentage and by how much. 

## Results
### Analysis Results
Running the macro on the 2017 data shows that the majoirty of the stocks had a positive 
return percentage excelt for "TERP" which had a 7.2 percent drop in yearly return. However
it is noted that for "DQ" the total volume is low in comparison to the other stocks in the
data so the high yearly return is not a true indicator. The macro results for the 2018 data 
shows that the majority of the stocks had a negative return percentage that year except for 
"ENPH" and  "RUN". Because t e total volume is also high then the positive yearly return 
is a true indicator. 

### Stock Analysis Refactored
As seen in the linked images is that the AllStocksAnalysisRefactored ran quickly but inputting "2018" resulted in a faster time by a slight margin.  
<img width="1440" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/102255823/166985779-02757c7d-7057-4dd7-b434-77f0eb2947ab.png">
<img width="1440" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/102255823/166985782-9c7325fa-4f85-4793-aa3a-e6c576b690aa.png">

But in comparison the AllStocksAnalysis, the refactored code ran much faster as it took half a second to run. The difference between the two codes 
is that in the refactored verison, three arrays are created to input information into and then a for loop is used to print out the array onto a excel
sheet. Whereas in the previous code, the information is output while iterating through the sheet of data. Therefore it would take longer for the code to 
run because it is switching between worksheets, getting data and then outputting. So it makes sense that it is faster the get the data first in its 
entirety and then output it afterwards. 

## Summary
### What are the advantages or disadvantages of refactoring code?
An advantage of refactoring code is the ability to optimize the run time. When working with larger data sets, the code will take longer to run and
the difference in run time is more noticable.  The disadvantage can be creating new bugs or errors that would need to be addressed.

### How do these pros and cons apply to refactoring the original VBA script?
As mentioned previously, the refactored code ran significantly faster. This demonstrates how refactored code can result in optimized run time. However,
a problem faced while refactoring was inconsistent vector sizes that caused an error within the calculation of the yearly return pencentage. In the previous code 
there was only one vector that was iterated through but it was hard coded and therefore less likely to produce errors. 
