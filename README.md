# VBA_Challenge2
Module Challenge 2
# Module 2
## Overview of Project
### Purpose
The purpose of this project was to help Steve with suggesting green stocks to invest in for his parents. More specific Steve wanted to know how certain green stocks performed over a two year period, from 2017-2018. We manipulated the data using mircosoft excel VBA macros to run loops through different sheets of data to return the values we desired and refactoring the original code to maximize the proformance of our macro VBA script.
## Results
### Stock Performance
* Green Stock tickers, "AY","CSIQ","DQ","ENPH","FSLR","HASI","JKS","RUN","SEDG","SPWR","TERP","VSLR" in the year 2017 all returned positive returns except "TERP" as can be seen in this photo. 

![VBA_2017](https://user-images.githubusercontent.com/93004710/148656012-c1b17ef2-912b-48b3-802b-fc96c70c5981.png)


* In the year 2018 almost all green stock tickers returned a negative return with the exception of tickers "ENPH" and "RUN" as can be seen in this photo. 

![VBA_2018](https://user-images.githubusercontent.com/93004710/148656074-2676a946-f711-42f5-8795-05888421a8d7.png)


### Macro Script Performance
* The original macro VBA script had a run time for 2017 of 0.8320313 as seen here.


![VBA_2017_original](https://user-images.githubusercontent.com/93004710/148656243-4d9061e1-09be-47f5-8900-8a09013f655b.png)


* The original macro VBA script had a run time for 2018 of 0.8164063 as seen here.


![VBA_2018_original](https://user-images.githubusercontent.com/93004710/148656320-f891df15-d906-4152-a161-df6b6bbdadc4.png)


* The refactored macro VBA script had a run time for 2017 of 0.1289063 as seen here.


![VBA_2017_Refactored](https://user-images.githubusercontent.com/93004710/148656362-773af2c9-2f38-45aa-aac9-51a6f06ce40b.png)


* The refactored macro VBA script had a run time for 2018 of 0.140625 as seen here. 


![VBA_2018_Refactored](https://user-images.githubusercontent.com/93004710/148656414-2bfdb0e1-1e35-47d0-87aa-81c3b63e126e.png)


* The refactored code is here. 


![VBA_Refactored_Script1](https://user-images.githubusercontent.com/93004710/148656817-2860aba3-f3a5-467d-8820-d43a644aba13.png)![VBA_Refactored_Script2](https://user-images.githubusercontent.com/93004710/148656826-4a35573e-3f9a-41ff-84ea-bd705ceef9de.png)


* Compared to the original code here.


 ![VBA_Original_Script](https://user-images.githubusercontent.com/93004710/148656897-c297abb6-6b43-4684-9703-498626248740.png)
 
 
* The photos above show just how much more performance we get out of the refactored script for both years 2017, and 2018. This big difference in the refactored code from the original code in the nested for loop "For j" that allows the script to run more efficient.
## Summary
### Advantages and Disadvantages of Refactoring Code
* An advantage of using refactored code is increasing the run time for the script as can be seen with the Macro Script Performance photos. The code of each year ran considerably faster with the refactored code.
* A disadvantage of using refactored code is syntax errors. When refactoring it is important to make sure your for loops and nested loops are correctly labeled otherwise the code will not run properly. In the photo above you can see in the refactored script the nested loop variable changes from "i" to "j" if you do not correctly label this your script will run but improperly fill your cells. This can be a con when refactoring code.
### Advantages and Disadvantages of Refactored VBA Script
* An advantage to using the refactored VBA script is being able to use the same pattern of code without having to write an entire new script. The original VBA script did not run as efficient, however it was a more simple script to follow and debug when necessary.
* A disadvantage of using the original VBA script was the poor efficiency of the script. While the refactored VBA script runs impressivly it is more complex and time consuming to write and follow along.


 
