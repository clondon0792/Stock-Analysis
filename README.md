# VBA Challenge for DQ Analysis

## Overview of Project

###### Situation

This situation was requested for the sole purpose of providing easier to digest data when looking through stocks for a specific client. In this instance we have 12 different stock types and the data sets for how well the stock did within a 12 month period, amongst two years (2017-2018). There was the initial code (noted as AllStockAnalysis) in Module 2 of the Visual basic editor, which provided a fairly slow output of approx 91 seconds.

###### Task
For this project, adjustments were made to refactor the code in a way that may be more efficient then the previous code. On worksheet All Stocks Analysis, four buttons are present. One button to clear the work sheet entirely. One to run the initial AllStocksAnalysis script, one to run the refactored script, and one button to apply conditional formatting to the present data in a while that is easily discernable for someone who is not as familiar with excel or data sheets. This is to distinguish which code is actually running in a more efficient manner to further benefit the client. Both scripts must receive the ticker, total volume, and return of each stock and print it into the All stock Analysis work sheet.

##  Results

Between the sets of code, the refactored code surpasses the initial code by a wide margin. As can be seen in images 2017_Non_Refactor.PNG and 2017_Refactor.PNG

###### 2017_Non_Refactor
![2017_Non_Refactor](https://github.com/clondon0792/Stock-Analysis/blob/main/2017_Non_Refactor.PNG)

###### 2017_Refactor
![2017_Refactor](https://github.com/clondon0792/Stock-Analysis/blob/main/2017_Refactor.PNG)

As can be seen evidently through the photos above, the initial code took approximately 91 miliseconds in order to run the code that was produced through the course of Module 2. With various degrees of modifications to the code, allowing the script to run through the data given at a much quicker pace, we see a decrease in the amount of time it takes to run the specific code. This is seen as a net gain, both for the developer and for the client as well. Therefore, the refactored code would be more useful to use in this circumstance to quicker provide the requested service to the client.

This can also be deduced, by looking at the 2018 results as well.

###### 2018_Non_Refactor

![2018_Non_Refactor](https://github.com/clondon0792/Stock-Analysis/blob/main/2018_Non_Refactor.PNG)

###### 2018_Refactor

![2018_Refactor](https://github.com/clondon0792/Stock-Analysis/blob/main/2018_Refactor.PNG)

In this example as well, much like the 2017 results, it can be determined that the refactored code has improved over all functionality of the script and is running at an ideal speed of what one would desire when running data anaylze macros.

## Summary

In summary, looking at the results above, it can be determined that refactoring and adjusting code to perform at a higher capacity in a shorter amount of time will always be beneficial to the developer and the client. However, it is not always the best option to continue tweaking code in order to just lower run time. In some circumstances when one has plenty of time to adjust and change things without fear of missing a given deadline, then a refactoring of ones code would be extremely beneficial. If a deadline is approaching far too quickly and refactoring code will only break existing code, this could result in missed deliverables and potentially lost clientele. 

In the above example, it may not be necessary to refactor the code if Steve was perfectly happy with the amount of time the initial macro had taken. To a developers eye, .72 miliseconds is a long time for a macro to run because it could be considered inefficient. In a laymans understanding, however, it is possible that the refactor/change to the functionality of the script would go unnoticed by anyone other than the developer. That's not to say it would still not be beneficial in its efficiency, but that it is possible the work may be looked over entirely if it isn't what the client had requested.
