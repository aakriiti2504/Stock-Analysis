# Stock-Analysis

## Overview of Project: 

### Background:
My friend Steve just recently finished his finance degree. His parents are absolutely proud of him and have decided to be his first clients. Steve’s parents are very passionate about green energy. Since fossil fuels are rapidly getting used up, they believe there is going to be more reliance on alternative sources of energy in near future. There are a number of alternative green energies like wind energy, geothermal energy, hydro-electricity and bioenergy to invest in. Steve’s parents have not done much research but have decided to invest their money in DAQO New Energy Corporation.  This company makes wafers for solar panels and that’s all the information his parents know of. DAQO’s ticker symbol is DQ. Steve has promised his parents to look into the DAQO’s stocks but he is more concerned about diversifying their funds. He wants to analyze some other green energy stocks for that matter.

### Purpose:
Steve is assisting his parents with financial investment decisions. They are green energy enthusiasts and want to invest their money in DAQO New Energy Corporation. For stock analysis, Steve has created an excel file with data of all other green energy stocks and seeking help from me to help him with the analysis. Steve is well versed with Excel but is looking for help with the stock analysis.

### Tools used:
For the purpose of stock analysis, we will be using an extension of Microsoft Excel. Visual Basic for Applications (VBA) will be used to automate the tasks. VBA is a programming language that interacts with Microsoft Excel. Through VBA, one can read and write to cells in worksheets. It can make calculations and use complex project to perform analysis. Using code to automate the analysis, Steve can reuse it with any stock and highly reduce the chances of errors and misinterpretations. The code will also be time efficient. The code written in VBA will automate the analysis done via calculations in Excel. These automated tasks are called Macros. VBA is the visual form of BASIC created by Microsoft which is a visual form builder to build graphical desktop applications. To enable VBA in Excel, the Developer tab was added to the ribbon in Excel. A macro-enabled workbook is created with an extension of ‘xlsm’.


### Analysis:
Steve wants to find the total daily volume and the yearly return of each of the stocks. The total number of shares traded throughout the day is called as the ‘Daily Volume’. It measures how actively a stock is traded. The percentage difference in price from the beginning of the year to the end of the year is called as the “Yearly Return’. Steve’s parents are curious about the DAQO’s stocks so we start with DQ first.

### Procedure:
We will first setup a worksheet to hold the stock data. We will access the data in VBA from Excel using the Range () and Cells () functions. Here, the Range () method belongs to the worksheet object that we activated. We will set the value of cell A1 to ‘DAQO (Ticker:DQ)’ with code:
Range(“A1”).Value = “DAQO (ticker: DQ)”

We then use cells () to create a header for cells A3 through C3 with column names Year, Total daily Volume and Return. Here, instead of using Range () we will use Cells () because it adds more flexibility when we move to automated code because individual numbers are easier to work with than strings of cell coordinates. When filling in the table below the header, we use the similar pattern of code with the row value using the variable is specified.
Calculating Total daily Volume for DQ in 2018:
Steve’s parents want to know how actively DQ was traded in the year 2018. According to their belief, if a stock is traded often, then the price accurately reflects the value of the stock. If all the daily volumes for DQ are summed up, the yearly volume can be calculated. We can also get a rough idea of how often the stock gets traded.
In order to find the total daily volume, we will loop through all rows in the stock data worksheet and check if the ticker for that row is DQ. If so, we then add its daily volume to the total volume.

### Calculating yearly return for DQ performance in 2018:
The yearly return is the percentage increase or decrease in price from the beginning of the year to the end of the year. If someone invested in DQ at the beginning of the year and never sold, then the yearly return is how much the investment grew or shrunk by the end of the year.


### Finding:
-	DAQO dropped over 63% in 2018. Hence Steve definitely would want to offer better stocks to his parents.


Since DAQO might not be the best option for Steve’s parents, we can analyze multiple stocks to find better choices for them. Analysis code used for DQ Analysis can now be repurposed to analyze various multiple stocks.

All Stocks Analysis worksheet has the output for analysis of multiple stocks.
To run analysis on all of the stocks, its important to create a program flow that loops through all of the tickers. We can create a list of tickers and have VBA handle the code, using a for loop and an array.
If Steve may want to look into a different set of stocks in future, we need to create a flexible macro for running multiple stocks. 




### Results: 
Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.





### Summary: In a summary statement, address the following questions.



#### 1. What are the advantages or disadvantages of refactoring code?






#### 2. How do these pros and cons apply to refactoring the original VBA script?
