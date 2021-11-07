# Stock-Analysis

This project analyzes stock datasets for the years 2018 and 2017. Details are explained below:


# Overview of Project: 

### Background:
My friend Steve just recently finished his finance degree. His parents are absolutely proud of him and have decided to be his first clients. Steve’s parents are very passionate about green energy. Since fossil fuels are rapidly getting used up, they believe there is going to be more reliance on alternative sources of energy in near future. There are a number of alternative green energies like wind energy, geothermal energy, hydro-electricity and bioenergy to invest in. Steve’s parents have not done much research but have decided to invest their money in DAQO New Energy Corporation.  This company makes wafers for solar panels and that’s all the information his parents know of. DAQO’s ticker symbol is DQ. Steve has promised his parents to look into the DAQO’s stocks but he is more concerned about diversifying their funds. He wants to analyze some other green energy stocks for that matter.

### Purpose:
Steve is assisting his parents with financial investment decisions. They are green energy enthusiasts and want to invest their money in DAQO New Energy Corporation. For stock analysis, Steve has created an excel file with data of all other green energy stocks and seeking help from me to help him with the analysis. Steve is well versed with Excel but is looking for help with the stock analysis.

### Tools used:
For the purpose of stock analysis, we will be using an extension of Microsoft Excel. Visual Basic for Applications (VBA) will be used to automate the tasks. VBA is a programming language that interacts with Microsoft Excel. Through VBA, one can read and write to cells in worksheets. It can make calculations and use complex project to perform analysis. Using code to automate the analysis, Steve can reuse it with any stock and highly reduce the chances of errors and misinterpretations. The code will also be time efficient. The code written in VBA will automate the analysis done via calculations in Excel. These automated tasks are called Macros. VBA is the visual form of BASIC created by Microsoft which is a visual form builder to build graphical desktop applications. To enable VBA in Excel, the Developer tab was added to the ribbon in Excel. A macro-enabled workbook is created with an extension of ‘xlsm’.


### Analysis:
Steve wants to find the total daily volume and the yearly return of each of the stocks. The total number of shares traded throughout the day is called as the ‘Daily Volume’. It measures how actively a stock is traded. The percentage difference in price from the beginning of the year to the end of the year is called as the “Yearly Return’. 

### Procedure:
We will first setup a worksheet to hold the stock data. We will access the data in VBA from Excel using the Range () and Cells () functions. Here, the Range () method belongs to the worksheet object that we activated.All Stocks Analysis worksheet has the output for analysis of multiple stocks.

To run analysis on all of the stocks, its important to create a program flow that loops through all of the tickers. We can create a list of tickers and have VBA handle the code, using a for loop and an array.
If Steve may want to look into a different set of stocks in future, we need to create a flexible macro for running multiple stocks. Using VBA the starter code provided is refactored so that we can loop through the dta one time and collect all the information. The refcatored code is expected to be more efficient than the original code.

![Capture](https://user-images.githubusercontent.com/23488019/140642084-f8fb748d-2d2b-47dd-9bab-f8049ae78964.JPG)

# Results: 
Using images and examples of the code, we will compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script. The steps taken for refactoring are given below:

- #### Step 1a - 
Variable 'tickerIndex' was created and set to 0. TickerIndex will be used to access the correct index across the four different arrays that will be used, the ticker array and the three output arrays created in the next step.


![1a](https://user-images.githubusercontent.com/23488019/140641067-7c4bf99f-35e8-4625-ae45-7cbe55deccd4.PNG)

- #### Step 1b - 
The three output arrays 'tickerVolumes', 'tickerStartingPrices' and tickerEndingPrices are created of data types Long, Single and Single respectively.


![1b](https://user-images.githubusercontent.com/23488019/140641152-23a752b3-31e7-4001-8980-8c223e36f7d1.PNG)

- #### Step 2a -
A for loop to initialize the 'tickerVolumes' to 0 is created.


![2a](https://user-images.githubusercontent.com/23488019/140641569-c14f7412-2488-4512-b212-7eb25fb2dd68.PNG)

- #### Step 2b - 
A for loop that will loop over all the rows in the spreadsheet is created.


![2b](https://user-images.githubusercontent.com/23488019/140641585-71c3ca09-6786-4cd0-89c5-cd46948774b7.PNG)

- #### Step 3a - 
Script for inside the for loop in Step 2b a script that increases the current 'tickerVolumes' variable and adds th eticker volume for the current stock ticker is written.

![3a](https://user-images.githubusercontent.com/23488019/140641599-6012134f-8ce3-4985-9288-c5d2c9dc33fa.PNG)

- #### Step 3b - 
An if- then statement is written to check if the current row is the first row with the selected 'tickerIndex'. If it is, then the current starting price is assigned to the 'tickerStartingPrice' variable.



![3b](https://user-images.githubusercontent.com/23488019/140641610-3e0e4a47-a2e1-44a1-a1f8-c39258597b5f.PNG)

- #### Step 3c - 
An if- then statement is written to check if the current row is the last row with the selected 'tickerIndex'. If it is, then the current closing price is assigned to the 'tickerEndingPrice' variable.



![3c](https://user-images.githubusercontent.com/23488019/140641624-1998f5be-097c-4fb0-9bc1-13fdfe97439c.PNG)

- #### Step 3d - 
A Script that increases the 'tickerIndex' if the next row's ticker doesnt match th eprevious row's ticker is written.


![3d](https://user-images.githubusercontent.com/23488019/140641633-fde27e95-2238-4d8b-8080-25c67a0e6f33.PNG)


- #### Step 4 - 
A for loop is used to loop through the arraays to output the three columns in the spreadsheet.


![4](https://user-images.githubusercontent.com/23488019/140641644-73958898-c094-47d2-b45d-55a59b4f7b50.PNG)


- #### Run Stock Analysis -
On running the stock Analysis it can be noted that the outputs for the years 2017 and 2018 are same after refactoring the original code. It can be noted that all the output values are same after running both the sets of codes. However,  due to the timer messagebox in the code we can compare the time of original script and refactored script and note that the refactored timer was lesser than the original. Hence the refactored code is more efficient.
    
- #### Output of Original Script and Refactored Script for Year 2017
![2017 a](https://user-images.githubusercontent.com/23488019/140641795-78fc3a51-2cd4-4f3d-a857-733f6686cfec.PNG)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/23488019/140641775-e295f538-415a-48ff-a67c-5aab27951ad4.PNG)

The first image is the original one and the last image is of the refactored code. We can note that the time taken to get output from original script in the year 2017 is 0.5429688 seconds. The time taken to get output from the refactored code in the year 2017 is 0.1137695 seconds.

- #### Output of Original Script and Refactored Script for Year 2018

![2018 timer](https://user-images.githubusercontent.com/23488019/140641918-a2625fda-cdf1-48e1-acb2-9214ac7bef1c.PNG)

![VBA_Challenge_2018](https://user-images.githubusercontent.com/23488019/140641884-a43fce35-aaba-4418-a354-959da0d0103a.PNG)

The first image is the original one and the last image is of the refactored code. We can note that the time taken to get output from original script in the year 2018 is 0.5460205 seconds. The time taken to get output from the refactored code in the year 2018 is 0.1171875 seconds.

# Summary: 

### 1. What are the advantages or disadvantages of refactoring code?
#### Advantages -
Refactoring is the process of writing reusable code so that it is time efficient and easier to use. Refactoring is an important part of any coding script. We do not add new functionality but just make the code much more efficient. Efficiency can be increased by taking fewer steps, using lesser memory or by improving the logic of the code to make it easier for future readers to read. It is an integral part of coding also because sometimes the first attempt to write a code is not always the best way to accomplish a given task. Improvement in the overall code results into shorter lines of code. Extra lines of redundant loops and unnecessary codes can be easily removed to avoid complexity.

#### Disadvantages -
A thorough understanding of the VBA syntax is essential in refactoring and making the code more efficient. If errors occur in the code, it gets challenging to fix the errors and make the code run successfully. The syntax's exact requirements makes it difficult to type sinec error is returned before an edit is completed. Hence, typing needs to be done carefully. Refactoring an existing code may introduce errors and make it difficult to fix bugs. Hence it can be challenging to decide between refactoring a code and starting fresh.

### 2. How do these pros and cons apply to refactoring the original VBA script?

Refactoring the code decreases the processing memory needed for processing big chunks of datasets. Hence the refactored code reduces the number of loops and extra lines of code to make the code more efficient and easy to understand. By refactoring the original VBA script, the runtime between th eoriginal and the refactored script decreased significantly. while refactoring, it is essential to test the code so that the efficiency of the code can be determined.

For example we can see that in 2017 original script yielded result in 0.5429688 seconds, whereas after refactoring the same code the result was calculated in 0.1137695 seconds. Hence there was an improvement in time and the refactored code was much more efficient. 
![2017 a](https://user-images.githubusercontent.com/23488019/140640447-f844691f-0d9a-4b93-955e-21c17262bfd7.PNG)

![VBA_Challenge_2017](https://user-images.githubusercontent.com/23488019/140640558-c682484d-45d0-4717-b696-58ab44d070ca.PNG)

The first image is the original one and the last image is of the refactored code. We can note that the time taken to get output from original script in the year 2017 is 0.5429688 seconds. The time taken to get output from the refactored code in the year 2017 is 0.1137695 seconds.
