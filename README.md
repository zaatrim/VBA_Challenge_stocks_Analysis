# **Stocks Performance Analysis**

## *Project Overview*
   Steve's Parents would like to invest in renewable clean energy and asked Steve to help them with their investment decision. For this purpose, Steve needs to run a stock performance analysis to recommend stocks for investment for his parents and be able to run a similar analysis on all stocks in Dataset for future analysis. His analysis uses the following 2 key performance indicators (KPIs), Yearly return as measured by calculating “Year-end Price / Year start price-1”, and trading Volume; assuming if a stock is traded often, then the price will accurately reflect the value of the stock. Steve asked for Help in running his analysis, In one push-button Steve would like to get the above two KPIs in a workbook. I will use Excel VBA Code, Excel Workbook to run through the Dataset & calculate the above two KPIs to help Steve with his analysis. 
                 
## *Analysis & Results*
### Analysis
Steve already collected Stocks trading historical data, which includes Pre each stock Name &trading Date stocks Names, Stock Opening, High, Low, closing prices, and Stocks trading volume through the years. To help Steve, I will reuse, Edit or refactor, the current solution EXCEL VBA code (from Module 2) to loop through all the data one time to collect the following same information:                                                                                                                                            * Total Daily Trading Volume Per stock for the selected year Per each stock.
    Stock close price on the 1st trade date in the Year and the  stock close price on the last date in the year. And then calculate per each stock the Yearly return using the following formula
       • Yearly Return per stock= stock Ending Price/stock Starting Price - 1
    
Since Steve wants to run his analysis on the entire dataset, I will refactor the code to run it more efficiently. 
        • I will define Arrays for the following parameters:
          •  Stocks List,"tickers"
          •	Stock Volumes, stock starting price, and stock Ending Price 

  Isnset tickers index image 

   Define stock indexing parameter and then Loop one time      over-all rows in stocks dataset per Year to assign and calculate the values for the above-defined Arrays (per each stock Index).
            Insert Loop image   

   The last step in this analysis, store the analysis outcome in Table in a new worksheet to help Steve make his recommendation.

   Insert output image here 
          
### Results
   1) Analysis results will focus on two factors:
     a.	Stock’s analysis conclusions
         i.	2017 was a good year for most renewable energy stocks (except for “TERP” stock).

   insert "allstocks 2017 output" image

   ii.	 2018 presented Negative return on most stocks except for “ENPH” and “RUN” stocks. 
               insert "allstocks 2018 output" image
       b.	There is no clear correlation between Stock Trading volume and yearly stock return.
            
   insert stock performace charts  image
 
   2)	In Code refactoring, I edited the code to loop over all rows at one time (Rather than doing a nested Loop, the system has to loop over all rows for every single stock in the Array). The refactoring significantly reduced the run time. For the Year 2018 original Code, the runtime was 0.8710938 seconds, while the refactored code runtime is 0.2109375 seconds. For the year 2017 original Code, Run time was 0.9179688 seconds vs. the refactored code runtime is 0.1953125 seconds

        Refactored Code Runtime 			Original Code Run time 
                insert 2018runtime image
                   
        Refactored Code Runtime 			Original Code Run time 
                 insert 2017runtime image

## *Summary*
### Advantages
-   a. Refactoring is a Key part of the coding process. When refactoring code, I did add new functionality; I want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.
-   b. Code refactoring is Important. It will enable developers to reuse someone else’s code to optimize their code.
-   c. In the specific stock analysis code. The refactored Code present ~81% - 85% Improvement in code Runtime. For large-scale stocks data set, this is a significant improvement to enable the code to scale up on large numbers of stocks.

### Disadvantages

-   a. The refactored code will work under the conditions that Dataset is sorted by the Stock Name and then sorted ascending by trade Date. If these two conditions are not met, the code will not run properly.
- b. The Array for the refactored code is not dynamic, the user has to: provide the list of the stocks and the Year. There is a place for additional refactoring to make the code mode dynamic such as:
    - •	The code will identify the year through the data.
    - •	The code will identify the list of the stocks from the Dataset instead of having it hardcoded. 
