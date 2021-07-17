# Analyzing stocks using Visual Basic
## Overview of Project
   - I have prepared a comprehensive excel workbook using VBA that can, with the click of a button, analyze a desired year of stock data and give individual stock information regarding total daily volume, and percent return that is formatted to quickly see a positive or negative return for the year instantaneously. 


## Results
   - I ran the refactored code through the years of 2017 and 2018 to compare the difference in performance in the two years. Let's start with 2017 and look at the results below:

## 2017
<img width="1000" alt="Screen Shot 2021-07-17 at 1 37 00 PM" src="https://user-images.githubusercontent.com/86524863/126045359-4b0ddb0e-659a-4afc-b2b1-0f1eced8f91c.png">
    
   -As we can see, 2017 was a very good year in regards to a positive returns for all the stocks except TERP. Some had a positive return of almost 200% in the case of DQ. This was a lot of data to analyze so it was crucial to refactor my base code I used to analyze one stock (DQ), so it was efficient and ran the entire year as quickly as possible. To differentiate between multiple tickers I had to check if the next row was the first row to contain that stocks ticker with the code below:
   
   >***If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1) <> tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value***

   -Then to make sure we stopped at the last row containing that ticker I had code that read as follows:
   
   >***If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1) <> tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value***
  
   -If these conditions were met, our tickerIndex would increase by 1; this was how we tallied up the individual ticker information row by row and compiled the data for individual stocks for the entire year desired. My original base code program ran in 0.8 seconds, as you can see below with refactoring my code now runs in about a fifth of that time:

<img width="1417" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/86524863/126045399-04a2be8a-eb78-459d-83e4-1c9f06bf98f3.png">


## 2018

<img width="1000" alt="Screen Shot 2021-07-17 at 1 37 31 PM" src="https://user-images.githubusercontent.com/86524863/126045387-8a3334a4-a5e7-407a-b643-c0b816daab7c.png">
    
   -Not every year in the stock market can be bullish, and we can see in 2018 the market took a downward turn. We had every stock lose on their return except two: ENPH and RUN. All of the data for the year ran in just under 0.8 seconds using my base code, as you can see below we have greatly improved efficiency by reducing the time to run the macro in just over 0.13 seconds:

<img width="1413" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/86524863/126045407-639c3e02-a487-479c-9e7b-eccd7cb0ad29.png">


## Summary

   1. When you start a project, the main goal is to make sure you can get the code to do exactly what you want it to do- fixing bugs, syntax, and typo errors. It is very important to not stop there though, a code should always be reanalyzed after the first attempt at writing it to try and increase the efficiency and reduce the chance of errors while running it. After finishing a project it's very important to step back, look at the code as a whole, and analyze how the code was written from start to finish and see where improvements can be made. Oftentimes as we write the code we get better as we go along and we can apply those new efficient methods to improve the code before it. The only pitfall to try and avoid can be rewriting code that you think will make it more efficient, but changes the outcome of the analysis to something unintended. Always make sure you haven't affected the integrity of the results when making improvements. 
   2. Originally our code analyzed one particular stock- in the refactored code we looked at a whole year of information containing 12 different stocks and all of their data. This gives our refactored code more utility and scalability. We get a broader picture of multiple stocks in a given year. We also made our code run more efficiently thus saving time- a half a second may not seem long but if this calculation was scaled up and done multiple times for more stocks or with more information, the time saving would be immense over multiple calculations. Refactoring always has to be done carefully and correctly to make sure the initial integrity of our code wasn't compromised unexpectedly. 

**The stock market can be a volatile place, but with the correct amount of caution and perspective you can turn a great return into profit, and a downward slide into an investment opportunity. Keep at it!** 


