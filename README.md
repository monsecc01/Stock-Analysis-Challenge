
# Stock Analysis with Excel VBA
*Module 3 Challenge-VBA Stock Analysis

---

## Project Overview: 

The purpose of this analysis is to refactor our current VBA code for 2018 and 2017 stock analysis to allow for higher efficiency. Success of script efficiency is then measured by script run time; by comparing the refactored run time with the non-refactored script run time we will determine if refactoring succesfully made the VBA script run faster. Our findings will be presented to Steve as he is interested potentially using our code with an expanded data set to include the entire stock market.

## Results: 

Analysis of stock data between 2017 and 2018 show that stock performances were more successful in the year 2017 as 11 of the 12 tickers show a positive return. However, only 2 stocks had a positive return in the year 2018, which means that most stocks were not a good long term option for investments. The only exceptions being ENPH and RUN who kept a positive return in the year 2018.

<img width="200" alt="All Stocks 2017 Results" src="https://user-images.githubusercontent.com/81447450/113518701-97660c80-954d-11eb-97dd-b44b16cc0885.png">
<img width="200" alt="All Stocks 2018 Results" src="https://user-images.githubusercontent.com/81447450/113518703-9a60fd00-954d-11eb-8eda-8b8cacae4f48.png">

In terms of our code execution times, the refactored script was proven succesful at reducing script run time. The original script took over 1 second to run. In comparison, the refactored script took drastically less time to run, approximately 0.25 seconds. This is a time reduction of over 75%, which proves that the refactored code is more efficient. 

*Original Script run time:
<img width="239" alt="Original 2017 Run Time" src="https://user-images.githubusercontent.com/81447450/113518837-5b7f7700-954e-11eb-896b-c8f403fd85da.png">

*Refactored 2017 Run Time:
<img width="245" alt="Refactored 2017 Run Time" src="https://user-images.githubusercontent.com/81447450/113518839-5e7a6780-954e-11eb-9185-4d80a504f761.png">

## Summary: 
### What are the advantages or disadvantages of refactoring code?
Two major advantages of refactoring code are maintainability and efficient execution. Simplyfing the code and improving the design of software allows for code that is easier to understand and therefore maintain. As seen in our Stock Analysis code, refactoring also improves execution of code by reducing script run time. 

Some disadvantages of refactoring code is that it can be time consuming and, in some cases, counterintuitive. Testing if the new code is working may require a lot of trial and error to make sure bugs are not introduced. Also, refactoring code requires the programmer to think about the computer logic behind the code which is sometimes difficult but can be a great skill to adquire.

### How do these pros and cons apply to refactoring the original VBA script?
Refactoring of the original VBA script allowed for improvement in code execution. Script run time was reduced from over 1 second to approximately 0.25 seconds for both 2017 and 2018 data. 
<img width="245" alt="Refactored 2017 Run Time" src="https://user-images.githubusercontent.com/81447450/113519619-05610280-9553-11eb-8c23-e89f19d5bb04.png">

<img width="242" alt="Refactored 2018 Run Time" src="https://user-images.githubusercontent.com/81447450/113519622-072ac600-9553-11eb-8713-84076e95fd93.png">

When refactoring the original VBA, the code had to be evaluated for bugs to make sure it ran as intended. This was time consuming and tedious at times. 
