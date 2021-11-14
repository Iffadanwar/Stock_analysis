# Stock_analysis
Analysis of stock for Module 2 Challenge

### Purpose
The purpose of this assignment is to edit or refactor, the Stock Market Dataset with a VBA module that loops through all the data set once to collect all the data needed to analyse the stock market provided. Then we need to check if refactoring our code made it fast. lastly, we need to make our code more reliable and efficient by using less memory or by improving the logic of the code used. Using the green_stock file that provides us with the data required for this experiment we will try to find the "total daily volume of trade" and the "per cent growth" that particular ticker brings.

### Analysis and Challenges
We need to do the following tasks:

* Prepare ourVBA_Challenge.vbs file.
* Create our resources folder in GitHub to hold the run-time pop-up messages that we’ll screenshot after running refactored analyses for 2017 and 2018.
* Create and convert our XLSM file from *.vbs dataset that you used in this module as VBA_Challenge.xlsm.
* Add the VBA_Challenge.vbs script to the Microsoft Visual Basic editor.
* Use the steps Refactor VBA code and measure performance to add code where indicated by the numbered comments in the starter code file.

### Background
Steve loves the workbook you prepared for him. At the click of a button, he can analyze an entire dataset. Now, to do a little more research for his parents, he wants to expand the dataset to include the entire stock market over the last few years. Although your code works well for a dozen stocks, it might not work as well for thousands of stocks. And if it does, it may take a long time to execute.

In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. Then, you’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, you’ll present a written analysis that explains your findings.

Refactoring is a key part of the coding process. When refactoring code, you aren’t adding new functionality; you just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read. Refactoring is common on the job because first attempts at code won’t always be the best way to accomplish a task. Sometimes, refactoring someone else’s code will be your entry point to working with the existing code at a job.

### Results
##### Results of the stock analysis
2017 seems to have a much better trading year than 2018 for these particular tickers. Every company, except for TERP, had shown strong growth in the year 2017 but most shows a decline in stock price for the year 2018. The only 2 companies that show an increase in stock price had the ticker "ENPH" and "RUN". I would recommend steve invest in these 2 companies as the data shows and they both have strong growth even during a bearish stock market.
##### Results of reformating our code
* The time to run the code for both years using my original code took 0.468755 seconds
* The time to run the code for both years using refactored code took 0.484375 seconds
The time to run the code increased after refactoring meaning that I added complexity/ steps to the code. Use of the original "tickers" but directly referencing its value, e.g (ticker(i)) in a loop may have reduced the number of steps it takes as compared to using a ticket index. I also added additional formatting which may have also increased the time needed to run the code. There is also an additional for loop to ID the ticket index to give it the tickers value.

### Summary 
##### listing the advantages and disadvantages of refactoring code
###### Advantages 
* Refactoring code allows the reduction of processing time which is vital for large sets of data. 
* It allows us to re-write code to avoid code repetition, and duplicate lines and form into one location that can be back referenced. 
* logic errors are easier to sight and change.
* shows us the programming logic in a more comprehensive manner.
* It helped me avoid "Magic Numbers".
* It also better formated my code and allowed me to add descriptive comments in the middle.
###### disadvantages 
* allows us to split our code into functions that can be hard to manage.
* can affect the outcome if further mistakes are made. 
* Can be time consuming.
* For our code specifically is increased processing time.
