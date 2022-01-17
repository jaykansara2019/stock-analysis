# **Stock Analysis**

## **Purpose**
This analysis aimed to work on the existing code to make it run faster, also known as Refactoring.

## **Results**

The time to run the original code v/s the refactored code has been depicted in the image below:

![Refactored vs Original code](https://github.com/jaykansara2019/stock-analysis/blob/531536d787a1169cf73592ccb696ba43a3e89bee/Refactored%20vs%20Original%20code.png)

As shown in the image below, the creation of TicketIndex, instead of making code go through the individual ticker, made a significant difference in the run time.

![Refactored code](https://github.com/jaykansara2019/stock-analysis/blob/531536d787a1169cf73592ccb696ba43a3e89bee/Refactored%20code.png)

The time was decreased by five fold in the refactored code compared to the original version.

The stocks performed remarkably well in the calendar year 2017. The return was particularly great with ***DQ, ENPH, FSLR and TERP***. However, the performance of most of the stocks plummeted sharply in the year 2018. The only stock that performed well in 2018 was ***RUN***. Although ***ENPH*** return was good, it did not perform as well as in 2017.


## **Summary**
### **Advantages of refactoring a VBA code**
- The major advantage of the refactored code is the time reduction by fivefold. Although it might not seem quite significant, if the spreadsheet has 10 to 100 times more data than the existing one, it would make a notable difference.

- Refactoring VBA code made us realise what kind of functions work well but take more time to run. Hence to avoid such using functions unless absolutely necessary.

### **Disadvantages of refactoring a VBA code**
- For a small volume of data, refactoring a code might not be a fruitful exercise, since it may not save huge time.

### **Advantages of refactoring a code in general**
- Refactoring a code, in general, is a good practice as a Quality Assurance to seek continuous improvement.
- At the organisation level, refactoring existing code also gives us a chance to implement newly updated functions that have been pushed recently by the creators.

### **Disadvantages of refactoring a code in general**
- One of the major disadvantages of the refactoring code at the organisation level is a lack of a proven track record. The existing code might have been in production for years and running perfectly fine. In the hope of achieving marginal improvement, the refactored code might give errors in the long run if not meticulously designed.

- In an environment like healthcare and aviation, refactored code has to go to rigorous testing before being pushed into production to avoid any dire outcomes. The refactored code that only results in incremental improvement might not be beneficial in a highly regulated industry.
