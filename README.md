Module 2 Challenge

## Overview of Project

Steve was interested in knowing about the performance of thousands of stocks performance in the past couple years.  To do this, I must refactor the code created in a previous excersize to 
account for this, and to speed up the analysis.  This is due to the increase in information that will be processed.

In the end we need to assess stock performance so Steve can choose which investements he will make a move on.

## Results - How did we accomplish this, and what were the results.

In the original script, we were examining several stocks, and doing so by looping thru the individual stocks and identifying each individual stock, then looping back thru the set to find the next.  
Due to the detailed nature of the loop, the initial script would run between 1.6 and 2.1 seconds respectively.  After the code refactoring, the elapsed time for the script was reduced significantly.  After the refactoring the code ran in roughly .4 seconds respectively.  

The difference in processing speed between the two years is negligable. 

The 2017 performance:  https://github.com/JohnJohnson913/Stock-Analysis/blob/main/VBA_Challenge_2017.png
The 2018 performance:  https://github.com/JohnJohnson913/Stock-Analysis/blob/main/VBA_Challenge_2018.png

The introduction of the for loop and the conditional statements created a much more efficient code, allowing for the elapsed time to complete much more effeciently, and work for much larger datasets.

The sheet was formatted, much like it is prior the refactoring, however the formatting doesn't appear to have impacted the elapsed time for the code to run.

### In Summary, what are the advantages and disadvantages that apply to refactoring the original VBA Script?

What are the advantages of refactoring code:

Refactoring code allows you to use what's already there.  This can be a significant time saver, allowing the developer to focus on the performance of the code as
opposed to creating the code from scratch.  Many times it is easier to start from scratch, but if the existing code is of good quality it will allow quick and effecient insight
to assess and modify code to desired outcome.

What are the challenges of refactoring code:

Like many other situations, it can be difficult to enter into something mid-stream.  Often, it can take time to get a full understanding of what the code is trying to do.  Additionally, refactoring code 
may take an unknown amount of time, given the variables may not completely be accounted for prior to starting off.
