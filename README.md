# Stock-analysis assignment

## Overview

This project involves the preparation of a workbook for Steve to help him and his parents decide on their options for the stock market using data from 2017 and 2018. More specifically, this iteration of the project is a refinement of the previous code to make it leaner and, therefore, more efficient.

### What the VBA code does

The VBA code that I use to analyze this dataset does essentially four things: _(i)_ Calculate the total value of transactions on each stock; _(ii)_ Find the annual change in the value of each stock; _(iii)_ Do the previous computations for each year, depending on user input; and _(iv)_ Keep the time spent running the whole code..

## Results

The results of the refinement in the code are visible, first and foremost, in the time in which the code compiles. As can be seen from the images below, the original code compiles the data for 2017 in around ___0.27___ seconds and the same for the 2018 in around ___0.28___ seconds.

![Original code for 2017](https://github.com/mustail/stock-analysis/blob/bec628ddd6467da5c25223af21281da10f9f8e90/Resources/Green_Stocks_2017.png)
![Original code for 2018](https://github.com/mustail/stock-analysis/blob/bec628ddd6467da5c25223af21281da10f9f8e90/Resources/Green_Stocks_2018.png)


The refactored code, however, is remarkably faster, running the dataset for 2017 in less than ___0.08___ seconds and the data of the following year, even faster, in around ___0.07___ seconds. On both datasets, the refactored code is roughly ___four times faster___. While in this particular dataset, this does not amount to much, but in cases involving tens of thousands of stocks, the difference would prove to be significant both in terms of time as well as energy spent on the processing.

![Refactored code for 2017](https://github.com/mustail/stock-analysis/blob/bec628ddd6467da5c25223af21281da10f9f8e90/Resources/VBA_Challenge_2017.png)
![Refactored code for 2018](https://github.com/mustail/stock-analysis/blob/bec628ddd6467da5c25223af21281da10f9f8e90/Resources/VBA_Challenge_2018.png)

### Some reasons why the refactored code is faster

There are a number of reasons the refactoring produces faster results. First of all, as can be clearly seen, the refactored code is much leaner and shorter. Most importantly, the original code includes a ___nested for loop___ that surveys all 3013 of the rows 12 times for each stock. Remarkably, the new code performs the same survey only once.

Moreover, the original code has an ___additional if statement___ in this for loop to determine whether to add a value to the total value. This obviously is an extra processing burden, which could become meaningful when we have much larger datasets than the existing one. The refactored code does away with this if statement thanks to its use of an array that can hold the total values of all tickers.

## Summary

As I said above, the refactoring results in a much leaner code, which, in turn, causes the macro to run much faster (almost four times). The refactoring takes advantage of the arrays to hold variable data for multiple stocks at once. This solution eliminates the need to have a nested for loop that goes through the whole data twelve times. Needless to say, this means a greater burden on the processor and the memory, a point we should keep in mind when dealing with much larger datasets.

I do not see any disadvantages to refactoring in this case. With the refactoring, we were able to reuse a significant amount of code from the original macro while keeping the final work shorter. Most importantly, the code runs more efficiently thanks to less intense processing with the refactored version.
