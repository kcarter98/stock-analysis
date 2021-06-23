# Refactoring Stock Analysis Code

## Overview

In this analysis, it was asked whether or not it is beneficial to refactor code, and if it is beneficial then how beneficial is it? Here I coded a script in VBA that looked at a list of stock information and turned the data into a readable way to analyze stock performance. After coding the inital script, I took extra time to improve it in an attempt to make it faster and more efficient.

## Results

### Change from Refactor

After refactoring the code, it was clear after it was ran how advantageous it was over the original first attempt in both the 2017 and 2018 runs. As you can see below, the original code took over 40 seconds and it was easy to tell that excel was struggling with the "not responding" message, but after a while it would pull through.

##### Original Code Results

![Image1](/resources/VBA_Original_Runtime.PNG)

### 2017 vs 2018

There was a slight increase in the performance for 2017 when compared against the 2018 run. Although, this was not as drastic of a difference when comparing the original vs. refactored code. It is possible that the consistent slight difference between 2017 and 2018 is that 2017 had a lower total volume. So, it could be a differnce in integer types causing more memory usage, but this is not conclusive.

##### 2017 Refactored Code Results

![Image2](/resources/VBA_Challenge_2017.PNG)

##### 2018 Refactored Code Results

![Image3](/resources/VBA_Challenge_2018.PNG)

## Summary

- *What are the advantages and disadvantages of refactoring code?*

The advantages of refactoring code is that it is good for catching bugs and helps make the program more efficient. It can also help to improve how the code is designed to make it easier to read and build on in the future.

Some disadvantages to refactoring code is that it is cost and time consuming. If you need something to be released quickly it might not be the best thing to do, although it can improve the software.

- *How do these pros and cons apply to refactoring the original VBA script?*

The refactoring of the VBA script unquestionably improved the performance of the code. Admittedly, the refactored script didn't even have the formatting functionality with it like the refactored did, and it wasn't even close. It did take me a few more hours to perform the refactor, but it was worth saving me those 30 seconds of runtime each time I run my operation!

