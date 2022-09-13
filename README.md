# VBA_challenge

## Project Overview

The purpose of this exercise was to evaluate the performance of 12 different stocks in the years 2017 and 2018 using VBA macros. As the analysis was done for a client/acquaintance, the presentation of the analysis needed user-friendly mechanisms to searching for information and interpret results. 

For the student, the exercise illustrated how to extract analysis from raw data by writing code on VBA. The first part of the exercise offered step by step coding instructions. The second part of the exercise illustrated a more efficient way to run code by a method called "refactoring." Refactoring consists of optimizing a code's execution time by trimming redundancies.

A comparison between the long and refactored codes shows that refactoring significantly reduces a query's execution time. 

The refactored code is also easier to read because the instructions are given in a continuous fashion, rather than a collection of subqueries as in the initial (i.e., long) exercise. 


## Results: Long v. Refactored Code

The following screenshots illustrate the efficiency of the refactored script/code. 

The screenshot below shows that the long script took .25 seconds to retrieve stock performance results in 2018. 

<img width="846" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/111469959/189672720-3674a142-8759-4f8f-8240-9feb26957d53.png">


By contrast, the screenshot below shows the refactored script took .08 seconds to retrieve the same results. 

<img width="846" alt="VBA_Challenge_2018_RF" src="https://user-images.githubusercontent.com/111469959/189672735-fea35fc9-a520-4bf1-9bfe-2c59bed905ae.png">


## Summary: Pros and Cons of Refactoring Code 

### Advantages and disadvantages of refactoring code

The main advantages of refactoring are the optimizaiton of execution time and reading instructions in one go. The disadvantage of refactoring could be following the flow of the code if comments are not actively used to explain what each instruction is intended to do. 

### Pros and cons of refactoring the original VBA script

The advantage of the original VBA script is creating instructions in chunks, or subroutines, with each subroutine accomplishing a specific task. If anything, the disadvantage of writing code in the original VBA script and then refactoring is that it can be a time consuming exercise. However, re-writing the code allows to review and reenforce how to write instructions on VBA. This is useful for beginners just learning to write instructions on VBA. 

