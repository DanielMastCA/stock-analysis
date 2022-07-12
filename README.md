In this challenge, you’ll edit, or refactor, the Module 2 solution code to loop through all the data one time in order to collect the same information that you did in this module. 
Then, you’ll determine whether refactoring your code successfully made the VBA macro run faster.
Finally, you’ll present a written analysis that explains your findings.


# Stock Analysis Refactoring

# Overview of Project
Our client, Steve, is in need of a software to quickly analyse stocks. Using VBA and Excel we have presented Steven with a fully working analysis tool utilizing data he provided. Steve is wondering how well this program will perform against larger databases containing all stock market data from spesific years. In this analysis we take a look at our macro and refactor the code to see if we can make it more efficient.

# Results

### Refactoring process
In the first version of the macro, we loop over each individual ticker, find the values in the sheet for that ticker, and save the results and so forth. 
```VBA
For i = 0 To UBound(tickers)

    For j = 2 To lastRow
    
        'code
    
    Next j
    
Next i
```

This proves to be a highly inefficient method as it loops over the whole data set several times until all the tickers are filled. Instead we went with a simpler and more efficient solution of looping through the data set once and gathering all the data.

```VBA
For i = 2 To RowCount
    
    If Cells(i - 1, 1) <> tickers(tickerIndex) Then
    
        'code
    
    End If

Next i
```

## Refactoring was a success
After refactoring our code, we recorded the time it took our macro to provide us with the results Steve was looking for. As a result of this refactoring, our macro became significantly faster!

With the data Steve provided, we ran tests on both his 2017 data and his 2018 data.
#### 2017
![Pre Refactoring Analysis 2017](/Resources/VBA_Challenge_2017_Pre_Refactor.png)
![Post Refactoring Analysis 2017](/Resources/VBA_Challenge_2017.png)

#### 2018
![Pre Refactoring Analysis 2018](/Resources/VBA_Challenge_2018_Pre_Refactor.png)
![Post Refactoring Analysis 2018](/Resources/VBA_Challenge_2018.png)

As we can see, in both cases the macro ran significantly faster, cutting the processing time by over 600ms! In conclution, code refactoring is an essential step in developing that can make your code more readable and increase performance.

# Summary
### What are the advantages or disadvantages of refactoring code?
Refactoring code makes the code easier to understand, more efficient, and easier to maintain, but refactoring code can also be a tedious process, both time and mentality wise. You can also lose track of what to do, or break something.

### How do these pros and cons apply to refactoring the original VBA script?
Refactoring the original VBA script made the analysis faster and increased the script readability, but the code refactoring was timeconsuming. There were some bugs that needed to be solved along the way, and the additional arrays added into the refactored code made it slightly harder to comprehend. 
