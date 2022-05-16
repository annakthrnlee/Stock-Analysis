# Stock-Analysis
## VBA Challenge
### Overview of Project:
The purpose of my analysis is to help Steve analyze different stock data covering a two year period. Steve’s goal is to help his parents decide which 
stocks are financially profitable so that they can determine which investments are smarter in the long-run. My analysis helps him see each stock's 
potential while also separating them by year (2017 and 2018). After reviewing my analysis, Steve can now help his parents invest in more reliable 
sources that have proven to be financially respectable over the past years. 
### Results:
Steve’s parents originally wanted to invest in Daqo’s Stock (DQ) and were bestering him about looking into past return rates. As we can see from 
my analysis, in 2018 DQ had a -63% return rate. Steve can now explain to his parents that even though they want to invest in a sustainable company, 
DQ may not be the best option based on that year's return rate. However, Steve could also use my analysis and run the same code but for 2017. 
In 2017, DQ was up by 199.4% which is an exceptional return rate compared to 2018. Based on these findings, it's up to Steve and his family to 
determine their individual risk based on the data provided.

#### Image: Representing 2018 DQ stock 
<img width="373" alt="2018 DQ return" src="https://user-images.githubusercontent.com/104043438/168684005-eba0077f-ff08-4ebf-ae16-9f705aa95dbf.png">

### Summary:
After completing the same analysis throughout module 2, it was convenient to have the ability to refactor the code and use it for my new analysis. 
While refactoring code does save a lot of time, it also comes with a price. Once the code has been pasted back into VBA, you must look through it 
and commit any changes necessary in order for it to run with your new project. The smallest thing will create an error, which I experienced multiple 
times throughout completing my analysis. The first error that kept recurring concerned the date. I attempted multiple times to enable a button to run 
the analysis for each year but I was unable to separate 2017 and 2018. Luckily the challenge didn’t require a button so I was able to avoid having to 
refactor my old code yet again. I added a clear button in my first analysis which helped me revert back to my original findings if anything went wrong. 

#### The code I used was: 
Sub Worksheets("All Stocks Analysis").Activate
Cells.Clear
End Sub

Another problem that could arise when refactoring code is bugs, sometimes your old code won’t do exactly what you anticipated in the new one so it’s 
very important to look over your code before running it. The link below is a great tool to decrease the chance of errors or bugs from occuring 
when you're refactoring code.

#### Helpful link: https://docs.microsoft.com/en-us/visualstudio/ide/find-and-fix-code-errors?view=vs-2022

