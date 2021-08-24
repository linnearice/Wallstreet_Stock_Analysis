# VBA of Wallstreet
## Overview of Project
Our company Data Analysts, Inc. (“DAI”) has worked with Steve, a Wall Street stockbroker, to deliver a tool Steve can use to analyze a large array of stocks.  Steve needs this information to advise his clients.  Steve has provided DAI the data points needed to properly analyze a stock.  It is DAI’s mission to incorporate all measurements into the analysis and to meet Steve’s expectations of accuracy and expedience.

### Improved Deliverable
DAI delivered our first draft analysis to Steve and after our follow-up meeting it was determined DAI needed to improve the analysis to incorporate the potential to analyze more stocks at one time and to do it faster.
	 
## Results
To facilitate an improved (or ***“refactored”***) deliverable, DAI accomplished the following:
1.	**Results:**
![2017 and 2018 Screenshot for report](https://user-images.githubusercontent.com/35401581/130656419-7a9e1f11-9894-48f1-9c97-54cfb0378a34.png)
 	 
	The notes in the message boxes at bottom of the charts read:
	-	2017 – The code ran in 0.1640625 seconds for the year 2017
	-	2018 – The code ran in 0.1445313 seconds for the year 2018
	
2.	**Calculation Accuracy:**  While the coding for the analysis was refactored for efficiency the calculation results for Total Daily Volume and Return remained the same which was the goal.
3.	**Analysis:**  Steve is now able to provide his analysis.  In the charts above 2017 results showed 11 of the 12 stocks evaluated traded with favorable returns but many took a nosedive in 2018 where only 2 of the 12 stocks had favorable returns.  The stock DQ for example had a favorable return of 199.4% in 2017 but fell to a -62.6% return in 2018.  Since the majority of stocks in this portfolio declined sharply in 2018, perhaps an overall market downturn may be at play.   
4.	**Flexibility:**  DAI incorporated the ability to ***run more stocks*** by changing data points (i.e., variables) from single variables to arrays which allowed the number of stocks within the array to be more easily changed.  By doing so, Steve was able to analyze the current set of 12 stocks (set up in the Tickers array) or if needed add stocks to the array to calculate many more.  In this event, adjustments to the counts within the arrays can be adjusted to accommodate the stock additions.  

- An example of added code is:

'Create a variable for the stock (tickers) array index (tickerIndex) which allows for more easily assigned values to each stock by index number (for example:  the variable tickerVolumes(tickerIndex) for the DQ stock (which holds a tickerIndex=3) is 35,796,200 in 2017.)

            Dim tickerIndex As Integer
            
'Create three output arrays.  For each stock (12) there are 3 variables with 36 results as defined below: 

            Dim tickerVolumes(12) As Long
            Dim tickerStartingPrice(12) As Single
            Dim tickerEndingPrice(12) As Single

Note:  The number of stocks evaluated can be changed as mentioned by adding stock(s) to the Tickers array.  Thus, the indexes above can be changed from 12 to a new number, let’s say 15.  

‘A loop using tickerIndex is used to calculate total daily volumes by stock:  

		For i = 2 To RowCount
            'calculates volume by tickerIndex
	       tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value  and so on…

5.	**Reduced Elapsed Time:**  The newly developed arrays also allowed DAI to implement ***more streamlined code*** that produced the same output in far fewer iterations.  DAI was then able to take out the nested loop statements and this produced a dramatic decrease in iterations.  Specifically, the ***refactored code*** for 12 stocks used 3,037 iterations while the ***original code*** used 36,156 iterations to produce the same data - a dramatic decrease indeed.  2017 and 2018 have the same number of total records so the reduced number of iterations was the same from year to year.  The result of implementing the ***refactored code*** with fewer iterations is:

	a.	***2017 data run*** – the elapsed time refactored was 0.1640625 seconds compared to the original code at 0.882813 seconds.
	
	b.	***2018 data run*** - the elapsed time refactored was 0.1470032 seconds compared to the original code at 0.960938 seconds 


## Summary
DAI was successful in providing Steve an improved deliverable.  Steve can now be confident he can analyze and present performances on an array of stocks in a timely fashion to his clients.

DAI would like to point out; however, that one must be very careful in conducting a ***refactoring*** project.  Code often includes complex calculations or detailed presentations of data and if all factors are not taken into consideration, the refactoring can result in incorrect or misleading outcomes.  One must weigh the advantages and disadvantages to refactoring code.

The **advantages** of refactoring code may include:

-  As in Steve’s case, refactoring can reduce the elapsed time needed to run the code.  If you are dealing with 1,000’s of records with multiple data points this time savings can add up and become significant.
-	Time often equates to cost and your cost of running the data may be reduced.
-	The code is easier to read and understand which means any programmer can come back to it and more easily make needed changes as time goes by, etc.

The **disadvantages** of refactoring code may include:

-	Calculations to produce certain output is often very complex and precise.  What seems like improved code may change the way or order in which a calculation computes which can result in data errors or in some cases slight changes to the data that aren’t caught in a check of only a few data points.  However, when this data is run over 1000’s of iterations or more this slight change can add up resulting in imprecise outcomes and misdirected strategy.
-	While the cost to run a program may improve due to reduced elapsed times, the cost for resources to make the changes may be too high and those funds may not be available.
-	As in cost, the time it takes to complete a refactor may be excessive and a company may not have the manpower to devote to the project.

In summary, one must outline a detailed plan for a successful refactor and identify all pros and cons before deciding to move forward.
	

