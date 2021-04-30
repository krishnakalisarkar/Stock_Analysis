# Stock_Analysis

**Title: VBA Analysis of Wall Street Stock Market in 2017 and 2018**

**Purpose of this Analysis:**

The main purpose of this analysis is to get an insight into the stock market returns of few companies in two consecutive years namely 2017 and 2018. Based on the output of this analysis, it is easier for the stakeholders to decide on the fate of their stocks, if they want to keep the stocks or tradeoff for a better blooming one. This analysis would also be a guide for future investors to decide on which company to invest their money on. By comparing few consecutive years gives a clear picture of stock trends over the same period.

**Results:**
**Analysis for Stock returns of 2017:**
   
![Figure:1](Resource/Succesful_backers.png)
Looking at the above Excel sheet and the bar graph visualization of the same data, it is clearly evident that as far as overall market goes most of the stocks did fairly well in 2017. Some companies had a much higher percentage returns as compared to others.

**Positive returns in percentage 2017**
	**Above 100% annual return:**
The top 3 companies with high returns are DQ with 199.4%, followed by SEDG at 184.5% and ENPH with 129.5%. FSLR had a fairly good return with 101.3%. 
	**50% to 100% annual return:**
Two companies VSLR (50%) and JKS (53.9%) had an annual return between 50 % to 100% in 2017. 
**Less than 50% annual return:**
The companies trailing at less than 50% returns are CSIQ at 33.1%, HASI at 25.8% and SPWR at 23%. 

**Negative returns in percentage 2017**

The company that failed to leave a positive mark at the stock market in 2017 is TERP, trailing behind at -7.2% return.

**Analysis for Stock returns of 2018:**

  
![Figure:2](Resource/Failed_backer.png)
Looking at the above Excel sheet of 2018 returns and the bar graph visualization of the same data, it is shocking to see that most of the stocks failed miserably in 2018.
**Positive returns in percentage 2018**

	**Above 100% annual return:**
There were no companies with an annual return of above 100%.
	**50% to 100% annual return:**
The two companies standing tall in this market are ENPH with 81.9% and RUN with a return of 84%.
**Less than 50% annual return:**
There are no other companies who made a mark in 2018.

**Negative returns in percentage 2018**
Most of the companies had a negative annual return in 2018. The worst hit is DQ with a return of -62.6% and JKS with -60.5%. This is followed by SPWR with -44.6%, FSLR with – 39.7%, HASI at -20.7% and CSIQ at -16.3%.
The other companies that had a low return are VSLR (-3.5%), TERP (-5.0%), AY (7.3) and SEDG at (7.8%).

**Comparing each company’s stock returns in 2017 and 2018:**
 
![Figure:3](Resource/Failed_backer.png)
By comparing the stocks in 2017 and 2018 some interesting trends come out in light. 
•	Of all the stocks analyzed only RUN stocks jumped from a mere 5.4% return in 2017 to a whooping 84.0% return 2018. Apart from this all-other stock experienced a fall.
•	The companies that experienced huge fall in returns are DQ from 199.4% in 2017 to -62.6% in 2018) and SEDG from 184.5% in 2017 to -7.8% in 2018. Interestingly, these two companies that had almost 200% return in 2017 and in 2018 they are experiencing negative returns.
•	Although, ENPH experienced loss in return from 2017 to 2018, yet it managed to have a positive return in 2018, as compared to all other stocks which had a negative return.

Execution time:
![Figure:4](Resource/Failed_backer.png)
•	The execution time for my refactored code for 2017 and 2018 both ran for 0.0703125 seconds.
  
![Figure:5](Resource/Failed_backer.png)
•	The execution time for the original code for years 2017 and 2018 was 0.46875 seconds. 
•	The refactored code execution time is much shorter than the original code execution time.

**Summary:**
**1.Pros and Cons of Refactoring Code:**
*"Refactoring is a disciplined technique for restructuring an existing body of code, altering its internal structure without changing its external behavior."*
-	*Martin Fowler*

Code Refactoring is a way of restructuring and optimizing existing code without changing its behavior. It is a way to improve the code quality. 
**Pros of Refactoring:**
•	The main goal of code refactoring is to make it easy to enhance and maintain system performance in the future.
•	Remove hard codes and make it open.
•	Remove code smells like duplicate codes, lack of design and structure.
•	Makes it easy to understand in the future by having more messages in between codes.
•	Reduces code size and shorter run time.
•	Fixes bugs.
**Cons of Refactoring:**
*"You shouldn't refactor if a deadline is near.” Says Martin Fowler*
•	Code refactoring is time-consuming. It takes anywhere many hours to update the existing codes in small projects. Big projects require even longer hours which is disadvantageous.
•	If big changes and modifications are needed to the system’s structure, it is easier to build new code from scratch.
•	Refactoring can mess up with the existing code and the code can break apart, taking longer hours to fix that.
**2.How do these pros and cons apply to refactoring the original VBA script?**
•	**Pros:**
	The execution time was shorter 0.46875 seconds. 
	The code size is shorter, removed duplicated codes.
	The code is more structured and well designed with messages.
	Easy to maintain and understand in the future.
•	**Cons:**
	Took longer time to clean up.
	The code would break from time to time and debugging took a long time.



