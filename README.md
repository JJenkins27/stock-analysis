# Analysis of Stock Data

Performing analysis on stock prices in 2017 and 2018.

## Overview of Project

An analysis of greenspace stock data and their returns to assist in investment choices. 

### Purpose

Steve's parents are interested in investing in a company that utilizes greenspace. They invested in DAQO without fully investigating the risks. They are now wondering if they have made a wise investment choice. It would also be wise to consider diversifying their portfolio in other greenspace stocks.

## Results

- VBA Macros were used to compile an analysis of stock returns data. The macro compares the close price for the first day of the year to the last day of the year without considering all the normal ups and downs throughout the year of a typical stock price. This would help determine the yearly return of the stock.

- In 2017, the majority of the greenspace stocks were performing well. In particular, DQ had a yearly return of 199.4%, which made it the best investment choice for the year. Other stocks performing well were SEDG at 184.5% and ENPH at 129.5%.

![2017 Stock Data](https://user-images.githubusercontent.com/108373151/179380515-2e055979-2e18-4016-bebf-b2c6a798a64f.png)

- However, 2018 saw the opposite results in which most of the stocks were underperforming. DQ had a yearly return of -62.6% which was the most underperforming stock of the year. Similarly, SEDG decreased by 7.8% while ENPH actually had a positive return of 81.9%.

![2018 Stock Data](https://user-images.githubusercontent.com/108373151/179380521-313095da-5668-4feb-aeea-6723fcb51465.png)

- I would recommend to Steve's parents to invest in ENPH since they had consecutive positive years over the reported 2 years. However, this is no guarantee of future returns. 

- I would not recommend investing in TERP, which saw a negative return in both years: -7.2% in 2017 and -5.0% in 2018.

- The results convey that no stock is the best choice over time. To be a proper investment, one must consistently look at the performance of the company. There is never a guarantee that a stock that is performing well in one year will continue to perform well over time.

### Challenges and Difficulties Encountered

The biggest challenge of this report was refactoring the data. I had a hard time with getting the code right in the for loops but especially that 4th condition where you have to increase the ticker index if the next row doesn't match. However, once the code worked, it was easy to see how much faster it was from the original code. This would be especially beneficial with a larger dataset. 

We are only given 2 years of stock data. I would prefer to compare a longer length of time to really give a good investment recommendation.

## Summary

### General Advantages of Refactoring Code

- Using refactoring code has many advantages. The code is less complex, making it much easier to read, update and maintain. It is has a much better performance rate, especially when datasets are large. An example of this difference is in the screentimes below. 

The inital run times for the VBA code was this in 2017:

![VBA_Challenge_2017_OG](https://user-images.githubusercontent.com/108373151/179380529-c7811935-f9e8-42f7-b239-5a1cebb642b3.png)

and similar in 2018:

![VBA_Challenge_2018_OG](https://user-images.githubusercontent.com/108373151/179380539-3e655a35-fcef-4cb5-a97e-be8a6b527c11.png)

And the refactoring code sped up the execution time tremendously. For 2017 it was:

![VBA_Challenge_2017_Refactored](https://user-images.githubusercontent.com/108373151/179380547-1319656d-abc9-4819-a700-146a8885d901.png)

and similar in 2018:

![VBA_Challenge_2018_Refactored](https://user-images.githubusercontent.com/108373151/179380554-e8064a42-cece-4655-bdbf-cbb31b85eb7e.png)

I created one button for the old code and one button for the new refactoring code and the results even ran visibly faster to the naked eye. 

### General Disadvantages of Refactoring Code

- A disadvantage of refactoring is that it is time consuming to convert old VBA code to new refacotoring code. 

### Comparison of VBA Code to Refactoring Code

- The original code was pretty straightforward and easy to program. It didn't take a lot of time and was easy to understand.

- The original VBA code had a nested For loop which why it takes so much longer to compute. The nested For loop would circle back through the data many times to output the results, thus greatly increasing the amount of memory and time it takes to run. In the refactored code, the nested For loop is taken away, which greatly impacted the speed as shown in the screenshots above.

- If there needs to be updates to the VBA code or if the dataset is large, running the old VBA code can be time and memory cumbersome. I spent a lot of time updating the code to refactoring and getting the formulas right. The benefit I had during updating is that I had good code that I know worked and it was easy in a lot of places to update this to the array/refactoring code. 

- When updating the VBA code, I encountered some overflow errors because I neglected to update one area to the array. It was difficult to find the problem as the overflow error occurred on a formula line that wasn't where the error in my code was - a common problem with VBA programming. However, once the code was successfully updated, things ran a lot more smoothly and quickly.

- It was very helpful to have the old code to compare the results to, so I knew when my refactored code was working correctly.
