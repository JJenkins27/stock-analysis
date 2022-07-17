# Analysis of Stock Data

## Overview of Project

An analysis of greenspace stock data and their returns to assist in investment choices. 

## Results

- In 2017, the majority of the greenspace stocks were performing well. DQ had a yearly return of 199.4%. 

![2017 Stock Data](https://user-images.githubusercontent.com/108373151/179380515-2e055979-2e18-4016-bebf-b2c6a798a64f.png)

- However, 2018 saw the opposite results in which most of the stocks were underperforming. 

![2018 Stock Data](https://user-images.githubusercontent.com/108373151/179380521-313095da-5668-4feb-aeea-6723fcb51465.png)

## Summary

- Using refactoring code has many advantages. The code is less complex, making it much easier to update and maintain. It is much faster to run, especially when datasheets are large. A disadvantage of refactoring is that it may introduce more bugs in the code. 

The inital run times for the code was this in 2017:
![VBA_Challenge_2017_OG](https://user-images.githubusercontent.com/108373151/179380529-c7811935-f9e8-42f7-b239-5a1cebb642b3.png)

and similar in 2018:
![VBA_Challenge_2018_OG](https://user-images.githubusercontent.com/108373151/179380539-3e655a35-fcef-4cb5-a97e-be8a6b527c11.png)

And the refactoring code sped up the execution time tremendously. For 2017 it was:
![VBA_Challenge_2017_Refactored](https://user-images.githubusercontent.com/108373151/179380547-1319656d-abc9-4819-a700-146a8885d901.png)

and similar in 2018:
![VBA_Challenge_2018_Refactored](https://user-images.githubusercontent.com/108373151/179380554-e8064a42-cece-4655-bdbf-cbb31b85eb7e.png)

- For example, when updating the old code to refactoring, I encountered some overflow errors because I neglected to update one area to the array. It was difficult to find the problem as the overflow error occurred on a formula line that wasn't where the error was. However, once the code was successfully updated, things ran a lot more smoothly and quickly.
