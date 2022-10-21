# A VBA Challenge


## **Overview of Project**
A recent finance graduate, Steve, needed some stock data analysis for his parents. He provided the data to analyze in an excel sheet. In order to automate the analysis, the VBA was used. Initial script calculated quite quickly the total volume and yearly return for dozen of stocks and for last few years. It is, however, not sure whether the code would work on much larger dataset - thousands of stocks. It may take a long time to execute. 

### Purpose
The purpose of this analysis is to refactor a VBA code for stock data analysis and measure the performance. 


## **Results**
Original and refactored script can be found here:
[VBA_Challenge](https://github.com/MSF2141/stock-analysis/blob/df71324ba2be48367c9bef26d1e495ddf80cc07d/VBA_Challenge.zip)

Using images and examples of your code, compare the stock performance between 2018 and 2018, as well as the execution times of the original script and the refactored script.
### Analysis of Outcomes Based on Launch Date
Briefly, to analyze campaign outcomes based on their launch dates, the Kickstarter dataset was further processed to make the data more detailed. First, the “Category and Subcategory” column was split into two separate columns, “Parent category” and “Subcategory”. Second, the “Launched_at” and “Deadline” columns were converted from Unix timestamps (i.e., a measure of time in seconds since midnight of January 1, 1970) into a day-month-year format in new columns “Date Created Conversion” and “Date Ended Conversion”, respectively. Third, year was extracted from the “Date Created Conversion” column to a new “Years” column.  

To summarize the data, a pivot table was created from the KickStarter worksheet and filtered based on the “Parent Category” and “Years”. The "Parent Category" was then filtered to show only the data for "Theater" subcategory. For the Columns value the "Outcome" was selected. Column labels were then filtered to show only "Successful," "Failed," and "Canceled" campaigns. For the Rows value the "Date Created Conversion" was selected. The "Row Labels" column was grouped to show only the months of the year. In the Values box, the “Outcome" was selected and it appeared as "Count of outcome."	

To visualize the data, a pivot chart was created from the pivot table. To display campaign outcomes with relation to the launched month, as continuous data over time, a line with markers chart type was selected. 

Refactored VBA Code
![VBA_Challenge_2018_refactored](https://github.com/MSF2141/stock-analysis/blob/9f3766bcac3e77ec054fa10f6a1a70b20d7aaa3e/Resources/VBA_Challenge_2018_refactored.png)


Original VBA Code
![VBA_Challenge_2018](

### Analysis of Outcomes Based on Goals
Briefly, to analyze campaign outcomes based on their campaign goals, a summary table was created that contained a “Goal” column with several dollar-amount ranges so campaigns could be filtered based on their goal amount.  A COUNTIF() function was used to collect information from the Kickstarter dataset about the outcome and goal amount for the “Plays” subcategory and to fill the "Number Successful," "Number Failed," and "Number Canceled" campaign columns. Per each goal-amount range, “Total  Projects” was calculated using a SUM() function.  

Percentage of successful, failed, and canceled campaign outcomes was calculated by dividing "Number Successful", "Number Failed", and "Number Canceled" by “Total projects”, respectively. To display data as percentage, the format of  “Percentage successful”, “Percentage failed”, and “Percentage canceled” columns was changed from general to percentage.

To visualize the data as a relationship between the goal-amount ranges and percentage of successful, failed, or canceled campaigns, a line chart was created.
![Outcomes_vs_Goals](https://github.com/MSF2141/kickstarter-analysis/blob/c1ef0c78db6bd53c981d6a32ab45600eeef5841d/Resources/Outcomes_vs_Goals.png)



## **Summary**

    What are the advantages or disadvantages of refactoring code?
    How do these pros and cons apply to refactoring the original VBA script?

- What are two conclusions you can draw about the Outcomes based on Launch Date?
The month that launched the most successful Kickstarter campaigns is May, followed by June and July. These months also have similar numner of failed campaigns, which suggest that May is truely the most ideal month to launch a successful campaign.

- What can you conclude about the Outcomes based on Goals?
The most successful campaing goals are within the range of less than $1,000, followed by $1,000-$4,999 and $35,000 and $44,999 range. The third most successful range is substentially higher to the first two most successful ones suggesting that another factor can explain its success.
