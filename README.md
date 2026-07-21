# lifestyle-wellness-excel-analysis
An Excel analysis of how exercise, sleep quality, stress and daily activity relate to overall health scores.


## Project Overview

This is my first exploratory data analysis project using Microsoft Excel.

The project uses a synthetic lifestyle and wellness dataset obtained from Kaggle. I selected this topic because I have recently developed a greater personal interest in the benefits of regular exercise and believed it would be a meaningful subject for my first data analytics project.

In this hypothetical scenario, I work as a junior data analyst supporting a wellness program. The objective is to examine how exercise, physical activity, sleep and stress are associated with overall health scores.


### Business Objective

The purpose of this analysis is to identify lifestyle factors associated with higher overall health scores and present the results through clear comparisons and visualisations.


### Business Questions:

1.  Do people who exercise more frequently have higher average health scores?
2.	How do sleep duration and sleep quality relate to average health scores?
3.	How do average health scores differ across stress levels?
4.	Do people with higher daily step counts have higher average health scores?
5.	How do exercise frequency and stress level together compare with average health scores?


### Dataset
•	Source: https://www.kaggle.com/datasets/nalisha/early-wakeup-health-and-lifestyle-dataset/data

•	Primary tool: Microsoft Excel

The analysis is focused on the following columns:

* Person_ID
* Age
* Gender
* Sleep_Duration_Hours
* Sleep_Quality_Score
* Exercise_Frequency_Per_Week
* Daily_Steps
* Stress_Level
* Health_Score


### Excel Skills Used:

•	Excel Tables
•	Structured references
•	ROWS
•	COUNTA
•	UNIQUE
•	COUNTBLANK
•	IFS
•	PivotTables
•	PivotCharts
•	Conditional formatting
•	Heatmap visualisation
•	Dashboard design
•	Grouped comparisons
•	Descriptive analysis


## Project Process

### 1. Data Preparation and Validation
   
As the dataset was already well structured and required minimal cleaning, the main preparation work involved validating the records and creating analytical groups for easier comparison.

A separate worksheet was created to preserve the original data. The working dataset was converted into an Excel Table and renamed:
tbl_clean

Using an Excel Table made it easier to filter the data, apply formulas and use structured column references.

**Total record count:**
The following formula was used to confirm the number of participant records:
=ROWS(tbl_clean[Person_ID])
Result: 10,000 records

**Unique participant IDs:**
The following formula was used to confirm that each participant ID was unique:
=COUNTA(UNIQUE(tbl_clean[Person_ID]))
Result: 10,000 unique participant IDs.

**Missing-value checks:**
The COUNTBLANK function was used to check whether the variables contained missing values.
Example:
=COUNTBLANK(tbl_clean[Health_Score])


### 2. Data Processing
   
New columns were created to convert numerical values into clear categories by using IFS function. This made it easier to compare average health scores using PivotTables and charts.
1.	**Exercise frequency groups**
•	0 days
•	1–2 days
•	3–4 days
•	5–6 days
•	7 days
=IFS(
[@[Exercise_Frequency_Per_Week]]=0,"0 days",
[@[Exercise_Frequency_Per_Week]]<=2,"1-2 days",
[@[Exercise_Frequency_Per_Week]]<=4,"3-4 days",
[@[Exercise_Frequency_Per_Week]]<=6,"5-6 days",
TRUE,"7 days")

2.	**Daily-step groups**
•	Under 5,000
•	5,000–7,000
•	7,000–9,000
•	Over 9,000
=IFS(
[@[Daily_Steps]]<5000,"Under 5,000",
[@[Daily_Steps]]<=7000,"5,000-7,000",
[@[Daily_Steps]]<=9000,"7,001-9,000",
TRUE,"Over 9,000")

3.	**Sleep-quality groups**
•	Low
•	Below Average
•	Good
•	Excellent
=IFS(
[@[Sleep_Quality_Score]]<4,"Low",
[@[Sleep_Quality_Score]]<6,"Below Average",
[@[Sleep_Quality_Score]]<8,"Good",
TRUE,"Excellent")

4.	**Stress groups**
•	Low
•	Medium
•	High
=IFS(
[@[Stress_Level]]<4,"Low",
[@[Stress_Level]]<=7,"Medium",
TRUE,"High")

5.	**Sleep-duration groups**
•	Under 5 hours
•	5–6.9 hours
•	7–9 hours
•	Over 9 hours
=IFS(
[@[Sleep_Duration_Hours]]<5,"Under 5 hours",
[@[Sleep_Duration_Hours]]<7,"5-6.9 hours",
[@[Sleep_Duration_Hours]]<=9,"7-9 hours",
TRUE,"Over 9 hours")


### 3. Data Analysis
   
PivotTables were used to calculate and compare average health scores across the different lifestyle groups.

The analysis compared:
•	Average health score by exercise frequency
•	Average health score by sleep duration
•	Average health score by sleep quality
•	Average health score by stress level
•	Average health score by daily-step group


### 4. Data Visualisation
   
The results were presented using:
•	Bar and column charts for category comparisons
•	Line charts for ordered lifestyle groups
•	A conditional-formatting heatmap
The heatmap was used to compare the combined relationship between exercise frequency, stress level and average health score.


## Key Insights

•	**Daily steps and health score**

Higher daily-step groups recorded progressively higher average health scores.
Participants completing fewer than 5,000 daily steps recorded an average health score of approximately 64.5, compared with approximately 80.3 among participants completing more than 9,000 daily steps.

•	**Exercise frequency and health score**

Average health scores increased consistently as weekly exercise frequency increased.
Participants who did not exercise recorded an average health score of approximately 61.9, while participants exercising seven days per week recorded an average score of approximately 86.3.

•	**Sleep quality and sleep score**

Average health scores increased as sleep quality improved.
Participants in the low sleep quality group recorded an average health score of approximately 64.2, compared with approximately 83.7 for participants in the excellent sleep quality group.

•	**Stress level and health score**

Average health scores decreased as stress levels increased.
The low-stress group recorded an average health score of approximately 79.3, compared with approximately 64.6 for the high-stress group.

•	**Sleep quantity and sleep score**

Average health scores increased as sleep duration increased.
Participants who slept under five hours recorded an average health score of approximately 71.8, compared with approximately 76.7 for participants who slept more than nine hours.

•	**Combined exercise and stress comparison**

The heatmap shows that participants with frequent exercise and low stress recorded the highest average health scores.
Participants exercising seven days per week with low stress recorded an average health score of approximately 91.7, while participants reporting no exercise and high stress recorded an average score of approximately 53.0.


## Recommendations

1.	Encouraging regular weekly exercise through fitness or movement programs.
2.	Promoting daily physical activity, including walking and step-based goals.
3.	Providing stress-management resources alongside physical-activity initiatives.
4.	Focusing on sleep quality and healthy sleep routines.
5.	Considering combined wellness strategies because exercise and stress appear to have related effects on overall health scores.

   
## Limitations

•	The dataset is synthetic and may not accurately represent a real population.
•	The findings describe associations within the dataset and do not prove that one lifestyle factor causes a change in health.
•	The analytical groups were created for this project and were not provided as official clinical categories.


## Conclusion

This project demonstrates my ability to prepare data, validate records, create calculated categories, use PivotTables, compare groups and communicate findings through an Excel dashboard.
It also strengthened my understanding of how analytical questions can be translated into structured Excel analysis and clear visual insights.


**Sammy Rana**
Aspiring Data Analyst developing practical skills in Excel, SQL and data visualisation.







