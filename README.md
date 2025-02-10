# Data-Analysis
### DESCRIPTION:
This program reads an Excel spreadsheet containing student performance data, cleans and organizes the data, conducts analysis, and generates
visualizations to highlight key patterns and insights.

### Data Cleaning
 1. Student Column: Renamed 'student' column to 'Student' & Standardized column data to ensure consistent capitalization (Title case)
 2. Study Hours Column: Renamed 'studyhours' column to 'Study Hours' & Deleted rows containing outliers (negative values) in the Study Hours column
 3. Attendance Rate Column: Renamed 'attendance_rate' column to 'Attendance Rate in %'
 4. Homework Column: Renamed 'homework' column to 'Homework (20 marks)'
 5. Participation Score Column: Renamed 'participationscore' column to 'Participation Score (15 marks)'
 6. Midterm Score Column: Renamed 'previousscore' column to 'Midterm Score (30 marks)'
 7. Final Exam Score Column: Renamed 'final_exam_score' column to 'Final Exam Score (35 marks)'
 8. Status Column: Renamed 'pass_fail' column to 'Status', Modified 0s to Fail & 1s to Pass
 9. Duplicates: Removed duplicates
 10. New Total Score Column: Created a 'Total Score' column to hold the sum of Homework, Participation, Midterm, and Final Exam
 11. Save: Saved the cleaned student performance excel spreadsheet to the same folder as the original spreadsheet 

### Data Analysis & Data Visualization
 1. Set Up Report: Opened a Word document and added a report title.
 2. Loaded Data: Read the cleaned Excel file with student performance data.
 3. Overall Statistics: Calculated summary statistics (using describe()) for the Total Score and added them to the word document.
 4. Histogram: Created and saved a histogram (with 20 bins and a density curve) to show the distribution of Total Scores, then inserted it
    into the word document.
 5. Boxplot (Total Score): Generated a boxplot comparing Total Scores between Pass and Fail groups, saved, and added it to the report.
 6. Countplot: Made a countplot (using pastel colors) to show the distribution of Pass/Fail statuses, saved and added it to the report.
 7. Study Hours Correlation: Computed the correlation between Study Hours and Total Score and added that number to the word document.
 8. Scatter Plot: Created a scatterplot showing the relationship between Study Hours and Total Score, then saved and inserted it into the word document.
 9. Boxplot (Study Hours): Generated a boxplot to compare Study Hours by Pass/Fail status, saved it, and inserted it.
 10. Participation Score Correlation: Calculated the correlation between Participation Score and Total Score and added it to the report.
 11. Boxplot (Participation Score): Created a boxplot comparing Participation Scores by Pass/Fail status, saved and inserted it into the word document.
 12. Pivot Table: Built a pivot table showing the average Total Score for each Study Hour value grouped by Status, and added it to the document.
 13. Final Save: Saved the Word document containing all the analysis and visualizations.
