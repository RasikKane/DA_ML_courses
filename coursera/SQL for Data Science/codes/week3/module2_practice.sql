--all the questions in this practice set, you will be using the Salary by Job Range Table.
--This is a single table titled: salary_range_by_job_classification. This table contains the following columns:
--SetIDJob_Code
--Eff_Date
--Sal_End_Date
--Salary_setID
--Sal_Plan
--Grade
--Step
--Biweekly_High_Rate
--Biweekly_Low_Rate
--Union_Code
--Extended_Step
--Pay_Type

-- Find the distinct values for the extended step.
SELECT DISTINCT Extended_step
FROM salary_range_by_job_classification

-- Excluding $0.00, what is the minimum bi-weekly high rate of pay
SELECT
MIN(Biweekly_high_Rate)
FROM salary_range_by_job_classification
WHERE Biweekly_high_Rate <> "$0.00"

--What is the maximum biweekly high rate of pay
SELECT
MAX(Biweekly_high_Rate)
FROM salary_range_by_job_classification

--What is the pay type for all the job codes that start with '03'?
SELECT Job_Code, Pay_Type
FROM salary_range_by_job_classification
WHERE Job_Code LIKE "03%"

--Run a query to find the Effective Date (eff_date) or Salary End Date (sal_end_date) for grade Q90H0?
SELECT eff_date, sal_end_date
FROM salary_range_by_job_classification
WHERE grade = "Q90H0"

--Sort the Biweekly low rate in ascending order.
SELECT Biweekly_Low_Rate
FROM salary_range_by_job_classification
ORDER BY Biweekly_Low_Rate ASC

--What Step are Job Codes 0110-0400?
SELECT DISTINCT Step
FROM salary_range_by_job_classification
WHERE Job_Code BETWEEN  "0110" AND "0400"

--What is the Biweekly High Rate minus the Biweekly Low Rate for job Code 0170?
SELECT (Biweekly_High_Rate - Biweekly_Low_Rate)
FROM salary_range_by_job_classification
WHERE Job_Code = "0170"

-- What is the Extended Step for Pay Types M, H, and D
SELECT DISTINCT Extended_Step
FROM salary_range_by_job_classification
WHERE Pay_Type in ("M","D","H")

--What is the step for Union Code 990 and a Set ID of SFMTA or COMMN?
SELECT DISTINCT Step
FROM salary_range_by_job_classification
WHERE Union_Code = "990" AND SetID IN ("SFMTA", "COMMN")
