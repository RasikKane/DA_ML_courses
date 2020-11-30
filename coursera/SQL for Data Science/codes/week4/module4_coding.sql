--all the questions in this practice set are based on Chinook data base. Schema are
--Artists
--Playlists
--Employees
--Albums
--Tracks
--Invoices
--Customers
-- 
-- 
-- Pull a list of customer ids with the customer’s full name, and address, 
-- along with combining their city and country together. 
-- Be sure to make a space in between these two and make it UPPER CASE. (e.g. LOS ANGELES USA)
SELECT C.CustomerId,
    C.Address,
    C.FirstName || ' ' || C.LastName,
    UPPER C.City || ' ' || ' ' || C.Country
FROM Customers AS C;
-- 
-- 
-- Create a new employee user id by combining the first 4 letters of the employee’s first name 
-- with the first 2 letters of the employee’s last name. 
-- Make the new field lower case and pull each individual step to show your work.
SELECT FirstName,
    LastName,
    LOWER(SUBSTR(FirstName, 1, 4)) AS A,
    LOWER(SUBSTR(LastName, 1, 2)) AS B,
    LOWER(SUBSTR(FirstName, 1, 4)) || LOWER(SUBSTR(LastName, 1, 2)) AS userId
FROM Employees -- 
    --
    -- Show a list of employees who have worked for the company for 15 or more years using the current date function. 
    -- Sort by lastname ascending
SELECT E.LastName,
    strftime('%Y', 'now') - strftime('%Y', E.HireDate)
FROM Employees AS E
WHERE (
        strftime('%Y', 'now') - strftime('%Y', E.HireDate)
    ) >= 15
ORDER BY E.LastName ASC;
--
--
-- Are there any columns with null values?
SELECT COUNT(*)
FROM Customers AS C
WHERE < columnName > IS NULL;
-- 
-- 
-- Find the cities with the most customers and rank in descending order.
SELECT C.City,
    COUNT(DISTINCT C.CustomerId) AS CustomerCount
FROM Customers AS C
GROUP BY C.City
ORDER BY CustomerCount DESC;
-- 
--
-- Create a new customer invoice id by combining a customer’s invoice id with their first and last name 
-- while ordering your query in the following order: firstname, lastname, and invoiceID.
SELECT C.FirstName,
    C.LastName,
    I.InvoiceId,
    C.FirstName || C.LastName || I.InvoiceId AS invID
FROM Customers AS C
    INNER JOIN Invoices AS I ON C.CustomerId = I.CustomerID -- WHERE invID LIKE "AstridGruber%"
ORDER BY C.FirstName,
    C.LastName,
    I.CustomerId -- 
    --