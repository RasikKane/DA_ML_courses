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
-- Find the names of all the tracks for the album "Californication".
SELECT T.Name
FROM Tracks AS T
WHERE T.AlbumId IN (
        SELECT AlbumId
        FROM Albums AS A
        WHERE A.Title = "Californication"
    );
-- 
-- 
-- Find the total number of invoices for each customer along with the customer's full name, city and email.
SELECT C.FirstName,
    C.LastName,
    C.City,
    C.Email,
    COUNT(I.invoiceId) AS num_invoices
FROM Customers AS C
    LEFT JOIN Invoices AS I ON C.CustomerId = I.CustomerId
GROUP BY C.CustomerId;
-- 
--
-- Retrieve the track name, album, artistID, and trackID for all the albums 
SELECT T.Name,
    A.Title,
    A.ArtistId,
    T.trackId
FROM Albums AS A
    LEFT JOIN Tracks AS T ON A.AlbumId = T.AlbumId;
--
--
-- Retrieve a list with the managers last name, and the last name of the employees who report to him or her
SELECT M.LastName AS Manager surname,
    E.LastName AS Employee surname
FROM Employees E,
    Employees M
WHERE E.ReportsTo = M.EmployeeId;
-- 
-- 
--Find the name and ID of the artists who do not have albums.
SELECT Artists.Name,
    Artists.ArtistId
FROM Artists
    LEFT JOIN Albums ON Artists.ArtistId = Albums.ArtistId
WHERE Albums.Title IS NULL;
-- 
--
-- create a list of all the employee's and customer's first names and last names ordered by the last name in descending order
SELECT FirstName,
    LastName
FROM Employees
UNION
SELECT FirstName,
    LastName
FROM Customers
ORDER BY LastName  DESC;
-- 
--
-- See if there are any customers who have a different city listed in their billing city versus their customer city.
SELECT C.CustomerId, I.InvoiceId, C.City, I.BillingCity AS Billing_City
FROM Customers AS C
    LEFT JOIN Invoices AS I ON C.CustomerId = I.CustomerId
WHERE C.City <> I.BillingCity;
-- 
--
