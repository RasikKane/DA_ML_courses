--all the questions in this practice set are based on Chinook data base. Schema are
--Artists
--Playlists
--Employees
--Albums
--Tracks : [TrackId, name, AlbumId, MediaTypeId, GenreId, Composer, Milliseconds, Bytes, UnitPrice]
--Invoices
--Customers

-- Find all the tracks that have a length of 5,000,000 milliseconds or more.
SELECT TrackId
FROM Tracks
WHERE Milliseconds >= 5000000

--Find all the invoices whose total is between $5 and $15 dollars.
SELECT InvoiceId
FROM Invoices
WHERE Total  BETWEEN  5 AND 15

--Find all the customers from the following States: RJ, DF, AB, BC, CA, WA, NY.
SELECT *
FROM Customers
WHERE State in ("RJ", "DF", "AB", "BC", "CA", "WA", "NY")

-- Find all the invoices for customer 56 and 58 where the total was between $1.00 and $5.00.
SELECT *
FROM Invoices
WHERE CustomerId in (56, 58) AND Total BETWEEN 1 AND 5

--Find all the tracks whose name starts with 'All'.
SELECT TrackId
FROM Tracks
WHERE Name LIKE "All%"

--Find all the customer emails that start with "J" and are from gmail.com.
SELECT Email
FROM Customers
WHERE Email LIKE "J%gmail.com"

--Find all invoices from the billing city Brasília, Edmonton, and Vancouver and sort in descending order by invoice ID.
SELECT *
FROM Invoices
WHERE BillingCity in ("Brasília", "Edmonton", "Vancouver")
ORDER BY InvoiceId DESC

--Show the number of orders placed by each customer
--(hint: this is found in the invoices table) and sort the result by the number of orders in descending order)
SELECT CustomerId, COUNT(DISTINCT InvoiceId) AS number_orders
FROM Invoices
GROUP BY CustomerId
ORDER BY number_orders DESC

--Find the albums with 12 or more tracks.
SELECT AlbumId, COUNT(DISTINCT TrackId) AS number_tracks
FROM Tracks
GROUP BY AlbumId
HAVING number_tracks >= 12