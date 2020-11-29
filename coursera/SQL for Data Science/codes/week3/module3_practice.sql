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
-- How many albums does the artist Led Zeppelin have?
SELECT COUNT(*) AS count_albums
FROM albums
WHERE albums.ArtistId IN (
        SELECT ArtistId
        FROM artists
        WHERE artists.Name = "Led Zeppelin"
    );
-- 
-- 
-- Create a list of album titles and the unit prices for the artist "Audioslave".
SELECT A.Title,
    T.UnitPrice
FROM (
        albums A
        LEFT JOIN tracks T ON A.AlbumId = T.AlbumId
    )
WHERE A.ArtistId IN (
        SELECT ArtistId
        FROM artists
        WHERE artists.Name = "Audioslave"
    );
-- 
--
-- Find the first and last name of any customer who does not have an invoice. 
-- Are there any customers returned from the query? Ans --> NO
SELECT C.FirstName,
    C.LastName
FROM customers C
    LEFT JOIN invoices I ON I.CustomerId = C.CustomerId
WHERE C.CustomerId NOT IN (I.CustomerId);
--
--
-- Find the total price for each album.
SELECT A.Title, SUM(T.UnitPrice)
FROM (
        albums A
        LEFT JOIN tracks T ON A.AlbumId = T.AlbumId
    )
group by A.AlbumId;
-- 
-- 
-- How many records are created when you apply a Cartesian join to the invoice and invoice items table?
SELECT *
FROM invoices
    CROSS JOIN invoice_items;
-- 
-- 