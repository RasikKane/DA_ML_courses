--Sub queries
-- Know region of each customer who had order with freight more than 100
SELECT
    CustomerID,
    Region
FROM
    Customers
WHERE
    CustomerID in (
        SELECT
            CustomerID
        FROM
            Orders
        WHERE
            Freight > 100
    );

-- Total order placed by every Customer
SELECT
    Customer_name,
    States,
(
        SELECT
            COUNT(*) AS Orders
        FROM
            Orders
        WHERE
            Orders.CustomerID = Customers.CustomerID
    )
    FROM Customers ORDER BY Customer_name;