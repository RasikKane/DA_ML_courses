--Joins & Unions
--
--
--Cross Join :  all produts from all suppliers [WRONG]
SELECT product_name,
    unit_price,
    company
FROM suppliers
    CROSS JOIN produts;
--
--
--INNER JOIN :  all produts from all suppliers
SELECT product_name,
    unit_price,
    company
FROM suppliers
    INNER JOIN produts ON suppliers.supplierID = products.supplierID;
--
--
--INNER JOIN with Alias and Pre-qulaifiers:  all produts, price and manufacturer company
SELECT P.product_name,
    O.unit_price,
    S.company
FROM (
        (
            suppliers S
            INNER JOIN products P ON S.supplierID = P.supplierID
        )
        INNER JOIN orders O ON P.productID = O.productID;
);
--
--
--LEFT JOIN :  all customers and any orders they have placed
SELECT C.customer_name,
    O.orderID
FROM customers C
    LEFT JOIN orders O ON C.customerID = O.customerID;
--
--
--RIGHT JOIN :  list all employees of courier company itself who have placed orders as customers
SELECT E.employee_name,
    O.orderID
FROM orders O
    LEFT JOIN employee E ON O.customerID = E.employeeID;
--
--
--FULL OUTER JOIN :  list all employees of courier company itself who have placed orders as customers
SELECT E.employee_name,
    O.orderID
FROM orders O
    OUTER JOIN employee E ON O.customerID = E.employeeID;
--
--
--Union :  Select cities which have supplier office
SELECT city,
    country
FROM countries C
WHERE C.country = "India"
UNION
SELECT city,
    country
FROM suppliers S
WHERE S.country = "India"
ORDER BY City