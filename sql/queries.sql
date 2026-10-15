-- =====================================================================
-- queries.sql  |  Analysis queries run by run_pipeline.py on db/sales.db
-- Business rule: net revenue = order_value of orders with status 'Delivered'.
-- Each query starts with a "-- name:" marker so the pipeline can export it.
-- =====================================================================

-- name: q01_overall_kpis
-- Aggregations: headline KPIs for delivered orders.
SELECT COUNT(*)                          AS delivered_orders,
       ROUND(SUM(order_value), 2)        AS net_revenue,
       ROUND(AVG(order_value), 2)        AS avg_order_value,
       COUNT(DISTINCT customer_id)       AS active_customers
FROM orders
WHERE status = 'Delivered';

-- name: q02_monthly_revenue_mom
-- CTE + window function (LAG): monthly revenue with month-over-month growth.
WITH monthly AS (
    SELECT substr(order_date, 1, 7) AS month,
           ROUND(SUM(order_value), 2) AS revenue,
           COUNT(*)                   AS orders
    FROM orders
    WHERE status = 'Delivered'
    GROUP BY substr(order_date, 1, 7)
)
SELECT month, revenue, orders,
       ROUND(100.0 * (revenue - LAG(revenue) OVER (ORDER BY month))
             / LAG(revenue) OVER (ORDER BY month), 1) AS mom_growth_pct
FROM monthly
ORDER BY month;

-- name: q03_revenue_by_category
-- JOIN + GROUP BY + share of total using a window function.
SELECT p.category,
       ROUND(SUM(o.order_value), 2) AS revenue,
       ROUND(100.0 * SUM(o.order_value) / SUM(SUM(o.order_value)) OVER (), 1) AS share_pct
FROM orders o
JOIN products p ON p.product_id = o.product_id
WHERE o.status = 'Delivered'
GROUP BY p.category
ORDER BY revenue DESC;

-- name: q04_revenue_region_category
-- Three-table JOIN: revenue by region and category.
SELECT c.region, p.category,
       ROUND(SUM(o.order_value), 2) AS revenue,
       COUNT(*)                     AS orders
FROM orders o
JOIN customers c ON c.customer_id = o.customer_id
JOIN products  p ON p.product_id  = o.product_id
WHERE o.status = 'Delivered'
GROUP BY c.region, p.category
ORDER BY c.region, revenue DESC;

-- name: q05_top10_customers
-- Window function RANK(): top 10 customers by revenue.
SELECT *
FROM (
    SELECT c.customer_id, c.city,
           ROUND(SUM(o.order_value), 2) AS revenue,
           COUNT(*)                     AS orders,
           RANK() OVER (ORDER BY SUM(o.order_value) DESC) AS revenue_rank
    FROM orders o
    JOIN customers c ON c.customer_id = o.customer_id
    WHERE o.status = 'Delivered'
    GROUP BY c.customer_id, c.city
)
WHERE revenue_rank <= 10
ORDER BY revenue_rank;

-- name: q06_top3_products_per_category
-- ROW_NUMBER() OVER (PARTITION BY ...): top 3 products inside each category.
SELECT category, product_id, product_name, revenue, rank_in_category
FROM (
    SELECT p.category, p.product_id, p.product_name,
           ROUND(SUM(o.order_value), 2) AS revenue,
           ROW_NUMBER() OVER (PARTITION BY p.category ORDER BY SUM(o.order_value) DESC) AS rank_in_category
    FROM orders o
    JOIN products p ON p.product_id = o.product_id
    WHERE o.status = 'Delivered'
    GROUP BY p.category, p.product_id, p.product_name
)
WHERE rank_in_category <= 3
ORDER BY category, rank_in_category;

-- name: q07_repeat_customer_rate
-- CTE + subquery logic: share of customers with 2+ delivered orders.
WITH per_customer AS (
    SELECT customer_id, COUNT(*) AS orders
    FROM orders
    WHERE status = 'Delivered'
    GROUP BY customer_id
)
SELECT COUNT(*)                                            AS customers_with_orders,
       SUM(CASE WHEN orders >= 2 THEN 1 ELSE 0 END)        AS repeat_customers,
       ROUND(100.0 * SUM(CASE WHEN orders >= 2 THEN 1 ELSE 0 END) / COUNT(*), 1) AS repeat_rate_pct
FROM per_customer;

-- name: q08_running_total_revenue
-- Window function SUM() OVER: cumulative revenue by month.
WITH monthly AS (
    SELECT substr(order_date, 1, 7) AS month, SUM(order_value) AS revenue
    FROM orders
    WHERE status = 'Delivered'
    GROUP BY substr(order_date, 1, 7)
)
SELECT month,
       ROUND(revenue, 2) AS revenue,
       ROUND(SUM(revenue) OVER (ORDER BY month ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW), 2) AS cumulative_revenue
FROM monthly
ORDER BY month;

-- name: q09_cancellation_return_rate_by_region
-- Conditional aggregation: cancelled and returned share of all orders per region.
SELECT c.region,
       COUNT(*) AS total_orders,
       ROUND(100.0 * SUM(CASE WHEN o.status = 'Cancelled' THEN 1 ELSE 0 END) / COUNT(*), 1) AS cancelled_pct,
       ROUND(100.0 * SUM(CASE WHEN o.status = 'Returned'  THEN 1 ELSE 0 END) / COUNT(*), 1) AS returned_pct
FROM orders o
JOIN customers c ON c.customer_id = o.customer_id
GROUP BY c.region
ORDER BY cancelled_pct DESC;

-- name: q10_customers_without_orders
-- LEFT JOIN anti-join: registered customers who never placed a valid order.
SELECT c.customer_id, c.city, c.signup_date
FROM customers c
LEFT JOIN orders o ON o.customer_id = c.customer_id
WHERE o.order_id IS NULL
ORDER BY c.signup_date
LIMIT 20;
