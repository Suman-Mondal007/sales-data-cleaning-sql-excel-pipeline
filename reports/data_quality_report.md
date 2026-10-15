# Data Quality Report

_Generated 2026-10-15 11:28_

## Row counts (before -> after)

| Table | Raw rows | Clean rows |
|---|---:|---:|
| orders | 20,600 | 19,078 |
| customers | 6,040 | 6,000 |
| products | 60 | 60 |

## Issues found and how they were handled

| Issue | Count | Action |
|---|---:|---|
| Duplicate order rows (after trimming IDs) | 600 | Removed, kept first |
| Duplicate customer rows | 40 | Removed, kept first |
| Status label variants | 16 -> 4 | Standardised |
| Payment label variants | 16 -> 5 | Standardised |
| City label variants | 31 -> 8 | Standardised, mapped to region |
| Category label variants | 14 -> 6 | Standardised |
| Prices with currency symbols / commas | 2,025 | Parsed to numbers |
| Missing unit prices | 400 | Filled from product master |
| Missing product master prices | 2 | Filled with most common order price |
| Missing discounts | 600 | Assumed 0% |
| Missing customer emails | 180 | Left blank |
| Invalid customer emails | 234 | Set to blank |
| Mixed date formats | 3 formats | Parsed to ISO (YYYY-MM-DD) |

## Rows quarantined (922 orders -> `data/rejected/orders_rejected.csv`)

| Reason | Rows |
|---|---:|
| missing_quantity | 291 |
| quantity_out_of_range | 219 |
| invalid_or_unparseable_date | 100 |
| unknown_customer_id | 99 |
| unknown_product_id | 96 |
| date_out_of_range | 60 |
| discount_out_of_range | 57 |

## Post-cleaning validation

| Check | Result |
|---|---|
| order_id is unique | PASS |
| customer_id is unique | PASS |
| product_id is unique | PASS |
| no nulls in key order columns | PASS |
| quantity within 1-50 | PASS |
| unit_price > 0 | PASS |
| discount within 0-50% | PASS |
| every order has a valid customer | PASS |
| every order has a valid product | PASS |
| order dates within 2024-01-01..2025-12-31 | PASS |
| order_value = qty x price x (1 - discount) | PASS |
| status values are standardised | PASS |
| payment methods are standardised | PASS |
| all cities mapped to a region | PASS |

**14/14 checks passed.**

## Assumptions

- Dates containing `/` are read as DD/MM/YYYY (Indian convention).
- Net revenue counts only orders with status `Delivered`.
- A missing discount means no discount (0%).
- Quantity must be 1-50; discount must be 0%-50%; dates must fall in 2024-01-01 to 2025-12-31.