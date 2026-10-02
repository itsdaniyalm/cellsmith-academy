---
title: Counting and summing with conditions
summary: SUMIFS, COUNTIFS, AVERAGEIFS, MAXIFS and MINIFS - the workhorses of every report - with criteria, wildcards, dates and OR logic.
fx: =SUMIFS(Sales[Revenue], Sales[Region], "West", Sales[Status], "Closed")
---

"Total revenue for the West region." "Number of open orders." "Average order size in March." Most reporting questions are a total, count or average of the rows that meet some conditions. Excel has a family of functions built exactly for this, and they're among the most-used functions there are.

## SUMIFS

```formula
=SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)
```

The first argument is what to add up. Then come pairs: a range to check, and the condition it must meet. A row is included only if it meets **every** condition.

```formula
=SUMIFS(Sales[Revenue], Sales[Region], "West")
→ 2874.25
=SUMIFS(Sales[Revenue], Sales[Region], "West", Sales[Status], "Closed")
→ 2009.55
```

All the ranges must be the same size. With table columns that's automatic.

## The rest of the family

They all share the same pattern of range-and-condition pairs:

| Function | Returns, for matching rows |
|---|---|
| `SUMIFS(sum_range, ...)` | the total |
| `COUNTIFS(...)` | how many rows match (no sum range) |
| `AVERAGEIFS(average_range, ...)` | the average |
| `MAXIFS(max_range, ...)` | the largest value |
| `MINIFS(min_range, ...)` | the smallest value |

```formula
=COUNTIFS(Sales[Status], "Open")
→ 4
=AVERAGEIFS(Sales[Revenue], Sales[Region], "East")
→ 246.2333333
=MAXIFS(Sales[Units], Sales[Rep], "Grace")
→ 25
```

::: note
You'll also see `SUMIF`, `COUNTIF` and `AVERAGEIF` without the S. They take a single condition, and `SUMIF` puts the sum range *last*. The S versions do everything the old ones do, so use them even for one condition and you only have one argument order to remember.
:::

## Writing criteria

A criterion is either a value to match exactly, or text that starts with a comparison operator:

| Criterion | Matches |
|---|---|
| `"West"` | exactly West (ignoring case) |
| `25` | exactly 25 |
| `">=20"` | 20 or more |
| `"<>Cancelled"` | anything except Cancelled |
| `"<>"` | anything that isn't blank |
| `"="` | blank cells |

The operator goes **inside the quotes**. To compare against a value from a cell or a calculation, join the operator on with `&`:

```formula
=COUNTIFS(Sales[Units], ">=" & 20)
→ 13
=SUMIFS(Sales[Revenue], Sales[Revenue], ">=" & MinOrder)
→ 11287
```

The second formula adds up orders worth at least the minimum order value from the Settings sheet.

::: gotcha
`">=MinOrder"` (all inside the quotes) doesn't work: Excel would look for values greater than or equal to the *text* "MinOrder". Keep cell references and names outside the quotes and join them on.
:::

### Dates in criteria

Dates are numbers, so the same operators work. Build the date with `DATE` and join it on:

```formula
=SUMIFS(Sales[Revenue], Sales[Date], ">=" & DATE(2026, 4, 1), Sales[Date], "<" & DATE(2026, 7, 1))
→ 5046.25
```

That's revenue for April to June. Notice the same column appears twice, once for the start and once for the end. Using `<` with the first day of the *next* period is safer than `<=` with the last day, because it also works if your dates include times.

### Wildcards

In text criteria, `*` matches any run of characters and `?` matches exactly one:

```formula
=SUMIFS(Sales[Revenue], Sales[Product], "Desk*")
→ 4811.5
=COUNTIFS(Sales[Product], "*Pen*")
→ 7
```

`"Desk*"` matches Desk Lamp and Desk Chair. `"*Pen*"` matches anything containing "Pen" anywhere.

## OR conditions

Every condition in `SUMIFS` must be met, so the conditions are joined by AND. For OR, like "North **or** South", give the criterion as a list in curly braces and add up the results:

```formula
=SUM(SUMIFS(Sales[Revenue], Sales[Region], {"North", "South"}))
→ 5216.05
```

`SUMIFS` runs once for each item in the list, giving two totals, and `SUM` adds them together. Chapter C1 explains what those curly braces are.

## A results table in one formula

Point the criterion at a whole column of values and you get one answer per value. This lists each product's total units:

```formula
=SUMIFS(Sales[Units], Sales[SKU], Products[SKU])
→ {61; 100; 161; 21; 15; 117; 185; 22}
```

The result **spills** into eight cells, one per product, in the same order as the Products table. Chapter C1 is all about spilling.

## SUMPRODUCT

Before these functions existed, people used `SUMPRODUCT`, and you'll still find it in older workbooks. It multiplies arrays together and adds up the result:

```formula
=SUMPRODUCT(Sales[Units], Sales[Price])
→ 12199.25
```

Combined with the TRUE-is-1 trick from chapter B1, it can apply any condition, including ones `SUMIFS` can't express, like a condition on a calculation:

```formula
=SUMPRODUCT((MONTH(Sales[Date]) = 3) * Sales[Revenue])
→ 2708.15
```

That's March revenue: `SUMIFS` can't take `MONTH(...)` as a range, but `SUMPRODUCT` doesn't mind. In modern Excel, `SUM` does the same job, so `=SUM((MONTH(Sales[Date]) = 3) * Sales[Revenue])` works too.

## Key ideas

- `SUMIFS`, `COUNTIFS`, `AVERAGEIFS`, `MAXIFS` and `MINIFS` take pairs of range and condition; every condition must hold.
- Put comparison operators inside quotes, and join cell values or dates on with `&`: `">=" & DATE(2026, 4, 1)`.
- `*` and `?` are wildcards in text criteria.
- For OR, pass a list like `{"North", "South"}` and wrap in `SUM`.

## Exercises

Use sheet **B2** of the practice workbook.

:::: exercise
What is the total revenue from closed orders in the East region?
::: answer
```formula
=SUMIFS(Sales[Revenue], Sales[Region], "East", Sales[Status], "Closed")
→ 2927.85
```
:::
::::

:::: exercise
How many orders were placed in the first quarter (January to March 2026)?
::: answer
```formula
=COUNTIFS(Sales[Date], ">=" & DATE(2026, 1, 1), Sales[Date], "<" & DATE(2026, 4, 1))
→ 25
```
:::
::::

:::: exercise
What is the average revenue of orders that had a discount?
::: answer
```formula
=AVERAGEIFS(Sales[Revenue], Sales[Discount], ">0")
→ 201.0552632
```
:::
::::

:::: exercise
What was Hiro's smallest order, in units?
::: answer
```formula
=MINIFS(Sales[Units], Sales[Rep], "Hiro")
→ 3
```
:::
::::

:::: exercise
What is the total revenue of all orders for Furniture products? The Sales table has no Category column, so match the three furniture products by name: Monitor Stand, Desk Chair and Whiteboard.
::: answer
```formula
=SUM(SUMIFS(Sales[Revenue], Sales[Product], {"Monitor Stand", "Desk Chair", "Whiteboard"}))
→ 5410.05
```
Chapter B5 shows a better way: look up each order's category from the Products table.
:::
::::

:::: exercise
How many orders are not cancelled? Use `COUNTIFS` with a "not equal" criterion.
::: answer
```formula
=COUNTIFS(Sales[Status], "<>Cancelled")
→ 45
```
:::
::::
