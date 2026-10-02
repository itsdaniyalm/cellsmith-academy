---
title: Making decisions
summary: IF, AND, OR, IFS and SWITCH, plus the trick that makes Excel's logic so powerful - TRUE and FALSE are numbers you can add and multiply.
fx: =IF(SUM(Sales[Revenue]) >= 12000, "Target met", "Keep going")
---

So far every formula has calculated one fixed thing. Real work needs formulas that choose: charge shipping only under a minimum order, flag a late delivery, label a region "on track" or "behind". This chapter covers the functions that make decisions.

## IF

`IF` asks a yes-or-no question and returns one value for yes and another for no:

```formula
=IF(logical_test, value_if_true, value_if_false)
```

The test is anything that produces `TRUE` or `FALSE`, usually a comparison (chapter A3):

```formula
=IF(SUM(Sales[Revenue]) >= 12000, "Target met", "Keep going")
→ Keep going
```

The results can be anything: text, numbers, references, or more calculations. Only the branch that's chosen gets used.

```formula
=IF(SUM(Sales[Revenue]) >= 12000, "Target met", "Short by " & TEXT(12000 - SUM(Sales[Revenue]), "#,##0.00"))
→ Short by 216.20
```

::: gotcha
Without the `TEXT`, this formula returns `Short by 216.200000000001`. Computers store decimal numbers in binary, so a calculation like 12000 − 11783.8 can come out a hair off. Excel normally hides that by showing 15 significant digits at most, but joining a number into text with `&` exposes it. Format numbers with `TEXT`, or `ROUND` them, before joining them into a message.
:::

If you leave out `value_if_false` and the test fails, `IF` returns `FALSE`, which is rarely what you want on a report. Give both branches.

::: tip
To return a blank-looking cell, use empty text: `""`. For example `=IF(B2 = 0, "", C2 / B2)` shows nothing instead of a `#DIV/0!` error when B2 is zero.
:::

## AND, OR and NOT

To test several conditions at once, combine them:

| Function | TRUE when |
|---|---|
| `AND(a, b, ...)` | every condition is TRUE |
| `OR(a, b, ...)` | at least one is TRUE |
| `NOT(a)` | a is FALSE |
| `XOR(a, b, ...)` | an odd number are TRUE |

```formula
=AND(COUNTA(Sales[Order ID]) = 48, SUM(Sales[Units]) > 500)
→ TRUE
=OR(MAX(Sales[Discount]) > 0.2, MIN(Sales[Units]) = 1)
→ TRUE
```

They're usually the test inside an `IF`. Here's a shipping rule: free shipping if the order is at least the minimum, or if the customer is in the North region. In a table row it would look like this:

```formula
=IF(OR([@Revenue] >= MinOrder, [@Region] = "North"), 0, 9.95)
```

## Nested IFs, and IFS

For more than two outcomes, you can put an `IF` inside another `IF`'s false branch:

```formula
=IF(AVERAGE(Sales[Units]) > 20, "High", IF(AVERAGE(Sales[Units]) > 10, "Medium", "Low"))
→ Medium
```

This gets hard to read quickly. `IFS` takes pairs of test and result, and returns the result for the **first** test that's TRUE:

```formula
=IFS(AVERAGE(Sales[Units]) > 20, "High", AVERAGE(Sales[Units]) > 10, "Medium", TRUE, "Low")
→ Medium
```

The last pair uses `TRUE` as its test, so it always matches. That's the "otherwise" branch. Without it, `IFS` returns `#N/A` when nothing matches.

::: gotcha
Order matters. Tests are checked from first to last and the first TRUE one wins. If you put `> 10` before `> 20`, a value of 25 would be labelled "Medium".
:::

## SWITCH

When you're comparing one value against a list of exact matches, `SWITCH` is clearer than either:

```formula
=SWITCH(value, match1, result1, [match2, result2], ..., [default])
```

```formula
=SWITCH(MAX(Sales[Discount]), 0.05, "Small", 0.1, "Medium", 0.15, "Large", "None")
→ Large
```

The optional last argument is the default, used when nothing matches.

## TRUE and FALSE are numbers

This is the most useful idea in the chapter. In arithmetic, `TRUE` counts as 1 and `FALSE` as 0. A comparison applied to a whole column gives a column of TRUEs and FALSEs, so adding them up **counts** the matches:

```formula
=SUM(--(Sales[Region] = "West"))
→ 9
```

The `--` (two minus signs) turns TRUE/FALSE into 1/0. The first minus makes them −1/0, the second flips them back. You'll also see `*1` or `+0` used the same way.

Multiplying two conditions gives 1 only where **both** are TRUE, so `*` works as AND:

```formula
=SUM((Sales[Region] = "West") * (Sales[Status] = "Closed"))
→ 6
```

Adding conditions gives a number above 0 where **either** is TRUE, so `+` works as OR. Wrap it in a comparison so rows that match both aren't counted twice:

```formula
=SUM(--((Sales[Region] = "North") + (Sales[Region] = "South") > 0))
→ 24
```

Multiply a condition by a column of numbers and you keep only the matching numbers. This sums revenue from closed orders:

```formula
=SUM((Sales[Status] = "Closed") * Sales[Revenue])
→ 9919.7
```

You'll meet functions designed for this (`SUMIFS`, `COUNTIFS`) in the next chapter. But `*` for AND and `+` for OR are exactly what `FILTER` uses for multiple conditions, so this idea pays off again in part C.

## IF works on whole columns too

`IF` also works element by element. This returns each order's revenue if it's closed and 0 otherwise, then sums them:

```formula
=SUM(IF(Sales[Status] = "Closed", Sales[Revenue], 0))
→ 9919.7
```

## Testing what kind of value you have

The `IS` functions answer TRUE or FALSE about a value:

| Function | TRUE when the value is |
|---|---|
| `ISBLANK` | an empty cell |
| `ISNUMBER` | a number (including dates) |
| `ISTEXT` | text |
| `ISERROR` | any error |
| `ISNA` | the `#N/A` error |

## Catching errors: IFERROR and IFNA

`IFERROR(value, value_if_error)` returns the value, unless it's an error, in which case it returns your alternative:

```formula
=IFERROR(1 / COUNTBLANK(Products[SKU]), "No blank SKUs")
→ No blank SKUs
```

It's convenient and dangerous. `IFERROR` hides *every* error, including the ones that mean your formula is wrong: a misspelt name, a deleted column. Prefer `IFNA` when you only expect "not found" (`#N/A`) from a lookup, and let other errors show so you notice them.

## Key ideas

- `IF(test, yes, no)` chooses between two results. `IFS` handles many tests in order; `SWITCH` matches one value against a list.
- `AND`, `OR` and `NOT` combine tests.
- TRUE is 1 and FALSE is 0 in arithmetic. Multiply conditions for AND, add them for OR, then `SUM` to count.
- Use `IFNA` for expected "not found" results; use `IFERROR` sparingly.

## Exercises

Use sheet **B1** of the practice workbook.

:::: exercise
If total revenue is at least 12,000 return `Target met`, otherwise return `Short by ` followed by the shortfall with two decimal places.
::: answer
```formula
=IF(SUM(Sales[Revenue]) >= 12000, "Target met", "Short by " & TEXT(12000 - SUM(Sales[Revenue]), "#,##0.00"))
→ Short by 216.20
```
:::
::::

:::: exercise
How many orders are closed **and** in the East region? Multiply two conditions and add up the result.
::: answer
```formula
=SUM((Sales[Region] = "East") * (Sales[Status] = "Closed"))
→ 12
```
:::
::::

:::: exercise
Label the average discount: `None` if it's 0, `Light` if it's under 5%, and `Heavy` otherwise. Use `IFS`.
::: answer
```formula
=IFS(AVERAGE(Sales[Discount]) = 0, "None", AVERAGE(Sales[Discount]) < 0.05, "Light", TRUE, "Heavy")
→ Light
```
:::
::::

:::: exercise
How many orders were for 20 units or more? Count them with a comparison and `--`.
::: answer
```formula
=SUM(--(Sales[Units] >= 20))
→ 13
```
:::
::::

:::: exercise
What is the total revenue of orders that are **not** cancelled? Use `IF` on the whole column.
::: answer
```formula
=SUM(IF(Sales[Status] <> "Cancelled", Sales[Revenue], 0))
→ 10868.15
```
:::
::::
