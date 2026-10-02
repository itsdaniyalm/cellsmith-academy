---
title: Reshaping arrays
summary: Pick columns and rows, take the top or drop the header, stack tables side by side or on top of each other, and fold lists into grids. Plus GROUPBY, a pivot table in a formula.
fx: =HSTACK(Products[Product], Products[Price] - Products[Cost])
---

`FILTER`, `SORT` and `UNIQUE` decide *which* values you get. The functions in this chapter decide what *shape* they come in: which columns, how many rows, stacked how. They're small and simple, and they're the glue that turns a list of results into a finished report.

## Choosing columns and rows

`CHOOSECOLS(array, col_num1, ...)` returns the columns you name, by position, in the order you give them. `CHOOSEROWS` does the same for rows.

```formula
=CHOOSECOLS(Products, 2, 4)
→ {"Desk Lamp", 34; "Notebook", 6.5; "Gel Pen Set", 9; "Monitor Stand", 49; "Desk Chair", 189; "Cable Kit", 14; "Sticky Notes", 4.25; "Whiteboard", 79}
```

That's the second and fourth columns of Products: names and prices. Negative numbers count from the end, so `CHOOSECOLS(Products, -1)` is the last column.

It's most useful straight after a `FILTER` that returned whole rows, when you only want a few columns:

```formula
=CHOOSECOLS(FILTER(Sales, Sales[Status] = "Cancelled"), 1, 4, 10)
→ {"SO-1029", "Farid", 324; "SO-1042", "Farid", 196.65; "SO-1045", "Grace", 395}
```

Order ID, Rep and Revenue of the cancelled orders.

## Taking and dropping

| Function | Returns |
|---|---|
| `TAKE(array, rows, [cols])` | the first rows (negative: the last rows) |
| `DROP(array, rows, [cols])` | everything except the first rows (negative: except the last) |

```formula
=TAKE(Sales[Order ID], 3)
→ {"SO-1001"; "SO-1002"; "SO-1003"}
=TAKE(Sales[Order ID], -2)
→ {"SO-1047"; "SO-1048"}
```

`TAKE` after a sort is the standard top-N pattern you saw in chapter C2. `DROP` is useful for removing a header row from data you've stacked yourself.

## Stacking arrays

`HSTACK(array1, array2, ...)` puts arrays **side by side**. `VSTACK` puts them **on top of each other**. This is how you assemble a table from separately calculated columns:

```formula
=HSTACK(Products[Product], Products[Price] - Products[Cost])
→ {"Desk Lamp", 18.5; "Notebook", 4.4; "Gel Pen Set", 5.6; "Monitor Stand", 22; "Desk Chair", 77; "Cable Kit", 8.8; "Sticky Notes", 2.95; "Whiteboard", 38}
```

Add a header row on top with `VSTACK`, using an array constant for the labels:

```formula
=VSTACK({"Product", "Profit"}, HSTACK(Products[Product], Products[Price] - Products[Cost]))
→ {"Product", "Profit"; "Desk Lamp", 18.5; "Notebook", 4.4; "Gel Pen Set", 5.6; "Monitor Stand", 22; "Desk Chair", 77; "Cable Kit", 8.8; "Sticky Notes", 2.95; "Whiteboard", 38}
```

`VSTACK` is also how you combine lists that live in different places, like two months of data on two sheets: `VSTACK(January!A2:D50, February!A2:D50)`.

::: note
If the arrays you stack are different sizes, the gaps are filled with `#N/A`. Wrap the stack in `IFNA(..., "")` to blank them.
:::

## Changing the shape

| Function | Does |
|---|---|
| `TOCOL(array)` | flattens a grid into one column |
| `TOROW(array)` | flattens a grid into one row |
| `WRAPROWS(vector, n)` | folds a list into rows of n items |
| `WRAPCOLS(vector, n)` | folds a list into columns of n items |
| `EXPAND(array, rows, [cols], [pad_with])` | pads an array out to a bigger size |
| `TRANSPOSE(array)` | swaps rows and columns |

```formula
=WRAPROWS(Staff[Rep], 2)
→ {"Ava", "Ben"; "Chloe", "Dev"; "Elena", "Farid"; "Grace", "Hiro"}
=TRANSPOSE(Products[Product])
→ {"Desk Lamp", "Notebook", "Gel Pen Set", "Monitor Stand", "Desk Chair", "Cable Kit", "Sticky Notes", "Whiteboard"}
```

The reps happen to be listed two per region, so `WRAPROWS` with 2 lays them out one region per row.

`TOCOL` has an optional second argument that skips blanks and errors, which makes it handy for cleaning up: `TOCOL(range, 1)` returns every non-blank cell of a range as one tidy column.

## A summary report from one formula

Here's everything so far working together: one row per rep, with their region, number of closed orders and closed revenue, biggest earner first.

```formula
=SORT(
    HSTACK(
        Staff[Rep],
        Staff[Region],
        COUNTIFS(Sales[Rep], Staff[Rep], Sales[Status], "Closed"),
        SUMIFS(Sales[Revenue], Sales[Rep], Staff[Rep], Sales[Status], "Closed")
    ),
    4,
    -1
)
→ {"Elena", "East", 9, 2325.85; "Dev", "South", 6, 1641.15; "Ben", "North", 8, 1495.65; "Chloe", "South", 5, 1284.35; "Hiro", "West", 3, 1038.75; "Grace", "West", 3, 970.8; "Farid", "East", 3, 602; "Ava", "North", 4, 561.15}
```

Notice how the formula is laid out over several lines with indentation. Excel accepts line breaks inside a formula: press <kbd>Alt</kbd>+<kbd>Enter</kbd> in the formula bar. Chapter D2 is about writing formulas this way.

## GROUPBY and PIVOTBY

The newest Microsoft 365 versions of Excel add two functions that build summaries directly, like a pivot table that's a formula:

```formula
=GROUPBY(row_fields, values, function)
```

```formula
=GROUPBY(Sales[Region], Sales[Revenue], SUM)
→ {"East", 3693.5; "North", 2290.55; "South", 2925.5; "West", 2874.25; "Total", 11783.8}
```

`GROUPBY` finds each distinct region, adds up its revenue, sorts the regions and adds a total row. The third argument is the function to apply: `SUM`, `AVERAGE`, `COUNT`, `MAX` and so on, written without brackets.

`PIVOTBY` adds a column field, giving a cross-tab:

```formula
=PIVOTBY(Sales[Region], Sales[Status], Sales[Revenue], SUM)
→ {"", "Cancelled", "Closed", "Open", "Total"; "East", 520.65, 2927.85, 245, 3693.5; "North", "", 2056.8, 233.75, 2290.55; "South", "", 2925.5, "", 2925.5; "West", 395, 2009.55, 469.7, 2874.25; "Total", 915.65, 9919.7, 948.45, 11783.8}
```

Both take optional arguments for headers, totals, sorting and filtering. If your Excel doesn't have them yet, the `HSTACK` and `SUMIFS` pattern above does the same job.

## Key ideas

- `CHOOSECOLS` and `CHOOSEROWS` pick by position; `TAKE` and `DROP` keep or remove rows from either end.
- `HSTACK` joins arrays side by side, `VSTACK` stacks them vertically. Together they build tables, headers included.
- `TOCOL`, `TOROW`, `WRAPROWS`, `WRAPCOLS` and `TRANSPOSE` change an array's shape.
- `GROUPBY` and `PIVOTBY` build summaries in one step, in current versions of Excel.

## Exercises

Use sheet **C3** of the practice workbook.

:::: exercise
Show the product name and category of every product, as a two-column table.
::: answer
```formula
=CHOOSECOLS(Products, 2, 3)
→ {"Desk Lamp", "Lighting"; "Notebook", "Paper"; "Gel Pen Set", "Writing"; "Monitor Stand", "Furniture"; "Desk Chair", "Furniture"; "Cable Kit", "Accessories"; "Sticky Notes", "Paper"; "Whiteboard", "Furniture"}
```
`Products[[Product]:[Category]]` gives the same result, because the two columns are next to each other.
:::
::::

:::: exercise
List the order IDs of the last five orders in the Sales table.
::: answer
```formula
=TAKE(Sales[Order ID], -5)
→ {"SO-1044"; "SO-1045"; "SO-1046"; "SO-1047"; "SO-1048"}
```
:::
::::

:::: exercise
Build a two-column table of each rep and their hire year, with the header row `Rep` and `Hired`.
::: answer
```formula
=VSTACK({"Rep", "Hired"}, HSTACK(Staff[Rep], YEAR(Staff[Hire Date])))
→ {"Rep", "Hired"; "Ava", 2019; "Ben", 2021; "Chloe", 2018; "Dev", 2023; "Elena", 2020; "Farid", 2022; "Grace", 2017; "Hiro", 2024}
```
:::
::::

:::: exercise
Lay out the eight product SKUs as a grid of two rows of four.
::: answer
```formula
=WRAPROWS(Products[SKU], 4)
→ {"BS-101", "BS-102", "BS-103", "BS-104"; "BS-105", "BS-106", "BS-107", "BS-108"}
```
:::
::::

:::: exercise
For each region (in the order North, South, East, West), show the region and its number of orders, as a two-column table.
::: answer
```formula
=HSTACK({"North"; "South"; "East"; "West"}, COUNTIFS(Sales[Region], {"North"; "South"; "East"; "West"}))
→ {"North", 13; "South", 11; "East", 15; "West", 9}
```
:::
::::
