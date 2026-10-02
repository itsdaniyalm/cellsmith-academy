---
title: Recipe book
summary: Ready-made formulas for jobs that come up again and again - running totals, ranks, top N, distinct counts, last values, missing items and more - each one explained.
fx: =TAKE(SORTBY(Products[Product], Products[Price], -1), 3)
---

This chapter is for keeping open while you work. Each recipe is a common task, a formula that does it on the practice data, and a short note on how it works and where it came from in the course. Adapt the table and column names to your own data.

## Totals and shares

### Percentage of total

What share of revenue did each region bring in?

```formula
=LET(
    regions, {"North"; "South"; "East"; "West"},
    revenue, SUMIFS(Sales[Revenue], Sales[Region], regions),
    HSTACK(regions, ROUND(revenue / SUM(revenue), 3))
)
→ {"North", 0.194; "South", 0.248; "East", 0.313; "West", 0.244}
```

Divide each part by the sum of the parts. Format the second column as a percentage. *(Chapters B2, C3, C4.)*

### Running total

```formula
=SCAN(0, TAKE(Sales[Units], 8), LAMBDA(total, u, total + u))
→ {28; 42; 48; 103; 108; 124; 135; 139}
```

`SCAN` keeps every step of a running sum (chapter C5). In a table column, the classic version is `=SUM(INDEX([Units], 1):[@Units])`: from the first row down to this one.

### Weighted average

The average price paid per unit, weighting each order's price by its units:

```formula
=SUMPRODUCT(Sales[Units], Sales[Price]) / SUM(Sales[Units])
→ 17.88746334
```

A plain `AVERAGE(Sales[Price])` would treat a 1-unit order the same as an 80-unit one. *(Chapter B2.)*

## Ranking and ordering

### Rank each value

Where does each product's price rank, 1 being the most expensive?

```formula
=HSTACK(Products[Product], RANK.EQ(Products[Price], Products[Price]))
→ {"Desk Lamp", 4; "Notebook", 7; "Gel Pen Set", 6; "Monitor Stand", 3; "Desk Chair", 1; "Cable Kit", 5; "Sticky Notes", 8; "Whiteboard", 2}
```

`RANK.EQ(value, list)` ranks descending by default; add `1` as a third argument for ascending. Ties get the same rank.

### Top N

The three most expensive products:

```formula
=TAKE(SORTBY(Products[Product], Products[Price], -1), 3)
→ {"Desk Chair"; "Whiteboard"; "Monitor Stand"}
```

Sort, then take. Use a negative number with `TAKE` for the bottom N. *(Chapters C2, C3.)*

### Nth largest value

The second-largest order by revenue:

```formula
=LARGE(Sales[Revenue], 2)
→ 567
```

`SMALL` is the opposite. To include ties properly, use `LARGE(UNIQUE(...), 2)`.

## Lists and duplicates

### Distinct count

How many different products have been ordered?

```formula
=COUNTA(UNIQUE(Sales[Product]))
→ 8
```

### Distinct list, sorted, with a condition

The reps who have sold Desk Chairs:

```formula
=SORT(UNIQUE(FILTER(Sales[Rep], Sales[Product] = "Desk Chair")))
→ {"Chloe"; "Dev"; "Elena"; "Grace"; "Hiro"}
```

*(Chapter C2.)*

### Items in one list but not another

Which products have never been ordered with a discount?

```formula
=FILTER(Products[Product], ISNA(XMATCH(Products[Product], FILTER(Sales[Product], Sales[Discount] > 0))), "None")
→ Sticky Notes
```

`XMATCH` returns `#N/A` for each product that isn't in the discounted list, and `ISNA` turns that into TRUE. This "anti-join" pattern finds missing customers, unbilled orders, unused codes. *(Chapters B5, C2.)*

### Count how often each value appears

```formula
=LET(reps, SORT(UNIQUE(Sales[Rep])), HSTACK(reps, COUNTIFS(Sales[Rep], reps)))
→ {"Ava", 5; "Ben", 8; "Chloe", 5; "Dev", 6; "Elena", 9; "Farid", 6; "Grace", 6; "Hiro", 3}
```

## Finding things

### The last value in a column

The most recent order's ID:

```formula
=TAKE(Sales[Order ID], -1)
→ SO-1048
```

For a range that might have blanks at the bottom, `XLOOKUP(TRUE, range <> "", range, , 0, -1)` finds the last non-blank cell.

### The last match

The last product Ava sold:

```formula
=XLOOKUP("Ava", Sales[Rep], Sales[Product], , 0, -1)
→ Sticky Notes
```

Search mode `-1` searches from the bottom. *(Chapter B5.)*

### Lookup on two conditions

The revenue target for the West region in Q2:

```formula
=XLOOKUP(1, (Targets[Region] = "West") * (Targets[Quarter] = "Q2"), Targets[Target])
→ 1000
```

*(Chapter B5.)*

## Text

### Split a full name

```formula
=HSTACK(TEXTBEFORE(Staff[Name], " "), TEXTAFTER(Staff[Name], " ", -1))
→ {"Ava", "Thompson"; "Benjamin", "Okafor"; "Chloe", "Martin"; "Devendra", "Rao"; "Elena", "Petrova"; "Farid", "Haddad"; "Grace", "Lee"; "Hiro", "Tanaka"}
```

First word, and everything after the last space. *(Chapter B3.)*

### Keep only the digits

```formula
=CONCAT(IFERROR(MID("(512) 555-0142", SEQUENCE(LEN("(512) 555-0142")), 1) * 1, ""))
→ 5125550142
```

`MID` with `SEQUENCE` splits the text into single characters; multiplying by 1 turns digits into numbers and everything else into errors, which `IFERROR` blanks; `CONCAT` glues the digits back together.

### Pad a number with zeros

```formula
=TEXT(42, "00000")
→ 00042
```

### Count words

```formula
=COUNTA(TEXTSPLIT(TRIM("  the quick   brown fox "), " "))
→ 4
```

`TRIM` first, so runs of spaces don't create empty "words".

## Dates

### First and last day of the month

```formula
=TEXT(EOMONTH(DATE(2026, 5, 17), -1) + 1, "yyyy-mm-dd")
→ 2026-05-01
=TEXT(EOMONTH(DATE(2026, 5, 17), 0), "yyyy-mm-dd")
→ 2026-05-31
```

*(Chapter B4.)*

### Revenue by month

```formula
=LET(
    months, SEQUENCE(6),
    HSTACK(
        TEXT(DATE(2026, months, 1), "mmm"),
        SUMIFS(Sales[Revenue], Sales[Date], ">=" & DATE(2026, months, 1), Sales[Date], "<" & DATE(2026, months + 1, 1))
    )
)
→ {"Jan", 1976.6; "Feb", 2052.8; "Mar", 2708.15; "Apr", 2120.3; "May", 1367.45; "Jun", 1558.5}
```

One row per month, built from `SEQUENCE`. Change the `6` to cover more months.

### Age or tenure in whole years

```formula
=DATEDIF(DATE(1990, 7, 15), DATE(2026, 6, 30), "y")
→ 35
```

## Checks

### Are there any duplicates?

```formula
=IF(ROWS(UNIQUE(Sales[Order ID])) = ROWS(Sales[Order ID]), "No duplicates", "Duplicates found")
→ No duplicates
```

### Do two totals agree?

Revenue by region should add up to total revenue. A check cell like this catches broken formulas early:

```formula
=IF(ROUND(SUM(SUMIFS(Sales[Revenue], Sales[Region], {"North", "South", "East", "West"})) - SUM(Sales[Revenue]), 2) = 0, "OK", "Mismatch")
→ OK
```

`ROUND` the difference before comparing it to zero, because of the floating-point effect from chapter B1.

## Exercises

Use sheet **D3** of the practice workbook.

:::: exercise
What share of total units did the West region sell? Round to 3 decimal places.
::: answer
```formula
=ROUND(SUMIFS(Sales[Units], Sales[Region], "West") / SUM(Sales[Units]), 3)
→ 0.117
```
:::
::::

:::: exercise
List the two cheapest products by price, cheapest first.
::: answer
```formula
=TAKE(SORTBY(Products[Product], Products[Price]), 2)
→ {"Sticky Notes"; "Notebook"}
```
:::
::::

:::: exercise
Which reps have never had a cancelled order? List them in Staff table order.
::: answer
```formula
=FILTER(Staff[Rep], ISNA(XMATCH(Staff[Rep], FILTER(Sales[Rep], Sales[Status] = "Cancelled"))))
→ {"Ava"; "Ben"; "Chloe"; "Dev"; "Elena"; "Hiro"}
```
:::
::::

:::: exercise
What product did Chloe sell most recently?
::: answer
```formula
=XLOOKUP("Chloe", Sales[Rep], Sales[Product], , 0, -1)
→ Whiteboard
```
:::
::::

:::: exercise
Show the order ID `SO-1001` with its number part padded to six digits: `SO-001001`.
::: answer
```formula
="SO-" & TEXT(VALUE(TEXTAFTER("SO-1001", "-")), "000000")
→ SO-001001
```
:::
::::
