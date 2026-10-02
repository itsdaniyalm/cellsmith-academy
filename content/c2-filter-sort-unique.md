---
title: FILTER, SORT and UNIQUE
summary: Pull out the rows you need, put them in order and remove duplicates. Several conditions with AND and OR, sorting by several columns, top-N lists and distinct counts.
fx: =SORT(FILTER(Sales[[Order ID]:[Revenue]], Sales[Region] = "West"), 10, -1)
---

These three functions do what used to take a filter, a copy, a paste and a sort, except the result updates itself when the data changes. If you already know `FILTER`, this chapter will show you how much further it goes when you combine it with the others.

## FILTER

```formula
=FILTER(array, include, [if_empty])
```

`FILTER` returns the rows of `array` where `include` is TRUE. The include argument is a condition on a column of the same length:

```formula
=FILTER(Sales[Order ID], Sales[Status] = "Open")
→ {"SO-1002"; "SO-1005"; "SO-1020"; "SO-1044"}
```

The first argument can be several columns wide, or the whole table. You get back complete rows:

```formula
=FILTER(Sales[[Order ID]:[Rep]], Sales[Status] = "Cancelled")
→ {"SO-1029", 46127, "East", "Farid"; "SO-1042", 46175, "East", "Farid"; "SO-1045", 46186, "West", "Grace"}
```

The Date column comes back as numbers like 46127. Spilled results carry values, not formatting, so select the spill range and format that column as a date. (46127 is 2026-04-15.)

### When nothing matches

If no row passes, `FILTER` returns `#CALC!`, meaning "empty result". The third argument gives a value to show instead:

```formula
=FILTER(Sales[Order ID], Sales[Units] > 500, "No orders that big")
→ No orders that big
```

### Several conditions

`FILTER` has no separate AND and OR arguments. It uses the arithmetic from chapter B1: multiply conditions for AND, add them for OR. Wrap each condition in brackets.

```formula
=FILTER(Sales[Order ID], (Sales[Region] = "West") * (Sales[Status] = "Closed"))
→ {"SO-1003"; "SO-1010"; "SO-1014"; "SO-1015"; "SO-1024"; "SO-1033"}
=FILTER(Sales[Order ID], (Sales[Discount] = 0.15) + (Sales[Units] >= 50))
→ {"SO-1001"; "SO-1004"; "SO-1021"; "SO-1025"; "SO-1027"; "SO-1032"; "SO-1033"; "SO-1044"}
```

The first keeps orders that are West **and** closed. The second keeps orders with a 15% discount **or** at least 50 units. With `+`, a row that meets both conditions adds up to 2, which `FILTER` still counts as TRUE: any non-zero number does.

::: tip
For "not", compare with `<>`, or wrap a condition in `NOT(...)`. For "matches any of these values", use `ISNUMBER(XMATCH(column, list))`: `FILTER(Sales[Order ID], ISNUMBER(XMATCH(Sales[Rep], {"Ava", "Ben"})))` keeps orders by Ava or Ben.
:::

### Partial text matches

`FILTER` has no wildcards, but `SEARCH` (chapter B3) returns a number when it finds text and an error when it doesn't. `ISNUMBER` turns that into TRUE/FALSE:

```formula
=FILTER(Products[Product], ISNUMBER(SEARCH("desk", Products[Product])))
→ {"Desk Lamp"; "Desk Chair"}
```

## SORT

```formula
=SORT(array, [sort_index], [sort_order], [by_col])
```

`sort_index` is which column of the array to sort by (1 by default), and `sort_order` is `1` for ascending or `-1` for descending.

```formula
=SORT(Products[[Product]:[Price]], 3, -1)
→ {"Desk Chair", "Furniture", 189; "Whiteboard", "Furniture", 79; "Monitor Stand", "Furniture", 49; "Desk Lamp", "Lighting", 34; "Cable Kit", "Accessories", 14; "Gel Pen Set", "Writing", 9; "Notebook", "Paper", 6.5; "Sticky Notes", "Paper", 4.25}
```

That's the products, most expensive first: sort by the third column of the array (Price), descending.

To sort by several columns, give arrays of indexes and orders: `SORT(data, {2, 3}, {1, -1})` sorts by column 2 ascending, then by column 3 descending.

## SORTBY

`SORTBY(array, by_array1, [order1], ...)` sorts by other arrays instead of column numbers. The sort key doesn't even have to be part of the result:

```formula
=SORTBY(Products[Product], Products[Price] - Products[Cost], -1)
→ {"Desk Chair"; "Whiteboard"; "Monitor Stand"; "Desk Lamp"; "Cable Kit"; "Gel Pen Set"; "Notebook"; "Sticky Notes"}
```

Products by profit per unit, highest first, though the profit itself isn't shown. Use `SORTBY` when the sort key is a calculation or a column you don't want in the output.

## UNIQUE

`UNIQUE(array, [by_col], [exactly_once])` removes duplicates:

```formula
=UNIQUE(Sales[Rep])
→ {"Elena"; "Grace"; "Ben"; "Farid"; "Ava"; "Hiro"; "Chloe"; "Dev"}
```

The list keeps the order of first appearance. Wrap it in `SORT` for an alphabetical list, which is one of the most common formulas in modern Excel:

```formula
=SORT(UNIQUE(Sales[Region]))
→ {"East"; "North"; "South"; "West"}
```

Count the distinct values with `ROWS` or `COUNTA`:

```formula
=ROWS(UNIQUE(Sales[Product]))
→ 8
```

Set the third argument to `TRUE` to get only the values that appear **exactly once**. In the Leads table, three leads came from the web, two from referrals and one from an event:

```formula
=UNIQUE(Leads[Source], , TRUE)
→ Event
```

Every rep has more than one order, so `UNIQUE(Sales[Rep], , TRUE)` finds nothing and returns `#CALC!`, the same "empty result" error as `FILTER`.

## Putting them together

The real power is in combining them. Each function takes an array and returns an array, so they nest like building blocks.

**The reps who have open orders, alphabetically:**

```formula
=SORT(UNIQUE(FILTER(Sales[Rep], Sales[Status] = "Open")))
→ {"Ava"; "Farid"; "Grace"}
```

**The three biggest orders.** Put the columns you want side by side with `HSTACK`, sort the rows by revenue with `SORTBY`, then keep the first three with `TAKE` (both are in chapter C3):

```formula
=TAKE(SORTBY(HSTACK(Sales[Order ID], Sales[Rep], Sales[Revenue]), Sales[Revenue], -1), 3)
→ {"SO-1015", "Grace", 567; "SO-1024", "Hiro", 567; "SO-1028", "Elena", 567}
```

**A summary table: each region with its revenue, biggest first.** `UNIQUE` gives the regions, `SUMIFS` with that list gives one total per region, and `HSTACK` (chapter C3) puts the two columns side by side:

```formula
=SORT(HSTACK(UNIQUE(Sales[Region]), SUMIFS(Sales[Revenue], Sales[Region], UNIQUE(Sales[Region]))), 2, -1)
→ {"East", 3693.5; "South", 2925.5; "West", 2874.25; "North", 2290.55}
```

That's a pivot-table-style summary from one formula, and it updates as soon as an order is added. Chapter C4 shows how `LET` makes formulas like this much easier to read.

## Key ideas

- `FILTER(array, condition, [if_empty])` keeps matching rows. Multiply conditions for AND, add them for OR.
- `SORT` sorts by column position; `SORTBY` sorts by any array, including a calculation.
- `UNIQUE` removes duplicates; `ROWS(UNIQUE(...))` counts distinct values.
- They nest: `SORT(UNIQUE(FILTER(...)))` is a filtered, de-duplicated, sorted list.

## Exercises

Use sheet **C2** of the practice workbook.

:::: exercise
List the order IDs of every Desk Chair order.
::: answer
```formula
=FILTER(Sales[Order ID], Sales[Product] = "Desk Chair")
→ {"SO-1015"; "SO-1024"; "SO-1025"; "SO-1028"; "SO-1034"; "SO-1039"}
```
:::
::::

:::: exercise
List the order IDs of closed orders in the North region with at least 10 units.
::: answer
```formula
=FILTER(Sales[Order ID], (Sales[Region] = "North") * (Sales[Status] = "Closed") * (Sales[Units] >= 10))
→ {"SO-1004"; "SO-1006"; "SO-1022"; "SO-1032"}
```
:::
::::

:::: exercise
Make an alphabetical list of the products that have been ordered with a discount.
::: answer
```formula
=SORT(UNIQUE(FILTER(Sales[Product], Sales[Discount] > 0)))
→ {"Cable Kit"; "Desk Chair"; "Desk Lamp"; "Gel Pen Set"; "Monitor Stand"; "Notebook"; "Whiteboard"}
```
:::
::::

:::: exercise
How many different reps have sold Notebooks?
::: answer
```formula
=ROWS(UNIQUE(FILTER(Sales[Rep], Sales[Product] = "Notebook")))
→ 2
```
:::
::::

:::: exercise
List the product names sorted by cost, cheapest first.
::: answer
```formula
=SORTBY(Products[Product], Products[Cost])
→ {"Sticky Notes"; "Notebook"; "Gel Pen Set"; "Cable Kit"; "Desk Lamp"; "Monitor Stand"; "Whiteboard"; "Desk Chair"}
```
:::
::::

:::: exercise
Show the order ID and revenue of the five largest orders, largest first.
::: answer
```formula
=TAKE(SORTBY(HSTACK(Sales[Order ID], Sales[Revenue]), Sales[Revenue], -1), 5)
→ {"SO-1015", 567; "SO-1024", 567; "SO-1028", 567; "SO-1039", 567; "SO-1045", 395}
```
:::
::::
