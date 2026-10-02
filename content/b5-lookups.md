---
title: Looking things up
summary: XLOOKUP from top to bottom - exact and approximate matches, defaults, searching backwards, several columns at once - plus XMATCH, INDEX/MATCH and how to read old VLOOKUPs.
fx: =XLOOKUP("BS-105", Products[SKU], Products[Price])
---

A lookup answers "find this, give me that": find a product's SKU, give me its price; find a sales total, give me its commission rate. Lookups connect tables, and they're behind almost every real workbook.

Modern Excel has one lookup function to learn first, `XLOOKUP`. This chapter covers it thoroughly, then shows the older functions you'll keep meeting in other people's workbooks.

## XLOOKUP

```formula
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

The three required arguments read as a sentence: *look for this value, in this column, and return the matching item from that column.*

```formula
=XLOOKUP("BS-105", Products[SKU], Products[Price])
→ 189
=XLOOKUP("BS-105", Products[SKU], Products[Product])
→ Desk Chair
```

`XLOOKUP` finds the position of "BS-105" in the SKU column (it's 5th) and returns the 5th item of the return column. The two columns just need to be the same length. They don't have to be next to each other, and the return column can be to the *left* of the lookup column, which the old `VLOOKUP` couldn't do.

### When nothing matches

If the value isn't found, `XLOOKUP` returns `#N/A`. The fourth argument replaces that with something friendlier:

```formula
=XLOOKUP("BS-999", Products[SKU], Products[Price])
→ #N/A
=XLOOKUP("BS-999", Products[SKU], Products[Price], "Unknown SKU")
→ Unknown SKU
```

That's better than wrapping the lookup in `IFERROR`, which would also hide real mistakes (chapter B1).

### Returning several columns

Give a return range that's several columns wide and you get them all, spilled across:

```formula
=XLOOKUP("BS-104", Products[SKU], Products[[Product]:[Price]])
→ {"Monitor Stand", "Furniture", 49}
```

### Looking up a whole column at once

Give `XLOOKUP` a column of values to find and it returns a column of answers. This finds the category for every order:

```formula
=XLOOKUP(Sales[SKU], Products[SKU], Products[Category])
≈ {"Accessories"; "Paper"; "Furniture"; …}
```

Combine that with what you learned in chapter B1, and you can total revenue by category, even though the Sales table has no Category column:

```formula
=SUM((XLOOKUP(Sales[SKU], Products[SKU], Products[Category]) = "Furniture") * Sales[Revenue])
→ 5410.05
```

## Match modes

The fifth argument decides what counts as a match:

| match_mode | Finds |
|---|---|
| `0` (default) | an exact match |
| `-1` | an exact match, or else the next **smaller** value |
| `1` | an exact match, or else the next **larger** value |
| `2` | a wildcard match, with `*` and `?` |

Mode `-1` is how you look up a value in bands, like tax brackets or commission tiers. The `Tiers` table says a rep earns 2% from 0, 4% from 2,000, 6% from 5,000 and 8% from 8,000. For sales of 5,600:

```formula
=XLOOKUP(5600, Tiers[Min Sales], Tiers[Rate], , -1)
→ 0.06
```

There's no 5,600 in the table, so `XLOOKUP` takes the next smaller value, 5,000, and returns its rate. Notice the empty slot for `if_not_found`: two commas in a row skip an optional argument.

Wildcard mode finds partial matches:

```formula
=XLOOKUP("*Chair*", Products[Product], Products[SKU], , 2)
→ BS-105
```

## Search modes

By default `XLOOKUP` searches from the top and returns the **first** match. The sixth argument changes that. `-1` searches from the bottom up, which gives you the **last** match. Because Sales is in date order, the last match is the most recent:

```formula
=XLOOKUP("Hiro", Sales[Rep], Sales[Date], , 0, -1)
→ 46142
```

The answer comes back as a plain number, because a lookup returns the date's value but not its formatting. Format the cell as a date and it shows 2026-04-30. Or wrap the lookup in `TEXT`:

```formula
=TEXT(XLOOKUP("Hiro", Sales[Rep], Sales[Date], , 0, -1), "d mmm yyyy")
→ 30 Apr 2026
```

## Lookups with two conditions

`XLOOKUP` looks for one value. To match on two columns, use the trick from chapter B1: multiply the conditions, which gives 1 only on rows where both are true, and look up the 1:

```formula
=XLOOKUP(1, (Sales[Region] = "East") * (Sales[Status] = "Open"), Sales[Order ID], "None")
→ SO-1005
```

That finds the first open order in the East region.

## XMATCH

`XMATCH(lookup_value, lookup_array)` returns the **position** of a match instead of a value from another column. It takes the same match and search modes as `XLOOKUP`.

```formula
=XMATCH("Desk Chair", Products[Product])
→ 5
```

On its own, a position isn't very useful. Paired with `INDEX`, it's the classic flexible lookup.

## INDEX

`INDEX(array, row_num, [column_num])` returns the item at a given position:

```formula
=INDEX(Products[Product], 3)
→ Gel Pen Set
=INDEX(Products, 3, 4)
→ 9
```

The second formula takes row 3, column 4 of the Products table: the Gel Pen Set's price.

## INDEX and MATCH

Before `XLOOKUP` arrived, the standard flexible lookup was `INDEX` around `MATCH`. `MATCH` is the older version of `XMATCH`; its third argument must be `0` for an exact match:

```formula
=INDEX(Products[Price], MATCH("BS-105", Products[SKU], 0))
→ 189
```

Read it inside out: find the position of BS-105 in the SKU column, then return that position from the Price column. It's exactly what `XLOOKUP` does in one step. You'll see this pattern in countless workbooks, so it's worth recognising.

`INDEX` with two `XMATCH`es does a two-way lookup, finding a row and a column by name:

```formula
=INDEX(Products, XMATCH("BS-108", Products[SKU]), XMATCH("Cost", Products[#Headers]))
→ 41
```

## VLOOKUP and HLOOKUP

`VLOOKUP` was the lookup function for decades:

```formula
=VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])
```

```formula
=VLOOKUP("BS-105", Products, 4, FALSE)
→ 189
```

It searches the **first column** of the table and returns the item from column number 4. Its weaknesses are why `XLOOKUP` exists:

- It can only look in the first column and return from columns to its right.
- The column is a number. Insert a column into the table and every `VLOOKUP` that counted past it now returns the wrong column, silently.
- Forget the `FALSE` at the end and it does an approximate match by default, which returns wrong answers on unsorted data without any error.

`HLOOKUP` is the same thing sideways, searching the first row. When you inherit these, they work; when you write new formulas, use `XLOOKUP`.

## Key ideas

- `XLOOKUP(find, in, return)` is the lookup to use. Add `if_not_found` for a friendly default.
- Match mode `-1` finds bands (the next smaller value); `2` allows wildcards. Search mode `-1` finds the last match.
- To match on two columns, look up `1` in a product of conditions.
- `INDEX` returns the item at a position; `XMATCH` finds a position. `INDEX`/`MATCH` and `VLOOKUP` are the older ways to look things up.

## Exercises

Use sheet **B5** of the practice workbook.

:::: exercise
What is the cost of product BS-108?
::: answer
```formula
=XLOOKUP("BS-108", Products[SKU], Products[Cost])
→ 41
```
:::
::::

:::: exercise
Look up product BS-200. It doesn't exist, so return the text `Not found` instead of an error.
::: answer
```formula
=XLOOKUP("BS-200", Products[SKU], Products[Product], "Not found")
→ Not found
```
:::
::::

:::: exercise
What commission rate applies to sales of 8,450? Use the Tiers table.
::: answer
```formula
=XLOOKUP(8450, Tiers[Min Sales], Tiers[Rate], , -1)
→ 0.08
```
:::
::::

:::: exercise
On what date did Grace place her most recent order? Show it as text like `1 Mar 2026`.
::: answer
```formula
=TEXT(XLOOKUP("Grace", Sales[Rep], Sales[Date], , 0, -1), "d mmm yyyy")
→ 13 Jun 2026
```
:::
::::

:::: exercise
What is the order ID of the first Desk Chair order in the North region?
::: answer
```formula
=XLOOKUP(1, (Sales[Product] = "Desk Chair") * (Sales[Region] = "North"), Sales[Order ID], "None")
→ None
```
:::
::::

:::: exercise
What is the total revenue from Paper products? Look up each order's category, then sum the matching revenue.
::: answer
```formula
=SUM((XLOOKUP(Sales[SKU], Products[SKU], Products[Category]) = "Paper") * Sales[Revenue])
→ 1408.3
```
:::
::::
