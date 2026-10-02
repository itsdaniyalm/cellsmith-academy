---
title: Reading and writing functions
summary: How a function call is put together, how to read any function's signature, and the everyday functions for totals, averages, counts and rounding.
fx: =ROUND(AVERAGE(Sales[Revenue]), 2)
---

Excel has more than 500 functions. Nobody memorises them all, and you don't need to. What you need is to recognise the shape every function shares, and to read a function's **signature** well enough to use one you've never seen before. This chapter teaches that, then covers the handful of functions you'll use every day.

## The shape of a function call

```formula
=ROUND(AVERAGE(Sales[Revenue]), 2)
→ 245.5
```

Every function call has the same parts:

- A **name**: `ROUND`, `AVERAGE`. Case doesn't matter; Excel upper-cases names for you.
- An opening bracket straight after the name, with no space.
- **Arguments**: the inputs, separated by commas. `ROUND` takes two here: the number to round, and how many decimal places to keep.
- A closing bracket.

An argument can be a value (`2`), a reference (`Sales[Revenue]`), an expression (`A1 * 2`), or another function call. When one function sits inside another it's **nested**. Excel works from the inside out: it calculates `AVERAGE(Sales[Revenue])` first, then hands the answer to `ROUND`.

## Reading a signature

Every function has a signature that lists its arguments. When you type a function name and its opening bracket, Excel shows the signature in a tooltip under the cell. Learn to read these and every function becomes approachable.

```formula
=ROUND(number, num_digits)
```

Two arguments, both required. The names hint at what goes there: a number, then a number of digits.

```formula
=IF(logical_test, [value_if_true], [value_if_false])
```

Arguments in **square brackets are optional**. You can stop after any required argument; Excel uses a default for the rest. The function's help page tells you what the default is.

```formula
=SUM(number1, [number2], ...)
```

The `...` means the last argument **repeats**. `SUM` accepts up to 255 arguments, each a number or a range.

```formula
=XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])
```

This is about as long as signatures get. Three required arguments, three optional ones. You'll meet it properly in chapter B5. To skip an optional argument but supply a later one, leave its slot empty between commas: `XLOOKUP(a, b, c, , 1)` skips `if_not_found` and sets `match_mode`.

::: tip
Stuck on a function? Click the **fx** button left of the formula bar (or press <kbd>Shift</kbd>+<kbd>F3</kbd>) to open the **Insert Function** dialog. It lets you search functions by what they do, and it shows a box for each argument with a description. Or use the [function finder](functions.html) on this site.
:::

## Functions with no arguments

A few functions need no input. They still need their brackets:

```formula
=PI()
→ 3.14159265358979
=TODAY()
≈ 2026-06-30
```

`TODAY` returns today's date, so its result changes every day. Functions like that are called **volatile**: Excel recalculates them every time anything in the workbook changes. A few of them are harmless; hundreds can slow a big workbook down.

## The everyday functions

These cover a surprising share of real work.

### Totals and averages

| Function | Returns |
|---|---|
| `SUM` | the total |
| `AVERAGE` | the mean |
| `MEDIAN` | the middle value |
| `MIN` / `MAX` | the smallest / largest |

```formula
=MAX(Sales[Revenue])
→ 567
=MIN(Sales[Date])
→ 46024
```

`MIN` of a date column gives the earliest date. It shows as a number here because dates are numbers (chapter A1). Format the cell as a date, or wrap the formula in `TEXT`, to see it as one:

```formula
=TEXT(MIN(Sales[Date]), "d mmm yyyy")
→ 2 Jan 2026
```

### Counting

| Function | Counts cells that |
|---|---|
| `COUNT` | contain numbers |
| `COUNTA` | aren't empty |
| `COUNTBLANK` | are empty |
| `ROWS` / `COLUMNS` | (counts the rows or columns of a range, filled or not) |

Picking the wrong one is a classic bug. `COUNT` ignores text, so counting a text column with it returns zero:

```formula
=COUNT(Sales[Order ID])
→ 0
=COUNTA(Sales[Order ID])
→ 48
=COUNTBLANK(Sales[Ship Date])
→ 7
```

The blank ship dates are orders that haven't shipped yet.

### Rounding

| Function | Does | `…(2.567, 1)` |
|---|---|---|
| `ROUND` | rounds to the nearest | 2.6 |
| `ROUNDUP` | always rounds away from zero | 2.6 |
| `ROUNDDOWN` | always rounds toward zero | 2.5 |
| `INT` | drops the decimals (rounds down) | (takes one argument) |

The second argument is the number of decimal places. It can be negative, which rounds to tens, hundreds and so on:

```formula
=ROUND(SUM(Sales[Revenue]), -2)
→ 11800
```

Rounding changes the value. Formatting a cell to show two decimals only changes how it looks. If you're adding up money that must match a printed report to the cent, round in the formula.

## Function categories

Excel groups its functions into categories. The ones this course covers, and where:

| Category | For | Chapter |
|---|---|---|
| Logical | decisions: `IF`, `AND`, `IFS`, `SWITCH` | B1 |
| Math | conditional totals: `SUMIFS`, `COUNTIFS` | B2 |
| Text | cleaning and building text | B3 |
| Date and time | dates, months, workdays | B4 |
| Lookup and reference | finding things: `XLOOKUP`, `FILTER` | B5, C2 |
| Dynamic arrays | reshaping lists and tables | C1–C3 |

## Key ideas

- A function call is a name, then arguments in brackets, separated by commas. Nested calls are worked out from the inside out.
- In a signature, `[square brackets]` mark optional arguments and `...` marks a repeating one.
- `COUNT` counts numbers, `COUNTA` counts anything that isn't empty.
- `ROUND` changes the value; number formatting only changes how it looks.

## Exercises

Use sheet **A4** of the practice workbook.

:::: exercise
What is the revenue of the largest single order?
::: answer
```formula
=MAX(Sales[Revenue])
→ 567
```
:::
::::

:::: exercise
What is the average number of units per order, rounded to one decimal place?
::: answer
```formula
=ROUND(AVERAGE(Sales[Units]), 1)
→ 14.2
```
:::
::::

:::: exercise
How many orders haven't shipped yet? Count the empty cells in the Ship Date column.
::: answer
```formula
=COUNTBLANK(Sales[Ship Date])
→ 7
```
:::
::::

:::: exercise
Show the date of the most recent order as text like `1 Mar 2026`.
::: answer
```formula
=TEXT(MAX(Sales[Date]), "d mmm yyyy")
→ 24 Jun 2026
```
:::
::::

:::: exercise
What is the median order revenue? It's less affected by one very large order than the average is.
::: answer
```formula
=MEDIAN(Sales[Revenue])
→ 228.875
```
:::
::::
