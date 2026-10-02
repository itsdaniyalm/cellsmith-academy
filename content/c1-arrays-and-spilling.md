---
title: Arrays and spilling
summary: The idea behind modern Excel. One formula can return many values, which spill into the cells around it. Array constants, the # reference, SEQUENCE, and how array maths works.
fx: =SEQUENCE(5, 1, 10, 10)
---

You've already seen formulas return more than one value: `PROPER(TRIM(Leads[Name]))` filled six cells, and `SUMIFS` with a column of criteria filled eight. This chapter explains what's going on. It's the single biggest change in Excel in thirty years, and everything in part C builds on it.

## Arrays

An **array** is a list or grid of values held together as one thing. A range like `A1:A5` is an array of five cells. A table column like `Sales[Units]` is an array of 48 numbers. And a formula can produce an array of its own.

You can also type an array directly into a formula, as an **array constant**, in curly braces:

| Constant | Shape |
|---|---|
| `{1, 2, 3}` | one row, three columns (commas separate columns) |
| `{1; 2; 3}` | three rows, one column (semicolons separate rows) |
| `{1, 2; 3, 4}` | two rows, two columns |

```formula
={10, 20, 30}
→ {10, 20, 30}
={"North"; "South"; "East"; "West"}
→ {"North"; "South"; "East"; "West"}
```

Array constants can only hold plain values: numbers, text in quotes, TRUE/FALSE and errors. No references or calculations inside the braces.

## Spilling

When a formula returns an array, Excel shows it by **spilling**: the formula lives in one cell, and its results flow into the cells below and to the right. Click any of those other cells and the formula bar shows the formula greyed out, because it belongs to the first cell, called the **anchor**. Excel draws a thin blue border around the whole spill range.

```formula cell=B2
=SEQUENCE(5, 1, 10, 10)
→ {10; 20; 30; 40; 50}
```

`SEQUENCE(rows, [columns], [start], [step])` generates numbers. This one makes 5 rows by 1 column, starting at 10, going up in steps of 10.

The spill grows and shrinks with the result. If the data changes and a formula returns twelve values instead of five, it spills into twelve cells. You never fill it down.

### When a spill is blocked: #SPILL!

A spill needs empty cells to flow into. If anything is in the way, even a single space someone typed, the anchor shows `#SPILL!` instead. Click the error's warning icon and choose **Select Obstructing Cells** to find the culprit, then clear it.

::: gotcha
Spilling formulas can't spill *inside* a table. Put them in ordinary cells next to the table or on another sheet. A formula in a table column calculates one value per row instead.
:::

### Referring to a spill: the # operator

To use a whole spill range in another formula, refer to its anchor followed by `#`. If `B2` holds the `SEQUENCE` above:

```formula
=SUM(B2#)
≈ 150
```

`B2#` means "the whole spill range that starts at B2, whatever size it is right now". It's how you build formulas on top of other dynamic results without guessing how big they'll be.

## Array maths

Operators and most functions work on arrays **element by element**. Multiply two columns of the same length and you get a column of products:

```formula
=Products[Price] - Products[Cost]
→ {18.5; 4.4; 5.6; 22; 77; 8.8; 2.95; 38}
```

That's each product's profit per unit, eight answers from one formula. Wrap it in a function that takes many values and you get one answer:

```formula
=MAX(Products[Price] - Products[Cost])
→ 77
```

If one side is a single value, it's applied to every element:

```formula
=Products[Price] * 1.1
→ {37.4; 7.15; 9.9; 53.9; 207.9; 15.4; 4.675; 86.9}
```

### Row times column makes a grid

Combine a column and a row and Excel pairs every element of one with every element of the other. A vertical array of 3 times a horizontal array of 4 gives a 3 × 4 grid:

```formula
=SEQUENCE(3) * SEQUENCE(1, 4)
→ {1, 2, 3, 4; 2, 4, 6, 8; 3, 6, 9, 12}
```

That's a multiplication table from one formula. The same idea gives you a summary grid of regions down the side and statuses across the top:

```formula
=COUNTIFS(Sales[Region], {"North"; "South"; "East"; "West"}, Sales[Status], {"Closed", "Open", "Cancelled"})
→ {12, 1, 0; 11, 0, 0; 12, 1, 2; 6, 2, 1}
```

Four regions × three statuses = twelve counts, laid out as a grid.

## Functions that return arrays

Some functions exist to produce arrays. You've met `SEQUENCE`. The next two chapters cover the rest: `FILTER`, `SORT`, `UNIQUE`, `TAKE`, `HSTACK` and friends. Here's a taste:

```formula
=TEXT(DATE(2026, SEQUENCE(6), 1), "mmm")
→ {"Jan"; "Feb"; "Mar"; "Apr"; "May"; "Jun"}
```

`SEQUENCE(6)` gives the months 1 to 6, `DATE` turns them into six dates, and `TEXT` gives their names. Three functions, each working on a whole array.

## Implicit intersection and the @ sign

Before 2019, an Excel formula could only show one value. If you gave it an array where it expected one value, Excel quietly picked the value *on the same row* as the formula. That's called **implicit intersection**.

Modern Excel spills instead. But when it opens an older workbook, it marks formulas that relied on the old behaviour with `@` so they keep working the same way:

```formula
=@Sales[Units]
```

`@` means "just the value on this row". Inside a table you'll recognise it from `[@Units]`. When you see `@` elsewhere in an old formula, it's Excel preserving how the formula used to behave. Usually you can delete it to let the formula spill.

::: note
You may also see old formulas wrapped in curly braces in the formula bar: `{=SUM(A1:A5*B1:B5)}`. These are **legacy array formulas**, which had to be entered with <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>Enter</kbd>. Modern Excel handles arrays natively, so you can just press Enter.
:::

## Key ideas

- An array is a list or grid of values. Write constants with `{ }`: commas between columns, semicolons between rows.
- A formula that returns an array **spills** from its anchor cell. `#SPILL!` means something is in the way.
- `A1#` refers to the whole spill range anchored at A1.
- Operators work element by element; a column combined with a row makes a grid.

## Exercises

Use sheet **C1** of the practice workbook.

:::: exercise
Make a column of the numbers 5, 10, 15, … up to 50 with `SEQUENCE`.
::: answer
```formula
=SEQUENCE(10, 1, 5, 5)
→ {5; 10; 15; 20; 25; 30; 35; 40; 45; 50}
```
:::
::::

:::: exercise
List each product's profit margin as a fraction of its price: price minus cost, divided by price. Round to 3 decimal places.
::: answer
```formula
=ROUND((Products[Price] - Products[Cost]) / Products[Price], 3)
→ {0.544; 0.677; 0.622; 0.449; 0.407; 0.629; 0.694; 0.481}
```
:::
::::

:::: exercise
Make a grid of order counts: the four regions down the side (North, South, East, West) and the statuses Closed and Open across the top.
::: answer
```formula
=COUNTIFS(Sales[Region], {"North"; "South"; "East"; "West"}, Sales[Status], {"Closed", "Open"})
→ {12, 1; 11, 0; 12, 1; 6, 2}
```
:::
::::

:::: exercise
Spill the full names of the months January to December across one row.
::: answer
```formula
=TEXT(DATE(2026, SEQUENCE(1, 12), 1), "mmmm")
→ {"January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"}
```
`SEQUENCE(1, 12)` makes one row of 12 numbers, so the result spills sideways.
:::
::::

:::: exercise
What is the total revenue of each quarter so far (Q1 and Q2)? Use `SUMIFS` with arrays of start and end dates.
::: answer
```formula
=SUMIFS(Sales[Revenue], Sales[Date], ">=" & DATE(2026, {1; 4}, 1), Sales[Date], "<" & DATE(2026, {4; 7}, 1))
→ {6737.55; 5046.25}
```
:::
::::
