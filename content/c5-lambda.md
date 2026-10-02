---
title: Your own functions with LAMBDA
summary: Write a formula once, give it a name, and use it like a built-in function. Then MAP, BYROW, REDUCE and SCAN, which run a LAMBDA over every item in an array.
fx: =LAMBDA(price, cost, (price - cost) / price)
---

`LET` names values inside one formula. `LAMBDA` goes a step further: it turns a calculation into a **function** with inputs, which you can name and reuse anywhere in the workbook. If you've ever copied the same tricky formula into twenty places and then had to fix it in all twenty, this chapter is for you.

## A function without a name

```formula
=LAMBDA(parameter1, [parameter2], ..., calculation)
```

The parameters are names for the inputs; the last argument is the calculation that uses them. Here's a function that works out a profit margin:

```formula
=LAMBDA(price, cost, (price - cost) / price)
```

Type that into a cell and you get `#CALC!`, because a function on its own has nothing to calculate. You have to **call** it, by putting its inputs in brackets straight after:

```formula
=LAMBDA(price, cost, (price - cost) / price)(189, 112)
→ 0.407407407
```

That's the margin on a Desk Chair. Calling a `LAMBDA` like this is mostly for testing. The real use is naming it.

## Naming a LAMBDA

To make your function available everywhere in the workbook:

1. Go to **Formulas → Name Manager → New**.
2. Type a **Name**, like `Margin`.
3. In **Refers to**, paste the `LAMBDA` (starting with `=`, without the test inputs).
4. Optionally add a **Comment**. Excel shows it as the function's description when you type it.
5. Click **OK**.

Now `Margin` works like any built-in function:

```formula
=Margin(189, 112)
→ 0.407407407
=ROUND(Margin(Products[Price], Products[Cost]), 3)
→ {0.544; 0.677; 0.622; 0.449; 0.407; 0.629; 0.694; 0.481}
```

The practice workbook already defines three named LAMBDAs: `Margin`, `WithTax` and `QuarterOf`. Open the Name Manager to see how they're written.

```formula
=QuarterOf(DATE(2026, 8, 14))
→ Q3
```

::: tip
Fix a named LAMBDA once, in the Name Manager, and every formula that uses it is fixed. That's the whole point: the logic lives in one place.
:::

## Optional parameters

Put a parameter in square brackets to make it optional, and test whether it was given with `ISOMITTED`. `WithTax` adds tax at the workbook's `TaxRate` unless you pass a different rate:

```formula
=LAMBDA(amount, [rate], amount * (1 + IF(ISOMITTED(rate), TaxRate, rate)))
```

```formula
=WithTax(100)
→ 108.25
=WithTax(100, 0.2)
→ 120
```

## Running a LAMBDA over an array

A handful of functions take a `LAMBDA` as an argument and apply it to every item of an array. They let you do things that ordinary array maths can't.

### MAP: one result per item

`MAP(array, lambda)` calls the LAMBDA once per item and returns an array of the results, the same shape as the input. It's useful when a calculation doesn't work element by element on its own, like the `IFS` functions, which need a single value as their criterion each time:

```formula
=MAP(Staff[Rep], LAMBDA(rep, SUMIFS(Sales[Revenue], Sales[Rep], rep, Sales[Status], "Closed")))
→ {561.15; 1495.65; 1284.35; 1641.15; 2325.85; 602; 970.8; 1038.75}
```

Each rep's closed revenue. `MAP` can take several arrays of the same size, passing one item from each:

```formula
=MAP(Products[Price], Products[Cost], LAMBDA(p, c, ROUND(Margin(p, c), 2)))
→ {0.54; 0.68; 0.62; 0.45; 0.41; 0.63; 0.69; 0.48}
```

### BYROW and BYCOL: one result per row or column

`BYROW(array, lambda)` passes each **row** of a grid to the LAMBDA as a small array, and returns one result per row. `BYCOL` does the same for columns.

```formula
=BYROW(TAKE(Sales[[Units]:[Price]], 3), LAMBDA(row, PRODUCT(row)))
→ {252; 126; 204}
```

That multiplies Units by Price on each of the first three orders. For a calculation like this, plain array maths (`Sales[Units] * Sales[Price]`) is simpler. `BYROW` earns its place when the calculation needs a whole row at once, like `MAX` or `TEXTJOIN` across several columns:

```formula
=BYROW(Products[[Product]:[Category]], LAMBDA(r, TEXTJOIN(" - ", TRUE, r)))
→ {"Desk Lamp - Lighting"; "Notebook - Paper"; "Gel Pen Set - Writing"; "Monitor Stand - Furniture"; "Desk Chair - Furniture"; "Cable Kit - Accessories"; "Sticky Notes - Paper"; "Whiteboard - Furniture"}
```

### REDUCE: boil an array down to one value

`REDUCE(initial_value, array, lambda)` walks through the array keeping a running result, called the **accumulator**. The LAMBDA gets two parameters: the accumulator so far and the current item. Whatever it returns becomes the new accumulator.

```formula
=REDUCE(0, Sales[Units], LAMBDA(total, units, total + units))
→ 682
```

That's a long way to write `SUM`, but it shows the mechanism. `REDUCE` is for the cases nothing else covers, like applying a list of replacements one after another:

```formula
=REDUCE("(512) 555-0142", {"(", ")", " ", "-"}, LAMBDA(text, ch, SUBSTITUTE(text, ch, "")))
→ 5125550142
```

Each pass strips one character. That's chapter B3's nested `SUBSTITUTE`, without the nesting.

### SCAN: every step of the way

`SCAN` works like `REDUCE` but returns every intermediate accumulator, not just the last. That makes running totals easy:

```formula
=SCAN(0, TAKE(Sales[Revenue], 5), LAMBDA(total, r, total + r))
→ {214.2; 333.9; 527.7; 885.2; 1130.2}
```

The running total of the first five orders' revenue.

### MAKEARRAY: build a grid from its coordinates

`MAKEARRAY(rows, cols, lambda)` calls the LAMBDA with each cell's row and column number:

```formula
=MAKEARRAY(3, 4, LAMBDA(r, c, r * 10 + c))
→ {11, 12, 13, 14; 21, 22, 23, 24; 31, 32, 33, 34}
```

## Shortcuts for common LAMBDAs

When the LAMBDA would just call one function, you can pass the function's name instead, without brackets. `BYROW(data, SUM)` is short for `BYROW(data, LAMBDA(r, SUM(r)))`:

```formula
=BYCOL(Sales[[Units]:[Price]], MAX)
→ {55, 189}
```

## Key ideas

- `LAMBDA(params, calculation)` defines a function. Call it with inputs in brackets, or give it a name in the Name Manager.
- A named LAMBDA puts logic in one place: fix it once, and every use is fixed.
- `[optional]` parameters and `ISOMITTED` give defaults.
- `MAP` applies a LAMBDA to each item, `BYROW` and `BYCOL` to each row or column, `REDUCE` folds to one value, `SCAN` keeps every step, and `MAKEARRAY` builds a grid.

## Exercises

Use sheet **C5** of the practice workbook, which has `Margin`, `WithTax` and `QuarterOf` defined.

:::: exercise
Call an unnamed LAMBDA that doubles a number, giving it the input 21.
::: answer
```formula
=LAMBDA(x, x * 2)(21)
→ 42
```
:::
::::

:::: exercise
What is the price of a Desk Lamp including tax at the default rate? Look up the price, then use `WithTax`.
::: answer
```formula
=WithTax(XLOOKUP("Desk Lamp", Products[Product], Products[Price]))
→ 36.805
```
:::
::::

:::: exercise
Use `MAP` to count each rep's orders of any status, in the order of the Staff table.
::: answer
```formula
=MAP(Staff[Rep], LAMBDA(rep, COUNTIFS(Sales[Rep], rep)))
→ {5; 8; 5; 6; 9; 6; 6; 3}
```
:::
::::

:::: exercise
Use `QuarterOf` on the whole Date column and count how many orders fall in Q2.
::: answer
```formula
=SUM(--(QuarterOf(Sales[Date]) = "Q2"))
→ 23
```
`QuarterOf` works on a whole column because everything inside it (`MONTH`, `ROUNDUP`, `&`) works element by element.
:::
::::

:::: exercise
Use `SCAN` to make a running total of units across all orders, and return just the last five values with `TAKE`.
::: answer
```formula
=TAKE(SCAN(0, Sales[Units], LAMBDA(total, u, total + u)), -5)
→ {667; 672; 674; 675; 682}
```
The last value should equal the total units, 682.
:::
::::
