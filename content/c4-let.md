---
title: Naming things with LET
summary: Give the parts of a formula names, calculate each part once, and turn unreadable one-liners into formulas you can follow line by line.
fx: =LET(total, SUM(Sales[Revenue]), total * TaxRate)
---

Long formulas are hard to read for the same reason a paragraph with no punctuation is: there's nowhere to pause. `LET` adds the pauses. It lets you name the pieces of a calculation and then use those names, so a formula reads as a short series of steps.

## The shape of LET

```formula
=LET(name1, value1, [name2, value2], ..., calculation)
```

You give pairs of a name and a value, then a final calculation that uses the names. The names exist only inside this one formula.

```formula
=LET(total, SUM(Sales[Revenue]), total * TaxRate)
→ 972.1635
```

Read it as: *let total be the sum of revenue; the answer is total times the tax rate.*

Each name can use the names defined before it, so the steps build on each other:

```formula
=LET(
    gross, SUM(Sales[Units] * Sales[Price]),
    net, SUM(Sales[Revenue]),
    discount, gross - net,
    ROUND(discount / gross, 4)
)
→ 0.0341
```

That's the share of gross value given away in discounts. Four lines, each saying one thing.

::: tip
Names can be any word that isn't a cell address or a function name. `total`, `west`, `isOpen` and `unit_price` are all fine. `A1` or `TAX2026` aren't, because they look like cell references. Pick names that say what the value means, not how it's calculated.
:::

## Calculate once, use many times

Here's a formula you might inherit. It labels the average closed West order:

```formula
=IFERROR(IF(SUMIFS(Sales[Revenue],Sales[Region],"West",Sales[Status],"Closed")=0,"No sales",TEXT(SUMIFS(Sales[Revenue],Sales[Region],"West",Sales[Status],"Closed")/COUNTIFS(Sales[Region],"West",Sales[Status],"Closed"),"$#,##0.00")),"Check inputs")
→ $334.93
```

The same `SUMIFS` appears twice. Excel calculates it twice, and anyone changing the formula has to remember to change both copies identically. With `LET`, each piece is written and calculated once:

```formula
=LET(
    isWest, Sales[Region] = "West",
    closed, Sales[Status] = "Closed",
    west, FILTER(Sales[Revenue], isWest * closed, ""),
    IF(
        COUNT(west) = 0,
        "No sales",
        TEXT(AVERAGE(west), "$#,##0.00")
    )
)
→ $334.93
```

Same answer. Now the formula has a vocabulary: `isWest`, `closed`, `west`. To change the region you edit one place. To check the logic, you read it top to bottom.

## Laying formulas out on several lines

The formula above is spread over several lines with indentation. Excel allows this: in the formula bar, press <kbd>Alt</kbd>+<kbd>Enter</kbd> for a new line, and type spaces to indent. Drag the bottom edge of the formula bar down (or press <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>U</kbd>) to see more than one line.

A layout that works well for `LET`:

- `=LET(` on the first line.
- One `name, value,` pair per line, indented.
- The final calculation indented, on its own line or lines.
- The closing bracket on its own line, lined up with `=LET(`.

Chapter D2 has more on formatting. The [Cellsmith](https://github.com/itsdaniyalm/cellsmith) add-in can lay formulas out like this for you.

## Debugging with LET

`LET` makes a formula easy to inspect. To see what any step produces, temporarily change the final calculation to just that name:

```formula
=LET(
    isWest, Sales[Region] = "West",
    closed, Sales[Status] = "Closed",
    west, FILTER(Sales[Revenue], isWest * closed, ""),
    west
)
→ {193.8; 210; 136; 567; 567; 335.75}
```

Now the cell spills the six revenues the average is taken from. When the step looks right, put the real calculation back.

## Building reports with LET

`LET` really pays off with the dynamic array functions from the previous chapters, where the same list is often needed several times. Suppose you want revenue by product category, a column the Sales table doesn't have. The natural first try is to look up each order's category, then hand it to `SUMIFS`:

```formula
=LET(
    category, XLOOKUP(Sales[SKU], Products[SKU], Products[Category]),
    SUMIFS(Sales[Revenue], category, "Paper")
)
≈ {"#VALUE!"; "#VALUE!"; "#VALUE!"; …}
```

That spills a column of `#VALUE!` errors, one per order, and it's a useful mistake. `SUMIFS` needs **ranges** for its criteria and sum arguments: real cells on a sheet. `category` is a calculated array, so `SUMIFS` refuses it. The same goes for `COUNTIFS`, `AVERAGEIFS` and the rest of the `IFS` family.

The fix is to use array arithmetic, which works on anything. For one category, multiply a condition by revenue as in chapter B1: `SUM((category = "Paper") * Sales[Revenue])`. For every category at once, compare every category to every name in a grid and let `MMULT` add up each row:

```formula
=LET(
    category, XLOOKUP(Sales[SKU], Products[SKU], Products[Category]),
    names, UNIQUE(category),
    totals, MMULT(--(TRANSPOSE(category) = names), Sales[Revenue]),
    SORT(HSTACK(names, totals), 2, -1)
)
→ {"Furniture", 5410.05; "Lighting", 2033.2; "Accessories", 1589; "Paper", 1408.3; "Writing", 1343.25}
```

`TRANSPOSE(category) = names` builds a grid with a row per category name and a column per order, TRUE where they match. `MMULT` multiplies that grid by the revenue column, which adds up each row's matching revenue. It's dense, but `LET` keeps it legible: each line still says one thing. In current Excel, `GROUPBY(category, Sales[Revenue], SUM)` does the same in one step.

## Key ideas

- `LET(name, value, ..., calculation)` names the parts of a formula. Later names can use earlier ones.
- Each named value is calculated once, however many times it's used.
- Lay `LET` formulas out one name per line with <kbd>Alt</kbd>+<kbd>Enter</kbd>.
- To debug, return an intermediate name instead of the final calculation.
- The `IFS` family needs real ranges, not calculated arrays. Use array arithmetic or `GROUPBY` for calculated columns.

## Exercises

Use sheet **C4** of the practice workbook.

:::: exercise
Using `LET`, name the total revenue of closed orders `closedTotal`, then return it divided by the number of closed orders, rounded to 2 decimals.
::: answer
```formula
=LET(
    closedTotal, SUMIFS(Sales[Revenue], Sales[Status], "Closed"),
    closedCount, COUNTIFS(Sales[Status], "Closed"),
    ROUND(closedTotal / closedCount, 2)
)
→ 241.94
```
:::
::::

:::: exercise
What percentage of all revenue came from the East region? Use `LET` to name the East total and the overall total, and return the share rounded to 3 decimals.
::: answer
```formula
=LET(
    east, SUMIFS(Sales[Revenue], Sales[Region], "East"),
    all, SUM(Sales[Revenue]),
    ROUND(east / all, 3)
)
→ 0.313
```
:::
::::

:::: exercise
Rewrite this formula with `LET` so the margin calculation appears only once: `=IF((SUM(Products[Price])-SUM(Products[Cost]))/SUM(Products[Price])>0.4,"Healthy: "&TEXT((SUM(Products[Price])-SUM(Products[Cost]))/SUM(Products[Price]),"0%"),"Thin")`
::: answer
```formula
=LET(
    price, SUM(Products[Price]),
    cost, SUM(Products[Cost]),
    margin, (price - cost) / price,
    IF(margin > 0.4, "Healthy: " & TEXT(margin, "0%"), "Thin")
)
→ Healthy: 46%
```
:::
::::

:::: exercise
List the order IDs of open orders worth more than the average order, using `LET` to name the average.
::: answer
```formula
=LET(
    avg, AVERAGE(Sales[Revenue]),
    FILTER(Sales[Order ID], (Sales[Status] = "Open") * (Sales[Revenue] > avg), "None")
)
→ SO-1020
```
:::
::::
