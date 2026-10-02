---
title: Errors and debugging
summary: What each error value means and how to fix it, when to catch errors and when to let them show, and the tools for taking a broken formula apart.
fx: =IFNA(XLOOKUP("BS-999", Products[SKU], Products[Price]), 0)
---

Errors aren't failures; they're messages. Each error value names a specific kind of problem, and once you can read them, most broken formulas take seconds to fix. This chapter is a field guide, followed by the tools Excel gives you for finding where a long formula goes wrong.

## The error values

### #DIV/0!

Something was divided by zero, or by an empty cell, which counts as zero.

```formula
=SUM(Sales[Revenue]) / COUNTIFS(Sales[Region], "Central")
→ #DIV/0!
```

There's no Central region, so the count is 0. Decide what the answer *should* be when there's nothing to divide by, and say so with `IF`:

```formula
=LET(n, COUNTIFS(Sales[Region], "Central"), IF(n = 0, "No orders", SUM(Sales[Revenue]) / n))
→ No orders
```

### #N/A

"Not available": a lookup didn't find what it was looking for.

```formula
=XLOOKUP("BS-999", Products[SKU], Products[Price])
→ #N/A
```

Check for typos, extra spaces (`TRIM` helps), and numbers stored as text: the lookup value `"105"` won't match the number `105`. If "not found" is a normal outcome, use `XLOOKUP`'s `if_not_found` argument or `IFNA`.

### #NAME?

Excel doesn't recognise a name in the formula. Almost always a misspelt function, or text without its quotes.

```formula
=XLOKUP("BS-105", Products[SKU], Products[Price])
→ #NAME?
```

Text without quotes is a sneakier case, because it doesn't always show up as an error:

```formula
=COUNTIFS(Sales[Region], North)
→ 0
```

`North` without quotes looks like the name of a cell or range. There's no such name, so it evaluates to `#NAME?`, and `COUNTIFS` then counts the cells that equal that error: none. You get a believable 0 instead of an error. Write `"North"`, and be suspicious of any count that comes back as zero.

### #VALUE!

A value is the wrong kind for the operation: usually text where a number was expected.

```formula
="12 units" * 2
→ #VALUE!
="12" * 2
→ 24
```

Notice that `"12"` works: Excel converts text to a number when the text *is* a number. `"12 units"` isn't. Clean the text first (chapter B3), or check whether a "number" column is really text.

### #REF!

The formula refers to a cell that no longer exists, usually because a row, column or sheet was deleted. Excel replaces the lost reference with `#REF!` *inside the formula*, so you'll see something like `=SUM(#REF!, B2)` in the formula bar. The fix is to rewrite the reference; undo (<kbd>Ctrl</kbd>+<kbd>Z</kbd>) straight after the deletion is quickest. Table references and names make this rarer, because they adjust when the table changes.

### #NUM!

A calculation produced a number that's impossible or too big:

```formula
=SQRT(-1)
→ #NUM!
```

It also appears when an iterative function like `IRR` can't find an answer.

### #SPILL!

A formula wants to spill but something is in the way (chapter C1). Clear the obstructing cells, or move the formula.

### #CALC!

A calculation with arrays produced nothing to show, most often `FILTER` finding no matches:

```formula
=FILTER(Sales[Order ID], Sales[Region] = "Central")
→ #CALC!
```

Give `FILTER` its third argument, `if_empty`.

### #NULL!

Two references don't intersect, usually because of a stray space between them where a comma or colon was meant (chapter A3).

### ######

Not an error at all. The cell is too narrow to show the number or date in it. Widen the column. (A negative date or time also shows as `#####`.)

## Catching errors: do it narrowly

`IFERROR(value, fallback)` replaces **any** error with a fallback. That's tempting, and it's often a mistake:

```formula
=IFERROR(XLOKUP("BS-105", Products[SKU], Products[Price]), 0)
→ 0
```

The function name is misspelt, so this formula can never work. But instead of `#NAME?`, which would tell you something's wrong, it calmly returns 0 forever. Errors that mean "your formula is broken" should stay visible.

Catch the specific outcome you expect, as close to its source as possible:

| Expecting | Use |
|---|---|
| a lookup that might not find anything | `XLOOKUP(..., if_not_found)` or `IFNA` |
| a filter that might be empty | `FILTER(..., if_empty)` |
| division by a count that might be zero | `IF(count = 0, ..., ...)` |

```formula
=IFNA(XLOOKUP("BS-999", Products[SKU], Products[Price]), 0)
→ 0
```

`IFNA` catches only `#N/A`, so a typo would still show as `#NAME?`.

### Ignoring errors in a total

If a column contains some errors, `SUM` returns an error too. To add up everything else, turn the errors into zeros first:

```formula
=SUM(IFERROR({10; 20; #N/A; 30}, 0))
→ 60
```

On a worksheet range, `AGGREGATE(9, 6, range)` does the same: its first argument picks the function (9 is SUM, 1 is AVERAGE, 4 is MAX) and option 6 means "ignore errors". It only accepts real cell ranges for those functions, though, so with a calculated array it returns `#VALUE!`:

```formula
=AGGREGATE(9, 6, {10; 20; #N/A; 30})
→ #VALUE!
```

Better still, fix whatever produced the errors.

## Taking a formula apart

When a long formula returns the wrong answer, don't stare at it. Take it apart.

### Evaluate one piece with F9

In the formula bar, select part of a formula (a function call, a reference, a sub-expression) and press <kbd>F9</kbd>. Excel replaces your selection with its current value, so you can check that piece on its own. Press <kbd>Esc</kbd> afterwards to put the formula back, or you'll save the value in place of the expression.

### Evaluate Formula

**Formulas → Evaluate Formula** steps through a formula one calculation at a time, underlining the next part to be worked out and showing each intermediate result. It's slow, but it shows exactly where a value turns into something you didn't expect.

### Trace precedents and dependents

**Formulas → Trace Precedents** draws arrows from the cells a formula reads; **Trace Dependents** draws arrows to the cells that read it. Use them to find where a wrong number is coming from, or what will break if you change a cell. **Remove Arrows** clears them.

### Debug with LET

In a `LET` formula, swap the final calculation for any name to see what that step produces (chapter C4). For long formulas this is the most effective debugging tool there is.

### The Watch Window

**Formulas → Watch Window** keeps chosen cells in a small floating panel, so you can see their values while you edit inputs on another sheet.

## Key ideas

- `#DIV/0!` divide by zero · `#N/A` not found · `#NAME?` unknown name · `#VALUE!` wrong kind of value · `#REF!` deleted reference · `#NUM!` impossible number · `#SPILL!` blocked spill · `#CALC!` empty array · `#NULL!` no intersection.
- Catch only the errors you expect, with `IFNA`, `if_not_found` or `if_empty`. Broad `IFERROR` hides real mistakes.
- Debug by evaluating pieces: <kbd>F9</kbd> on a selection, Evaluate Formula, or returning an intermediate `LET` name.

## Exercises

Use sheet **D1** of the practice workbook.

:::: exercise
Divide total revenue by the number of orders in the "Central" region, returning `0` instead of an error when there are none. Don't use `IFERROR`.
::: answer
```formula
=LET(n, COUNTIFS(Sales[Region], "Central"), IF(n = 0, 0, SUM(Sales[Revenue]) / n))
→ 0
```
:::
::::

:::: exercise
Look up the price of SKU `BS-110`, returning the text `Discontinued` if it isn't found. Catch only the "not found" case.
::: answer
```formula
=IFNA(XLOOKUP("BS-110", Products[SKU], Products[Price]), "Discontinued")
→ Discontinued
```
`XLOOKUP("BS-110", Products[SKU], Products[Price], "Discontinued")` is even simpler.
:::
::::

:::: exercise
List the order IDs of orders over 1,000 units, or the text `None` if there aren't any.
::: answer
```formula
=FILTER(Sales[Order ID], Sales[Units] > 1000, "None")
→ None
```
:::
::::

:::: exercise
Add up the array `{5; #DIV/0!; 15; #N/A; 20}` while ignoring its errors.
::: answer
```formula
=SUM(IFERROR({5; #DIV/0!; 15; #N/A; 20}, 0))
→ 40
```
:::
::::
