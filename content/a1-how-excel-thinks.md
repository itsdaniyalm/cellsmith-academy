---
title: How Excel thinks
summary: Cells, values and formulas. What happens when you press Enter, and the five kinds of value every formula produces.
fx: =Units * Price
---

A spreadsheet looks like a table, but it behaves more like a program. Every cell is a small box that holds one thing. Some boxes hold values you typed. Others hold instructions, called **formulas**, that tell Excel how to work a value out from other boxes. When anything changes, Excel re-runs the instructions that depend on it.

Once you see a workbook this way, most of Excel's behaviour stops being mysterious. This chapter builds that picture.

## Values and formulas

Type `42` into a cell and press <kbd>Enter</kbd>. The cell now holds the value 42. It will hold 42 until you type something else.

Now type this into another cell:

```formula
=6 * 7
→ 42
```

The cell *shows* 42 too, but it holds something different: an instruction. Click the cell and look at the formula bar above the grid. You'll see `=6*7`, not 42. The grid shows each cell's **result**; the formula bar shows what the cell **contains**.

The leading `=` is what makes it a formula. Without it, Excel treats what you type as plain text or a number. That one character is the switch between "store this" and "calculate this".

::: tip
Press <kbd>Ctrl</kbd>+<kbd>`</kbd> (the key left of <kbd>1</kbd>) to show formulas instead of results across the whole sheet. Press it again to switch back. It's the quickest way to see which cells are calculated and which are typed.
:::

## Formulas point at other cells

A formula that only uses fixed numbers is a calculator. The real power is pointing at other cells. Suppose `B2` holds a quantity and `C2` holds a price:

```formula cell=D2
=B2 * C2
≈ 156
```

`D2` now holds an instruction: *multiply whatever is in B2 by whatever is in C2*. Change B2 and D2 updates immediately. You never re-type the formula; you change the inputs.

This is called **recalculation**, and it's automatic. Excel keeps track of which cells depend on which, so when you edit one cell it recalculates everything downstream of it, in the right order. A formula can depend on a cell that depends on another cell, and so on, in long chains.

::: note
Excel won't let a chain loop back on itself. If D2 depends on E2 and E2 depends on D2, Excel warns you about a **circular reference**, because there's no order in which it could calculate them.
:::

## The five kinds of value

Every cell, and every formula result, is one of five kinds of value. Knowing which kind you have explains a lot of surprising results.

| Kind | Examples | Lines up |
|---|---|---|
| Number | `42`, `-3.5`, `8%`, dates, times | right |
| Text | `"North"`, `"SO-1001"`, `"42"` | left |
| Logical | `TRUE`, `FALSE` | centre |
| Error | `#N/A`, `#DIV/0!`, `#VALUE!` | centre |
| Blank | nothing at all | |

Inside a formula, text always goes in double quotes. Without quotes, Excel assumes a word is the name of something, like a function or a named cell:

```formula
="North"
→ North
=LEN("North")
→ 5
```

### Numbers that look like text, and text that looks like numbers

`42` and `"42"` are different values. One is a number you can add; the other is two characters of text. Data pasted in from other systems often arrives as text-that-looks-like-a-number, and it causes quiet trouble: sums skip it, lookups miss it. One clue is alignment. By default Excel lines numbers up on the right of a cell and text on the left.

```formula
=ISNUMBER(42)
→ TRUE
=ISNUMBER("42")
→ FALSE
```

### Dates are numbers

This is the single most useful fact in this chapter. Excel stores every date as a count of days since 1 January 1900. The *format* of the cell makes it look like a date, but underneath it is a whole number:

```formula
=DATE(2026, 1, 2) * 1
→ 46024
```

Multiplying by 1 strips off the date format so you can see the number. Because dates are numbers, you can subtract them to get a gap in days, or add days to them:

```formula
=DATE(2026, 3, 1) - DATE(2026, 1, 1)
→ 59
```

Times work the same way: they're fractions of a day. Noon is `0.5`, 6 a.m. is `0.25`. Chapter B4 is all about dates.

### Logical values are numbers in disguise

`TRUE` and `FALSE` are their own kind of value, but when you do arithmetic with them Excel treats `TRUE` as 1 and `FALSE` as 0:

```formula
=TRUE + TRUE
→ 2
```

That sounds like trivia. It turns out to be the key to `FILTER` with several conditions, and to a whole family of counting tricks in chapter B2.

### Errors are values too

When a formula can't produce a sensible answer, it returns an error value instead. Errors are not crashes: they're results, and they flow through any formula that uses them. If `B2` holds `#N/A`, then `=B2 * 2` is also `#N/A`.

```formula
=1 / 0
→ #DIV/0!
```

Each error tells you what went wrong. Chapter D1 covers every one, but here are the ones you'll meet first:

| Error | Means |
|---|---|
| `#DIV/0!` | Divided by zero (or by a blank cell). |
| `#N/A` | A lookup didn't find what it was looking for. |
| `#NAME?` | Excel doesn't recognise a name. Usually a misspelt function. |
| `#VALUE!` | Wrong kind of value, like adding a number to text that isn't a number. |
| `#REF!` | The formula points at a cell that has been deleted. |

## What happens when you press Enter

When you type a formula and press <kbd>Enter</kbd>, Excel:

1. **Reads** it. It breaks the text into pieces: numbers, text, cell addresses, function names, operators. If it can't make sense of the pieces (an unclosed bracket, say), it refuses the formula and shows a message instead of saving it.
2. **Evaluates** it, working from the innermost parts outwards, much like you would by hand.
3. **Shows** the result in the cell, using the cell's number format.
4. **Remembers** which cells the formula depends on, so it knows to recalculate when they change.

Step 3 matters more than it seems. The number format changes how a value *looks*, never what it *is*. A cell formatted to show `$54.00` still holds the number 54. A cell showing `8%` holds 0.08. When a total looks one cent off, it's almost always because the displayed values were rounded but the stored ones weren't.

## The practice data

From here on, nearly every example uses a small made-up company, **Brightside Supply Co.**, which sells office supplies. Its data lives in the [practice workbook](practice/cellsmith-academy-practice.xlsx) as a set of tables. The main one, `Sales`, has 48 orders:

```formula
=COUNTA(Sales[Order ID])
→ 48
=SUM(Sales[Revenue])
→ 11783.8
```

Don't worry about the `Sales[Revenue]` notation yet; it's how you point at a column of a table, and chapter A2 explains it. For now, read it as "the Revenue column of the Sales table". You can browse all of the tables on the [practice data](data.html) page.

## Key ideas

- A cell holds either a **value** you typed or a **formula** that calculates one. Formulas start with `=`.
- Formulas point at other cells, and Excel **recalculates** them automatically when those cells change.
- Every value is a number, text, logical, error or blank. Dates and times are numbers; `TRUE` and `FALSE` behave as 1 and 0 in arithmetic.
- Number formats change how a value looks, never what it is.

## Exercises

Open the practice workbook and go to sheet **A1**. Each exercise has a yellow cell for your formula and a check underneath it.

:::: exercise
How many orders are there in the Sales table? Count the Order ID column.
::: answer
```formula
=COUNTA(Sales[Order ID])
→ 48
```
`COUNTA` counts cells that aren't empty. `COUNT` would return 0 here, because order IDs are text and `COUNT` only counts numbers.
:::
::::

:::: exercise
What is the total number of units sold across all orders?
::: answer
```formula
=SUM(Sales[Units])
→ 682
```
:::
::::

:::: exercise
How many days are there from 1 January 2026 to 30 June 2026? Use `DATE` to build both dates and subtract.
::: answer
```formula
=DATE(2026, 6, 30) - DATE(2026, 1, 1)
→ 180
```
Because dates are numbers, subtracting one from the other gives the number of days between them.
:::
::::

:::: exercise
Is the text `"682"` a number? Write a formula that returns `FALSE`.
::: answer
```formula
=ISNUMBER("682")
→ FALSE
```
Anything in double quotes is text, even if it's made of digits.
:::
::::
