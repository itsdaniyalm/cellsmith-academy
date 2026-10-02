---
title: References
summary: Pointing at cells, ranges, other sheets, names and tables. Why some references move when you copy a formula and others don't.
fx: =SUM(Sales[Revenue]) * TaxRate
---

A formula is only as useful as what it points at. A **reference** is the address of the cells a formula reads. Excel has several ways to write one, and choosing the right kind is the difference between a formula you write once and copy everywhere, and one you have to fix by hand in every cell.

## Cell addresses

Columns are lettered (A, B, … Z, AA, AB, … up to XFD) and rows are numbered (1 to 1,048,576). A cell's address is its column letter followed by its row number, so `C7` is the third column, seventh row. This is called **A1 notation**, and it's what everyone uses.

A **range** is a rectangle of cells, written as its top-left and bottom-right corners joined by a colon:

| Reference | Means |
|---|---|
| `B2` | One cell |
| `B2:B10` | Nine cells in column B |
| `B2:D5` | A block, 3 columns by 4 rows |
| `B:B` | All of column B |
| `3:3` | All of row 3 |

Functions that take a range work on all of its cells at once:

```formula
=SUM(B2:B10)
≈ 1,250
```

::: gotcha
Whole-column references like `B:B` are convenient, but they include the header and every empty row below your data. Most functions cope, but they make some formulas slow and they're easy to misread. Tables, later in this chapter, are a better answer.
:::

## Relative references move

Here's the behaviour that confuses everyone at first. Suppose `D2` contains:

```formula cell=D2
=B2 * C2
≈ 54
```

Copy D2 and paste it into D3. The new formula is not `=B2 * C2`. It's:

```formula cell=D3
=B3 * C3
≈ 99.95
```

Excel doesn't store "B2". It stores "the cell two columns to my left, on my row". When you copy the formula down one row, the reference moves down one row with it. This is called a **relative reference**, and it's exactly what you want most of the time: write the formula once, then fill it down a whole column.

## Absolute references stay put

Sometimes a formula should always point at the same cell, like a tax rate in `G1`. Fill this down and it breaks:

```formula cell=E2
=D2 * G1
```

In E3 it becomes `=D3 * G2`, and G2 is empty. Put a dollar sign in front of the part that must not move:

```formula cell=E2
=D2 * $G$1
```

`$G$1` is an **absolute reference**. Copy it anywhere and it still points at G1. The dollar signs don't mean money; they mean "fixed".

You can fix just one half of an address. That's a **mixed reference**:

| You write | Column when copied | Row when copied |
|---|---|---|
| `G1` | moves | moves |
| `$G$1` | fixed | fixed |
| `$G1` | fixed | moves |
| `G$1` | moves | fixed |

Mixed references are for formulas you copy both across and down, like a grid of price × quantity. Prices run down column A and quantities across row 1:

```formula cell=B2
=$A2 * B$1
```

Fill it across and down, and every cell multiplies its own row's price by its own column's quantity.

::: tip
Click inside a reference in the formula bar and press <kbd>F4</kbd>. Each press cycles it through `G1`, `$G$1`, `G$1` and `$G1`.
:::

## Other sheets

To point at a cell on another sheet, put the sheet name and an exclamation mark in front:

```formula
=Settings!B5
→ 100
```

If the sheet name has spaces or punctuation, wrap it in single quotes: `='Q1 Sales'!B2`. You rarely type these by hand: start a formula, click the other sheet's tab, click the cell, and Excel writes the reference for you.

## Names

You can give a cell or range a **name** and use the name instead of the address. The practice workbook names the tax rate cell on the Settings sheet `TaxRate`, so you can write:

```formula
=TaxRate
→ 0.0825
=SUM(Sales[Revenue]) * TaxRate
→ 972.1635
```

The Settings sheet formats that cell as a percentage, so it shows 8.25% there. A new cell that refers to it shows the plain number, 0.0825, until you format it too.

That second formula reads like a sentence. Compare it with `=SUM(Sales!J2:J49)*Settings!$B$2`, which does the same thing and tells you nothing.

Names are always absolute: `TaxRate` points at the same cell wherever you copy the formula. To create one, select the cell, click the **Name Box** (the box left of the formula bar showing the cell's address), type a name and press <kbd>Enter</kbd>. **Formulas → Name Manager** lists, edits and deletes them.

## Tables and structured references

Format a block of data as a **table** (select it and press <kbd>Ctrl</kbd>+<kbd>T</kbd>) and Excel gives the table and each of its columns a name. Instead of `J2:J49` you write `Sales[Revenue]`: the Revenue column of the Sales table. This is a **structured reference**.

| Reference | Means |
|---|---|
| `Sales[Revenue]` | All the data in the Revenue column |
| `Sales` | All the data in the table (no headers) |
| `Sales[[Units]:[Price]]` | The columns from Units to Price |
| `Sales[#Headers]` | The header row |
| `Sales[@Units]` | Units on the *same row* as the formula |
| `[@Units]` | The same, inside the table itself |

Tables fix the weaknesses of ordinary ranges:

- **They grow.** Add a row to the bottom of the table and every formula that uses `Sales[Revenue]` includes it automatically. A formula using `J2:J49` wouldn't.
- **They read well.** `Sales[Revenue]` says what it is; `J2:J49` doesn't.
- **They fill themselves.** Type a formula in one cell of a new table column and Excel copies it down the whole column. The Revenue column in the practice workbook is one of these **calculated columns**:

```formula
=ROUND([@Units] * [@Price] * (1 - [@Discount]), 2)
```

`[@Units]` means "Units on this row", so each row multiplies its own values.

```formula
=ROWS(Sales)
→ 48
=AVERAGE(Products[Price])
→ 48.09375
```

From here on, the course uses table references almost everywhere. They're the habit that makes the biggest difference to how readable your workbooks are.

## Key ideas

- A range is two corners joined by a colon: `B2:D5`.
- Relative references (`B2`) move when you copy a formula; absolute ones (`$B$2`) don't; mixed ones (`$B2`, `B$2`) fix one half. <kbd>F4</kbd> cycles between them.
- `Sheet!A1` points at another sheet. Names like `TaxRate` replace addresses with words.
- Tables give you structured references like `Sales[Revenue]`, which grow with your data and say what they mean.

## Exercises

Use sheet **A2** of the practice workbook.

:::: exercise
What is the minimum order value stored in cell B5 of the Settings sheet?
::: answer
```formula
=Settings!B5
→ 100
```
The workbook also names that cell `MinOrder`, so `=MinOrder` gives the same answer.
:::
::::

:::: exercise
How much tax would be due on all revenue? Multiply the total of the Revenue column by the named cell `TaxRate`.
::: answer
```formula
=SUM(Sales[Revenue]) * TaxRate
→ 972.1635
```
:::
::::

:::: exercise
How many products are in the Products table? Use a structured reference to the whole table.
::: answer
```formula
=ROWS(Products)
→ 8
```
`Products` on its own means the table's data rows, without the header.
:::
::::

:::: exercise
What is the average cost of a product?
::: answer
```formula
=AVERAGE(Products[Cost])
→ 25.9375
```
:::
::::
