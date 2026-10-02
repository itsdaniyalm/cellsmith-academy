---
title: Formulas other people can read
summary: Habits that make workbooks easy to understand, check and change - for colleagues, and for you six months from now.
fx: =LET(net, SUM(Sales[Revenue]), net * (1 + TaxRate))
---

A formula is read many more times than it's written. Someone checks it, someone else inherits the workbook, you come back to it next quarter and can't remember why it's there. Every habit in this chapter makes that reading faster. None of them make a formula calculate differently; they make it *obvious* what it calculates.

## Point at meaning, not addresses

Compare:

```formula
=SUMIFS(Sales!J2:J49, Sales!C2:C49, "West") * (1 + Settings!$B$2)
```

```formula
=SUMIFS(Sales[Revenue], Sales[Region], "West") * (1 + TaxRate)
→ 3111.375625
```

They calculate the same thing. Only one tells you what. Use **tables** for data, so columns have names (chapter A2), and **named cells** for settings like rates and thresholds. A bonus: table references grow with the data, so you'll never have a total that silently stops at row 49.

## No numbers buried in formulas

A formula like `=B2 * 1.0825` contains a decision (the tax rate) that nobody can see without opening the formula. When the rate changes, someone has to find and edit every formula that contains it, and they'll miss one.

Put inputs in cells, label them, give them names, and refer to them. Everything that might change belongs where people can see it and change it once. The exceptions are numbers that are part of the logic and will never change: the 12 in "months per year", the 100 in "as a percentage".

## Name the steps with LET

Chapter C4 showed how `LET` turns a nested formula into a sequence of named steps. The same logic, before and after:

```formula
=IF(SUMIFS(Sales[Revenue],Sales[Rep],"Elena",Sales[Status],"Closed")>=XLOOKUP("East",Targets[Region],Targets[Target]),"Hit: "&TEXT(SUMIFS(Sales[Revenue],Sales[Rep],"Elena",Sales[Status],"Closed"),"#,##0"),"Miss")
→ Hit: 2,326
```

```formula
=LET(
    sold, SUMIFS(Sales[Revenue], Sales[Rep], "Elena", Sales[Status], "Closed"),
    target, XLOOKUP("East", Targets[Region], Targets[Target]),
    IF(sold >= target, "Hit: " & TEXT(sold, "#,##0"), "Miss")
)
→ Hit: 2,326
```

The second version says what it compares. It also calculates `sold` once instead of twice.

## Lay long formulas out like code

Excel ignores line breaks and spaces between the parts of a formula, so you can format it for people. In the formula bar, <kbd>Alt</kbd>+<kbd>Enter</kbd> starts a new line. Some layout rules that make formulas scannable:

- **One argument per line** when a function call doesn't fit on one line comfortably.
- **Indent** each level of nesting by the same amount (four spaces is common).
- **Closing bracket** on its own line, lined up with the line that opened it.
- **`LET`**: one `name, value` pair per line, then the result.

```formula
=IFS(
    sold >= target * 1.1, "Ahead",
    sold >= target, "On target",
    sold >= target * 0.8, "Close",
    TRUE, "Behind"
)
```

The [Cellsmith](https://github.com/itsdaniyalm/cellsmith) add-in formats formulas like this automatically, highlights them like code and flags mistakes as you type. It was built for exactly this problem.

## One formula per column

In a table, every row of a column should use the same formula. Calculated columns (chapter A2) do this for you. When one cell in a column has a different formula, a hand-typed override or a fix someone applied to a single row, it's almost invisible, and it's one of the most common causes of wrong totals. Excel flags these with a small green triangle in the cell's corner: **Inconsistent Calculated Column Formula**. Don't ignore it.

## Helper columns are fine

There's no prize for the longest formula. If a calculation is clearer in two steps, use two columns. A helper column like `Category`, looked up from Products, makes every later formula shorter, and anyone can check it by eye. With `LET` the same steps can live in one formula when you don't want the extra columns, but either way, **one step, one idea**.

## Be careful with volatile functions

`TODAY`, `NOW`, `RAND`, `RANDBETWEEN`, `OFFSET` and `INDIRECT` are **volatile**: they recalculate whenever anything in the workbook changes, and so does everything that depends on them. A few are fine. Hundreds, or one that sits under a big calculation, make a workbook sluggish.

`OFFSET` and `INDIRECT` have a second problem: they build references from text or counts, so Excel can't track what they point at. Renaming or inserting cells won't update them, and tools like Trace Precedents can't follow them. `INDEX`, `XLOOKUP`, `FILTER`, `TAKE` and `CHOOSECOLS` cover nearly everything they were used for.

## Leave notes

Excel has no comment syntax inside formulas, but there are ways to leave explanations:

- A **cell note** (right-click → **New Note**) beside a complicated formula.
- A **Comment** on a name in the Name Manager. It's shown as the description when the name is typed, which is perfect for named `LAMBDA` functions.
- In a pinch, `N("text")` returns 0, so `=SUM(Sales[Revenue]) + N("Revenue includes cancelled orders")` adds a note without changing the result. Use it sparingly; a note people can see is better.

## A checklist

Before you hand a workbook over:

- Data is in tables; formulas use structured references.
- Rates, thresholds and dates that might change are in labelled, named cells.
- Long formulas use `LET` and are laid out on several lines.
- Each table column has one consistent formula.
- Errors you expect are caught narrowly (`IFNA`, `if_empty`); others are left visible.
- No unexplained `OFFSET`, `INDIRECT` or hard-coded numbers.

## Key ideas

- Use table columns and named cells so formulas say what they mean.
- Keep inputs in cells, not inside formulas.
- Use `LET` and line breaks to show the structure of long formulas.
- One formula per column; helper columns are fine; avoid volatile functions where you can.

## Exercises

Use sheet **D2** of the practice workbook. Each exercise gives you a formula that works but is hard to read. Rewrite it so it returns the same result.

:::: exercise
Rewrite `=SUMIFS(Sales[Revenue],Sales[Region],"North")/SUM(Sales[Revenue])*100` with `LET`, naming the two totals, and round the percentage to 1 decimal place.
::: answer
```formula
=LET(
    north, SUMIFS(Sales[Revenue], Sales[Region], "North"),
    all, SUM(Sales[Revenue]),
    ROUND(north / all * 100, 1)
)
→ 19.4
```
:::
::::

:::: exercise
Rewrite `=IF(SUMIFS(Sales[Units],Sales[Rep],"Dev")>100,SUMIFS(Sales[Units],Sales[Rep],"Dev")*0.5,SUMIFS(Sales[Units],Sales[Rep],"Dev")*0.25)` so the `SUMIFS` appears only once.
::: answer
```formula
=LET(
    units, SUMIFS(Sales[Units], Sales[Rep], "Dev"),
    units * IF(units > 100, 0.5, 0.25)
)
→ 16.5
```
Moving the `IF` inside the multiplication also removes the repeated `units *`.
:::
::::

:::: exercise
This formula hides the minimum order value inside it: `=COUNTIFS(Sales[Revenue],">=100")`. Rewrite it to use the named cell `MinOrder` instead.
::: answer
```formula
=COUNTIFS(Sales[Revenue], ">=" & MinOrder)
→ 40
```
:::
::::
