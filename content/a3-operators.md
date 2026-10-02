---
title: Operators and order
summary: Arithmetic, joining text, comparisons and the reference operators, plus the order Excel applies them in, including the one rule that catches everyone out.
fx: ="Total: " & SUM(Sales[Units] * Sales[Price])
---

Operators are the symbols between the values in a formula: `+`, `*`, `&`, `>` and friends. There are only a handful of them, but the order Excel applies them in decides what your formula actually calculates.

## Arithmetic

| Operator | Does | Example | Result |
|---|---|---|---|
| `+` | add | `=7 + 2` | 9 |
| `-` | subtract | `=7 - 2` | 5 |
| `*` | multiply | `=7 * 2` | 14 |
| `/` | divide | `=7 / 2` | 3.5 |
| `^` | raise to a power | `=7 ^ 2` | 49 |
| `%` | divide by 100 | `=7%` | 0.07 |
| `-` (in front) | negate | `=-7` | −7 |

Spaces around operators are optional. Excel ignores them, so write whichever is easier to read.

## Order of operations

Excel doesn't calculate left to right. Like school maths, multiplication and division happen before addition and subtraction:

```formula
=2 + 3 * 4
→ 14
=(2 + 3) * 4
→ 20
```

Here is the full order, from first to last. Operators on the same line have equal priority and are applied left to right.

| Priority | Operators | |
|---|---|---|
| 1 | `:` `,` space | reference operators (below) |
| 2 | `-` | negation, as in `-A1` |
| 3 | `%` | percent |
| 4 | `^` | power |
| 5 | `*` `/` | multiply and divide |
| 6 | `+` `-` | add and subtract |
| 7 | `&` | join text |
| 8 | `=` `<>` `<` `>` `<=` `>=` | compare |

When in doubt, add brackets. They cost nothing and they make your intent obvious to the next reader.

::: gotcha
Negation comes **before** powers. In maths, −2² is −4. In Excel:

```formula
=-2^2
→ 4
=0 - 2^2
→ -4
```

The first formula squares −2. The second subtracts 2² from zero. If you mean "the negative of a square", write `=-(2^2)`.
:::

Equal-priority operators go left to right, which includes powers. Excel reads `2^3^2` as `(2^3)^2`, not `2^(3^2)` as a mathematician would:

```formula
=2^3^2
→ 64
```

## Joining text with &

The ampersand glues values together into one piece of text:

```formula
="SO-" & 1001
→ SO-1001
="Total units: " & SUM(Sales[Units])
→ Total units: 682
```

Numbers are converted to text as they're joined, using their plain form rather than any cell formatting. To control how a number looks inside text (two decimals, a thousands separator, a date), wrap it in `TEXT`, which chapter B3 covers:

```formula
="Revenue: " & TEXT(SUM(Sales[Revenue]), "#,##0.00")
→ Revenue: 11,783.80
```

## Comparisons

Comparison operators ask a question and answer `TRUE` or `FALSE`:

```formula
=5 > 3
→ TRUE
=SUM(Sales[Units]) >= 700
→ FALSE
="North" <> "South"
→ TRUE
```

`<>` means "not equal to". Comparisons are the raw material of every decision a formula makes; chapter B1 builds on them.

Two surprises are worth knowing now. Text comparisons ignore case, and text is always treated as greater than any number:

```formula
="apple" = "APPLE"
→ TRUE
="10" > 9999
→ TRUE
```

The second one is how text-that-looks-like-a-number breaks comparisons. `"10"` is text, so it's "bigger" than every number.

## Reference operators

Three operators combine references rather than values:

| Operator | Name | Example | Means |
|---|---|---|---|
| `:` | range | `A1:C3` | everything from A1 to C3 |
| `,` | union | `SUM((A1:A3, C1:C3))` | both ranges together |
| space | intersection | `=B:B 3:3` | the cell where column B meets row 3 |

You'll use `:` constantly. The other two are rare, but the space operator can bite: a stray space between two references gives you their intersection, which is often empty and returns `#NULL!`.

## Operators work on whole columns

Here's a preview of something chapter C1 explores fully. Operators work element by element on ranges and table columns:

```formula
=SUM(Sales[Units] * Sales[Price])
→ 12199.25
```

`Sales[Units] * Sales[Price]` multiplies each order's units by its price, producing 48 results at once. `SUM` then adds them up. That's the total value of all orders before discounts, in one formula with no helper column.

## Key ideas

- Excel applies operators in a fixed order: reference operators, negation, percent, power, multiply/divide, add/subtract, join, compare. Brackets override it.
- Negation binds tighter than `^`, so `=-2^2` is 4.
- `&` joins values into text; wrap numbers in `TEXT` to control their format.
- Comparisons return `TRUE` or `FALSE`. Text comparisons ignore case, and text ranks above numbers.
- Operators work element by element on whole columns.

## Exercises

Use sheet **A3** of the practice workbook.

:::: exercise
What is the total value of all orders before discounts? Multiply the Units column by the Price column and add up the results.
::: answer
```formula
=SUM(Sales[Units] * Sales[Price])
→ 12199.25
```
:::
::::

:::: exercise
Write a formula that returns the text `Total units: 682`, calculating the number from the Sales table.
::: answer
```formula
="Total units: " & SUM(Sales[Units])
→ Total units: 682
```
:::
::::

:::: exercise
How much revenue would there be including tax? Add `TaxRate` to the total revenue (multiply by one plus the rate) and round to 2 decimal places with `ROUND(value, 2)`.
::: answer
```formula
=ROUND(SUM(Sales[Revenue]) * (1 + TaxRate), 2)
→ 12755.96
```
The brackets around `1 + TaxRate` matter. Without them, Excel would multiply first and then add the tax rate.
:::
::::

:::: exercise
How much did discounts cost the business? Subtract total revenue from the value before discounts.
::: answer
```formula
=SUM(Sales[Units] * Sales[Price]) - SUM(Sales[Revenue])
→ 415.45
```
:::
::::
