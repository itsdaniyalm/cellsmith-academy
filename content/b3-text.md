---
title: Working with text
summary: Cleaning, splitting, searching and building text. TRIM and PROPER for messy data, TEXTBEFORE and TEXTSPLIT for pulling it apart, TEXT and TEXTJOIN for putting it back together.
fx: =PROPER(TRIM(Leads[Name]))
---

Data rarely arrives clean. Names come in with stray spaces and random capitals, phone numbers in five formats, city and state jammed into one cell. Excel's text functions are how you fix that without retyping anything. The practice workbook's `Leads` table is deliberately messy so you can try them on real problems.

## Measuring and slicing

Text is a sequence of characters, numbered from 1. These functions pull pieces out by position:

| Function | Returns |
|---|---|
| `LEN(text)` | the number of characters |
| `LEFT(text, n)` | the first n characters |
| `RIGHT(text, n)` | the last n characters |
| `MID(text, start, n)` | n characters from position start |

```formula
=LEN("SO-1001")
→ 7
=LEFT("SO-1001", 2)
→ SO
=RIGHT("SO-1001", 4)
→ 1001
=MID("BS-104", 4, 3)
→ 104
```

The results are always text, even when they look like numbers. `RIGHT("SO-1001", 4)` is the text "1001". Wrap it in `VALUE` to get a number you can calculate with:

```formula
=VALUE(RIGHT("SO-1001", 4)) + 1
→ 1002
```

## Cleaning up

| Function | Does |
|---|---|
| `TRIM` | removes spaces at both ends, and squeezes runs of spaces inside to one |
| `CLEAN` | removes non-printing characters (line breaks from other systems, for example) |
| `UPPER` / `LOWER` | changes everything to capitals / small letters |
| `PROPER` | capitalises the first letter of each word |

These combine naturally. Here's every name in the Leads table, tidied in one formula:

```formula
=PROPER(TRIM(Leads[Name]))
→ {"Maria Garcia"; "James O'Neil"; "Priya Sharma"; "Li Wei"; "Sofia Rossi"; "Noah Johnson"}
```

Because the formula is given a whole column, it returns a whole column of answers. They **spill** into the cells below. You'll learn more about spilling in chapter C1; for now, enjoy not having to fill the formula down.

::: gotcha
`PROPER` capitalises after every non-letter, which usually helps (O'Neil) and occasionally doesn't (McDonald becomes Mcdonald). Check names with unusual capitals by eye.
:::

## Finding and replacing

`FIND(find_text, within_text)` returns the position where one piece of text first appears inside another. `SEARCH` does the same but ignores case and allows the `*` and `?` wildcards. Both return `#VALUE!` if the text isn't there.

```formula
=FIND("-", "SO-1001")
→ 3
=SEARCH("lamp", "Desk Lamp")
→ 6
```

`SUBSTITUTE(text, old, new)` replaces every occurrence of one piece of text with another. It's the workhorse for stripping unwanted characters. Nest it to remove several:

```formula
=SUBSTITUTE(SUBSTITUTE("(512) 555-0142", "(", ""), ")", "")
→ 512 555-0142
```

`REPLACE(text, start, n, new)` replaces by position instead of by content.

## Splitting text

Older workbooks split text with combinations of `LEFT`, `MID` and `FIND`. Modern Excel has functions that say what you mean:

| Function | Returns |
|---|---|
| `TEXTBEFORE(text, delimiter)` | everything before the delimiter |
| `TEXTAFTER(text, delimiter)` | everything after it |
| `TEXTSPLIT(text, col_delimiter, [row_delimiter])` | every piece, spilled across columns |

```formula
=TEXTBEFORE(Leads[Location], ",")
→ {"Austin"; "Denver"; "Seattle"; "Boston"; "Chicago"; "Miami"}
=TEXTAFTER(Leads[Location], ", ")
→ {"TX"; "CO"; "WA"; "MA"; "IL"; "FL"}
=TEXTBEFORE(Staff[Email], "@")
→ {"ava.thompson"; "ben.okafor"; "chloe.martin"; "dev.rao"; "elena.petrova"; "farid.haddad"; "grace.lee"; "hiro.tanaka"}
```

`TEXTSPLIT` breaks one piece of text into many:

```formula
=TEXTSPLIT("Desk Lamp, Notebook, Cable Kit", ", ")
→ {"Desk Lamp", "Notebook", "Cable Kit"}
```

Both `TEXTBEFORE` and `TEXTAFTER` take an optional instance number. A negative one counts from the end, so `TEXTAFTER(name, " ", -1)` gives the last word of a name however many words come before it.

## Building text

The `&` operator (chapter A3) joins values. Two functions help when there are many pieces:

- `CONCAT(text1, ...)` joins everything with nothing in between.
- `TEXTJOIN(delimiter, ignore_empty, text1, ...)` puts a separator between the pieces and can skip blanks. It's the one you'll actually use.

```formula
=TEXTJOIN(", ", TRUE, Products[Product])
→ Desk Lamp, Notebook, Gel Pen Set, Monitor Stand, Desk Chair, Cable Kit, Sticky Notes, Whiteboard
```

To put a line break inside a cell, join `CHAR(10)` and turn on **Wrap Text** for the cell.

## Formatting numbers as text: TEXT

`TEXT(value, format_text)` turns a number into text using a number format code, the same codes you'd use in **Format Cells**:

| Code | 1234.5 shows as |
|---|---|
| `"0"` | 1235 |
| `"0.00"` | 1234.50 |
| `"#,##0"` | 1,235 |
| `"$#,##0.00"` | $1,234.50 |
| `"0%"` | 123450% |
| `"00000"` | 01235 |

```formula
="Revenue to date: " & TEXT(SUM(Sales[Revenue]), "$#,##0")
→ Revenue to date: $11,784
=TEXT(DATE(2026, 3, 9), "dddd d mmmm")
→ Monday 9 March
```

Date codes: `d` day, `dd` two-digit day, `ddd` short weekday, `dddd` full weekday, `m`/`mm`/`mmm`/`mmmm` the same for months, `yy`/`yyyy` the year.

::: gotcha
The result of `TEXT` is text. It looks like a number but won't add up. Use `TEXT` for labels and messages, and keep the real numbers in their own cells for calculating.
:::

## Key ideas

- `LEFT`, `RIGHT`, `MID` and `LEN` work by position; `FIND` and `SEARCH` locate text.
- `TRIM` and `PROPER` fix most messy names; `SUBSTITUTE` strips unwanted characters.
- `TEXTBEFORE`, `TEXTAFTER` and `TEXTSPLIT` split text by a delimiter.
- `TEXTJOIN` builds lists; `TEXT` formats numbers and dates inside text. Both return text, not numbers.

## Exercises

Use sheet **B3** of the practice workbook.

:::: exercise
Tidy up the lead names: remove extra spaces and capitalise each word. Your formula should spill all six names.
::: answer
```formula
=PROPER(TRIM(Leads[Name]))
→ {"Maria Garcia"; "James O'Neil"; "Priya Sharma"; "Li Wei"; "Sofia Rossi"; "Noah Johnson"}
```
:::
::::

:::: exercise
Get each lead's state code: the part of Location after the comma and space.
::: answer
```formula
=TEXTAFTER(Leads[Location], ", ")
→ {"TX"; "CO"; "WA"; "MA"; "IL"; "FL"}
```
`RIGHT(Leads[Location], 2)` works too, because every state code is two letters.
:::
::::

:::: exercise
Get each sales rep's last name from the Name column of the Staff table.
::: answer
```formula
=TEXTAFTER(Staff[Name], " ", -1)
→ {"Thompson"; "Okafor"; "Martin"; "Rao"; "Petrova"; "Haddad"; "Lee"; "Tanaka"}
```
The `-1` means "after the last space".
:::
::::

:::: exercise
List every rep's first name in one cell, separated by a comma and a space.
::: answer
```formula
=TEXTJOIN(", ", TRUE, Staff[Rep])
→ Ava, Ben, Chloe, Dev, Elena, Farid, Grace, Hiro
```
:::
::::

:::: exercise
Turn the order number `SO-1048` into the number 1048, so you could do arithmetic with it.
::: answer
```formula
=VALUE(TEXTAFTER("SO-1048", "-"))
→ 1048
```
:::
::::
