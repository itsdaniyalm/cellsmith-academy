---
title: Dates and times
summary: Dates are day counts, times are fractions of a day. Building dates, pulling them apart, month ends, working days and ages.
fx: =NETWORKDAYS(DATE(2026, 1, 1), DATE(2026, 6, 30))
---

Chapter A1 showed that Excel stores a date as the number of days since 1 January 1900. Hold onto that idea, because it explains everything in this chapter. Subtracting two dates gives days. Adding 7 to a date gives next week. A time is a fraction of a day, so 6 p.m. is 0.75.

## Building a date

`DATE(year, month, day)` turns three numbers into a date:

```formula
=DATE(2026, 3, 15)
→ 2026-03-15
```

Typing `3/15/2026` into a cell also creates a date, but how Excel reads typed dates depends on your computer's regional settings: in the UK, 3/4 is 3 April; in the US, it's 4 March. `DATE` is the same everywhere, so use it inside formulas.

`DATE` quietly rolls over out-of-range values, which is handy:

```formula
=DATE(2026, 13, 1)
→ 2027-01-01
=DATE(2026, 3, 0)
→ 2026-02-28
```

Month 13 is January of the next year. Day 0 is the last day of the previous month.

::: note
How a date *looks* depends on the cell's number format and your computer's regional settings. This course always shows dates as year-month-day, like 2026-03-15, so they read the same everywhere. Your Excel may show 15/03/2026 or 3/15/2026; the stored number is the same.
:::

## Taking a date apart

| Function | Returns, for 15 March 2026 |
|---|---|
| `YEAR(date)` | 2026 |
| `MONTH(date)` | 3 |
| `DAY(date)` | 15 |
| `WEEKDAY(date, 2)` | 7 (Sunday, counting Monday as 1) |

```formula
=MONTH(Sales[Date])
≈ {1; 1; 1; 1; 1; 1; 1; 1; 2; …}
```

A quarter isn't a built-in function, but it's just the month divided by 3, rounded up:

```formula
="Q" & ROUNDUP(MONTH(DATE(2026, 5, 20)) / 3, 0)
→ Q2
```

## Today

`TODAY()` returns the current date and `NOW()` the current date and time. They're volatile (chapter A4): they update whenever the workbook recalculates. Use them for "days until" and "age" calculations, but if you need to record *when* something happened, type the date in (<kbd>Ctrl</kbd>+<kbd>;</kbd> enters today's date as a fixed value).

```formula
=DATE(2026, 12, 25) - TODAY()
≈ 178
```

## Date arithmetic

Because dates are numbers, `+` and `-` do most of the work. The gap between order and shipping, for one order:

```formula
=DATE(2026, 1, 9) - DATE(2026, 1, 2)
→ 7
```

Across the whole table, the shipping gap is `Sales[Ship Date] - Sales[Date]`. But orders that haven't shipped have a blank ship date, and a blank counts as zero, which would give a huge negative gap. Leave them out with `IF`:

```formula
=AVERAGE(IF(Sales[Ship Date] <> "", Sales[Ship Date] - Sales[Date]))
→ 4
```

`IF` without a false branch returns `FALSE` for unshipped orders, and `AVERAGE` ignores `FALSE` values inside an array. That's the average number of days to ship.

## Months and month ends

Adding months is harder than adding days, because months differ in length. Two functions handle it:

- `EDATE(date, months)`: the same day, some months later (or earlier, with a negative number).
- `EOMONTH(date, months)`: the last day of the month, some months later. `EOMONTH(date, 0)` is the end of the date's own month.

```formula
=EDATE(DATE(2026, 1, 31), 1)
→ 46081
=EOMONTH(DATE(2026, 2, 10), 0)
→ 46081
=EOMONTH(DATE(2026, 2, 10), -1) + 1
→ 46054
```

::: gotcha
`EDATE`, `EOMONTH` and `WORKDAY` return the date's serial number, and Excel doesn't format the cell as a date for you. 46081 is 2026-02-28 and 46054 is 2026-02-01. Format the cell as a date (**Home → Number Format → Short Date**), or wrap the formula in `TEXT(..., "yyyy-mm-dd")`.
:::

Formatted as dates, the first two are both 28 February 2026: there's no 31 February, so `EDATE` stops at the month's last day. The third is a common trick: the end of the previous month plus one day is the first day of this month.

## Working days

`NETWORKDAYS(start, end, [holidays])` counts weekdays between two dates, including both ends. `WORKDAY(start, days, [holidays])` goes the other way: the date that's a number of working days after the start.

```formula
=NETWORKDAYS(DATE(2026, 1, 1), DATE(2026, 6, 30))
→ 129
=TEXT(WORKDAY(DATE(2026, 6, 26), 3), "yyyy-mm-dd")
→ 2026-07-01
```

Pass a range of holiday dates as the last argument and they're skipped too. For weekends other than Saturday and Sunday, use `NETWORKDAYS.INTL` and `WORKDAY.INTL`.

## Ages and tenure

`DATEDIF(start, end, unit)` counts complete years (`"y"`), months (`"m"`) or days (`"d"`) between two dates. Excel doesn't list it in its function menus, for historical reasons, but it works and it's the simplest way to get an age:

```formula
=DATEDIF(Staff[Hire Date], DATE(2026, 6, 30), "y")
→ {7; 4; 8; 3; 5; 4; 9; 1}
```

That's each rep's complete years of service at the end of June. `YEARFRAC(start, end)` gives a fractional number of years instead, like 7.3.

## Times

Times are fractions of a day: 0.5 is noon, 0.25 is 6 a.m. `TIME(hours, minutes, seconds)` builds one, and `HOUR`, `MINUTE` and `SECOND` take one apart.

```formula
=TIME(18, 0, 0)
→ 18:00
=TIME(18, 0, 0) * 1
→ 0.75
```

To turn a time into hours you can calculate with, multiply by 24. When adding up durations longer than a day, format the cell as `[h]:mm`. The square brackets stop the hours from wrapping back to zero at 24.

## Key ideas

- Dates are day numbers and times are fractions of a day, so ordinary arithmetic works on them.
- Build dates with `DATE(year, month, day)`; take them apart with `YEAR`, `MONTH`, `DAY` and `WEEKDAY`.
- `EDATE` and `EOMONTH` move by months; `NETWORKDAYS` and `WORKDAY` count working days.
- `DATEDIF(start, end, "y")` gives complete years between two dates.

## Exercises

Use sheet **B4** of the practice workbook.

:::: exercise
How many orders were placed in March 2026? Use `COUNTIFS` with a start and end date.
::: answer
```formula
=COUNTIFS(Sales[Date], ">=" & DATE(2026, 3, 1), Sales[Date], "<" & DATE(2026, 4, 1))
→ 9
```
:::
::::

:::: exercise
What is the last day of the month in which the first order was placed? Show it as text in the form `2026-03-31`.
::: answer
```formula
=TEXT(EOMONTH(MIN(Sales[Date]), 0), "yyyy-mm-dd")
→ 2026-01-31
```
:::
::::

:::: exercise
What is the longest time any order took to ship, in days?
::: answer
```formula
=MAX(IF(Sales[Ship Date] <> "", Sales[Ship Date] - Sales[Date]))
→ 7
```
:::
::::

:::: exercise
How many working days are there in the second quarter of 2026 (1 April to 30 June)?
::: answer
```formula
=NETWORKDAYS(DATE(2026, 4, 1), DATE(2026, 6, 30))
→ 65
```
:::
::::

:::: exercise
Which year was the longest-serving rep hired? Find the earliest hire date and take its year.
::: answer
```formula
=YEAR(MIN(Staff[Hire Date]))
→ 2017
```
:::
::::
