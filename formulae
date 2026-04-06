# Apple Numbers Budget Template -- Spilling Array Formulas

A version-controlled reference for all spilling array functions used in the
Budget Template Numbers project. Formulas reference the **Spending template**
table with the following column layout:

|Column|Content                                                                                                        |
|------|---------------------------------------------------------------------------------------------------------------|
|A     |Budget Item (title)                                                                                            |
|B     |Category                                                                                                       |
|C     |Payment Interval (`Yearly`, `Discretionary`, `Weekly`, `Monthly`, `Half-yearly`, `Two-monthly`, `Three-yearly`)|
|D     |Budgeted Amount                                                                                                |
|E     |Estimated Monthly Payment                                                                                      |
|G     |Payment Month (number 1–12)                                                                                    |
|H     |Payment Day                                                                                                    |

-----

## 1. Estimated Monthly Payment (Spending template Column E)

Per-cell formula -- paste in each cell of the Estimated Monthly Payment column.
Calculates a normalised monthly figure from the Budgeted Amount based on
Payment Interval.

```
=IFS(C2="Monthly", D2, C2="Half-yearly", D2/6, C2="Two-monthly", D2/2, C2="Yearly", D2/12, C2="Discretionary", D2/12, C2="Three-yearly", D2/36, TRUE, "")
```

-----

## 2. Yearly Payments Calendar

A calendar table with Categories as row titles and months (January–December)
as column headers, showing Estimated Monthly Payment totals for **Yearly**
items. Includes a column of row totals and a footer row of column totals.

Place in the top-left cell of the destination table.

```
=LET(
  months, MAP(SEQUENCE(1,12,1,1), LAMBDA(m, CHOOSE(m,"January","February","March","April","May","June","July","August","September","October","November","December"))),
  cats, UNIQUE(FILTER('Spending template'::B,'Spending template'::C="Yearly",)),
  raw, MAKEARRAY(ROWS(cats),COLUMNS(months),LAMBDA(row,col,SUMIFS('Spending template'::E,'Spending template'::B,INDEX(cats,row,1,),'Spending template'::G,MONTH(DATEVALUE("1 "&INDEX(months,1,col,)&" 2000")),'Spending template'::C,"Yearly"))),
  rowTotals, BYROW(raw,LAMBDA(r,SUM(r))),
  colTotals, BYCOL(raw,LAMBDA(c,SUM(c))),
  display, IFERROR(1/(1/raw),""),
  header, HSTACK("Category",months,"Total"),
  table, HSTACK(cats,display,rowTotals),
  footer, HSTACK("Total",colTotals,SUM(rowTotals)),
  VSTACK(header,table,footer)
)
```

**Notes:**

- `MONTH(DATEVALUE("1 "&INDEX(months,1,col,)&" 2000"))` converts month names
  to numbers for matching against Column G.
- `IFERROR(1/(1/raw),"")` displays blank instead of zero.
- `rowTotals` and `colTotals` are calculated from `raw` (not `display`) to
  keep arithmetic on clean numbers.

-----

## 3. Discretionary Summary by Category

Lists each unique Category for Discretionary items with its annual Budgeted
Amount total and a derived monthly average. Includes a totals footer row.

```
=LET(
  cats, UNIQUE(FILTER('Spending template'::B,'Spending template'::C="Discretionary",)),
  totals, SUMIFS('Spending template'::D,'Spending template'::C,"Discretionary",'Spending template'::B,cats),
  monthly, totals/12,
  header, HSTACK("Category","Annual Total","Monthly Average"),
  table, HSTACK(cats,totals,monthly),
  footer, HSTACK("Total",SUM(totals),SUM(monthly)),
  VSTACK(header,table,footer)
)
```

-----

## 4. Weekly Summary by Category

Lists each unique Category for Weekly items with its Budgeted Amount (Column D)
and Estimated Monthly Payment (Column E) totals. Includes a totals footer row.

```
=LET(
  cats, UNIQUE(FILTER('Spending template'::B,'Spending template'::C="Weekly",)),
  annual, SUMIFS('Spending template'::D,'Spending template'::C,"Weekly",'Spending template'::B,cats),
  monthly, SUMIFS('Spending template'::E,'Spending template'::C,"Weekly",'Spending template'::B,cats),
  header, HSTACK("Category","Budgeted Amount","Estimated Monthly"),
  table, HSTACK(cats,annual,monthly),
  footer, HSTACK("Total",SUM(annual),SUM(monthly)),
  VSTACK(header,table,footer)
)
```

-----

## 5. Ruby Items List

Lists all Budget Items whose title contains the text "Ruby" (case-insensitive),
with their Category and Estimated Monthly Payment. Includes a totals footer row.

```
=LET(
  mask, ISNUMBER(SEARCH("Ruby",'Spending template'::A)),
  items, FILTER('Spending template'::A,mask,),
  cats, FILTER('Spending template'::B,mask,),
  amounts, FILTER('Spending template'::D,mask,),
  header, HSTACK("Budget Item","Category","Estimated Monthly"),
  table, HSTACK(items,cats,amounts),
  footer, HSTACK("Total","",SUM(amounts)),
  VSTACK(header,table,footer)
)
```

-----

## 6. Monthly Payments by Day and Category (LET reference)

A matrix showing Monthly payment items, with Payment Day as row titles and
Categories as column headers. Includes row totals, column totals, and a grand
total. Sourced from the more complex LET formula developed in session.

```
=LET(
  days, SORT(UNIQUE(FILTER('Spending template'::H,
    ('Spending template'::H >= 1)
    * ('Spending template'::H <= 31)
    * ('Spending template'::C = "Monthly")
  ,),,),),
  cats, TOROW(SORT(UNIQUE(FILTER('Spending template'::B,
    ('Spending template'::H >= 1)
    * ('Spending template'::H <= 31)
    * ('Spending template'::C = "Monthly")
  ,),,),),),
  raw, SUMIFS('Spending template'::E,
    'Spending template'::H, days,
    'Spending template'::B, cats,
    'Spending template'::C, "Monthly"),
  display, IF(raw = 0, "", raw),
  rowTotals, BYROW(raw, LAMBDA(r, SUM(r))),
  colTotals, BYCOL(raw, LAMBDA(c, SUM(c))),
  grandTotal, SUM(colTotals),
  header, HSTACK("Payment day", cats, "Total"),
  table, HSTACK(days, display, rowTotals),
  footer, HSTACK("Category Totals", colTotals, grandTotal),
  VSTACK(header, table, footer)
)
```