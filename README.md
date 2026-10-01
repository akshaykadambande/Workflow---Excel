# Workflow---Excel
Excel vs Python vs SQL

| Task | Excel | Python | SQL |
| --- | --- | --- | --- |
| Filter | Filter | Boolean filtering | WHERE |
| Sort | Sort | `sort_values()` | ORDER BY |
| Group | PivotTable | `groupby()` | GROUP BY |
| Sum | SUM/SUMIFS | `sum()` | SUM() |
| Average | AVERAGE/AVERAGEIFS | `mean()` | AVG() |
| Median | MEDIAN | `median()` | Database-specific |
| Conditional | IF/IFS | `np.where()` | CASE |
| Lookup | XLOOKUP | merge/map | JOIN |
| Missing | IF/Find/Replace | `fillna()` | COALESCE |
| Duplicates | Remove Duplicates | `drop_duplicates()` | GROUP BY / ROW_NUMBER |
| Date year | YEAR | `.dt.year` | YEAR() |
| Date month | MONTH | `.dt.month` | MONTH() |
| Ranking | RANK | `.rank()` | RANK() |
| Pivot | PivotTable | `pivot_table()` | conditional aggregation / PIVOT where supported |
| Combine rows | Append/Power Query | `concat()` | UNION ALL |
| Combine columns | formulas/Power Query | merge/concat | JOIN |
| Repeatable cleaning | Power Query | Python script | SQL query/view |
Data Analyst — Excel Interview & Practical Handbook

A practical Excel reference for Data Analyst interview preparation and real-world data analysis.

The handbook follows this workflow:

**Import → Inspect → Clean → Transform → Validate → Analyze → Visualize → Report**

* * *
# 19. Recommended Excel Learning Sequence

    Excel Basics
          ↓
    Data Import
          ↓
    Data Inspection
          ↓
    Data Cleaning
          ↓
    Text Functions
          ↓
    Date/Time Functions
          ↓
    Logical Functions
          ↓
    Lookups
          ↓
    Aggregation
          ↓
    Conditional Formatting
          ↓
    Pivot Tables
          ↓
    Power Query
          ↓
    Frequent Cleaning Scenarios
          ↓
    Data Analyst Business Problems

# 1. Excel Repository Structure

    Excel
    │
    ├── 01_Excel_Basics
    ├── 02_Data_Import
    ├── 03_Data_Inspection
    ├── 04_Data_Cleaning
    ├── 05_Text_Operations
    ├── 06_Date_Time_Functions
    ├── 07_Logical_Functions
    ├── 08_Lookup_Reference
    ├── 09_Aggregation_Functions
    ├── 10_Conditional_Formatting
    ├── 11_Data_Analysis
    ├── 12_Power_Query
    └── 13_Frequent_Data_Cleaning_Scenarios

* * *

# 2. Common Dataset

The examples use a common e-commerce dataset.

| Order ID | Customer | City | Category | Order Date | Quantity | Unit Price | Discount | Rating |
| --- | --- | --- | --- | --- | --- | --- | --- | --- |
| 1001 | Amit | Pune | Electronics | 01-Apr-2024 | 1   | 65000 | 5%  | 4.5 |
| 1002 | Amit | Pune | Electronics | 05-Apr-2024 | 2   | 800 | 0%  | 4.0 |
| 1003 | Priya | Mumbai | Furniture | 07-Apr-2024 | 1   | 6500 | 10% | 3.5 |
| 1004 | Rahul | Delhi | Fashion | 12-Apr-2024 | 2   | 4000 | 5%  | 4.2 |
| 1005 | Neha | Pune | Furniture | 02-May-2024 | 1   | 10000 | 0%  | 4.8 |
| 1006 | Arjun | Bengaluru | Electronics | 08-May-2024 | 3   | 800 | 0%  | 3.9 |
| 1007 | Amit | Pune | Fashion | 15-May-2024 | 1   | 4000 | 10% | 4.1 |
| 1008 | Priya | Mumbai | Electronics | 01-Jun-2024 | 1   | 65000 | 0%  | 4.7 |

Add a calculated `Sales` column:

    =F2*G2*(1-H2)

Where:

* `F` = Quantity
* `G` = Unit Price
* `H` = Discount

* * *

# 3. 01_Excel_Basics

## Cell References

### Relative

    =A2*B2

Changes when copied.

### Absolute

    =A2*$B$1

`$B$1` remains fixed.

### Mixed

    =A2*B$1

Row remains fixed.

    =A2*$B1

Column remains fixed.

## Important shortcuts

| Shortcut | Purpose |
| --- | --- |
| Ctrl + C | Copy |
| Ctrl + V | Paste |
| Ctrl + X | Cut |
| Ctrl + Z | Undo |
| Ctrl + F | Find |
| Ctrl + H | Replace |
| Ctrl + 1 | Format Cells |
| Ctrl + Shift + L | Filter |
| Ctrl + Arrow | Jump to data edge |
| Ctrl + Shift + Arrow | Select data range |
| Alt + = | AutoSum |
| F2  | Edit cell |
| Ctrl + T | Convert range to Table |

* * *

# 4. 02_Data_Import

## CSV

    Data → From Text/CSV

Check:

* Delimiter
* Encoding
* Data types
* Headers
* Null values

## Excel workbook

    Data → Get Data → From Workbook

## Web

    Data → Get Data → From Web

## Multiple files

Power Query can combine files from a folder.

Typical workflow:

    Data
     ↓
    Get Data
     ↓
    From Folder
     ↓
    Combine & Transform
     ↓
    Clean
     ↓
    Load

* * *

# 5. 03_Data_Inspection

Before cleaning, inspect the dataset.

## Check dimensions

Identify:

* Number of rows
* Number of columns
* Header names
* Empty rows
* Empty columns

## Check missing values

Use:

* Filter
* Go To Special
* Conditional Formatting
* `COUNTBLANK()`

    =COUNTBLANK(I2:I1000)

## Count non-empty values

    =COUNTA(A2:A1000)

## Count numeric values

    =COUNT(F2:F1000)

## Identify unique categories

    =UNIQUE(D2:D1000)

## Count each category

    =COUNTIF(D:D,D2)

## Inspect duplicates

    =COUNTIF(A:A,A2)

If result > 1, the value appears multiple times.

* * *

# 6. 04_Data_Cleaning

The objective:

**Fix → Standardize → Validate → Prepare**

* * *

# 6.1 Missing Values

## Identify blanks

    =ISBLANK(I2)

## Count blanks

    =COUNTBLANK(I:I)

## Replace blank with constant

    =IF(I2="","Unknown",I2)

## Replace blank with mean

    =IF(I2="",AVERAGE($I$2:$I$1000),I2)

## Replace blank with median

    =IF(I2="",MEDIAN($I$2:$I$1000),I2)

## Replace blank with mode

    =IF(I2="",MODE.SNGL($I$2:$I$1000),I2)

## Fill missing Rating by Category median

Suppose:

* Category = column D
* Rating = column I

Modern Excel:

    =IF(
        I2="",
        MEDIAN(
            FILTER(
                $I$2:$I$1000,
                ($D$2:$D$1000=D2)*
                ($I$2:$I$1000<>"")
            )
        ),
        I2
    )

This is an important interview scenario because a category-specific statistic can preserve differences between groups.

* * *

# 6.2 Duplicates

## Highlight duplicates

    Home
    → Conditional Formatting
    → Highlight Cells Rules
    → Duplicate Values

## Formula

    =COUNTIF($A:$A,A2)>1

## Remove duplicates

    Data
    → Remove Duplicates

Select the columns that define a duplicate record.

Important:

Do not automatically remove duplicates simply because a value repeats. Determine what defines a duplicate business record.

* * *

# 6.3 Text Cleaning

## Remove leading/trailing/repeated spaces

    =TRIM(A2)

## Remove non-printing characters

    =CLEAN(A2)

## Best combined cleaning

    =TRIM(CLEAN(A2))

## Uppercase

    =UPPER(A2)

## Lowercase

    =LOWER(A2)

## Proper case

    =PROPER(A2)

## Replace

    =SUBSTITUTE(A2,"Old","New")

## Find text

    =FIND("@",A2)

## Search text

    =SEARCH("@",A2)

* * *

# 6.4 Standardize Categories

Suppose values contain:

    Electronics
    electronics
     ELECTRONICS
    Electronic

Clean spaces:

    =TRIM(A2)

Standardize case:

    =PROPER(TRIM(A2))

Map inconsistent values:

    =IF(
        OR(A2="Electronic",A2="Electronics"),
        "Electronics",
        A2
    )

For many mappings, use a lookup table rather than a long nested IF.

* * *

# 6.5 Numbers Stored as Text

## Detect

    =ISNUMBER(A2)

## Convert

    =VALUE(A2)

## Remove commas

    =VALUE(SUBSTITUTE(A2,",",""))

Example:

    "65,000"

becomes:

    65000

## Remove currency symbol

    =VALUE(SUBSTITUTE(A2,"₹",""))

* * *

# 6.6 Date Cleaning

Convert text to date:

    =DATEVALUE(A2)

If date contains date and time:

    =INT(A2)

Format:

    Home → Number → Short Date

or:

    Format Cells → Date

## Validate date

    =ISNUMBER(A2)

Excel stores valid dates as serial numbers, so this can help identify text pretending to be dates.

* * *

# 6.7 Date & Time Extraction

Assume:

    A2 = 15-May-2024 14:30:25

## Date

    =INT(A2)

## Time

    =MOD(A2,1)

## Year

    =YEAR(A2)

## Month number

    =MONTH(A2)

## Month name

    =TEXT(A2,"mmmm")

## Short month name

    =TEXT(A2,"mmm")

## Day

    =DAY(A2)

## Day of week number

    =WEEKDAY(A2,2)

`2` makes Monday = 1 and Sunday = 7.

## Day name

    =TEXT(A2,"dddd")

## Short day name

    =TEXT(A2,"ddd")

## Quarter

    ="Q"&ROUNDUP(MONTH(A2)/3,0)

## Week number

    =WEEKNUM(A2,2)

* * *

# 6.8 Numerical Cleaning

## Remove commas

    =VALUE(SUBSTITUTE(A2,",",""))

## Round

    =ROUND(A2,2)

## Round up

    =ROUNDUP(A2,2)

## Round down

    =ROUNDDOWN(A2,2)

## Absolute value

    =ABS(A2)

## Percentage

    =A2/B2

Format result as Percentage.

* * *

# 6.9 Outliers

## Quartile

    =QUARTILE.INC(A2:A1000,1)

    =QUARTILE.INC(A2:A1000,3)

## IQR

    =Q3-Q1

## Lower boundary

    =Q1-1.5*IQR

## Upper boundary

    =Q3+1.5*IQR

## Flag outlier

    =IF(
        OR(
            A2<Lower_Bound,
            A2>Upper_Bound
        ),
        "Outlier",
        "Normal"
    )

Do not automatically delete an outlier. First determine whether it is:

* A data-entry error
* A legitimate extreme value
* A rare business event
* A measurement issue

* * *

# 6.10 Rows & Columns

## Rename columns

Double-click header or type a standardized name.

Recommended:

    Customer Name → Customer_Name
    Order Date → Order_Date
    Review Rating → Review_Rating

## Delete unnecessary columns

    Select column
    → Right Click
    → Delete

## Filter rows

    Data → Filter

## Sort

    Data → Sort

## Remove blank rows

    Filter → Blanks
    → Select rows
    → Delete

* * *

# 6.11 Split Columns

Use:

    Data → Text to Columns

Common delimiters:

* Comma
* Space
* Colon
* Hyphen
* Semicolon

Example:

    Pune, Maharashtra, 411001

can become:

    City       State          PIN
    Pune       Maharashtra    411001

* * *

# 6.12 Flash Fill

Example:

    A2 = Amit Kadambande

Enter:

    Amit

in another column, then:

    Ctrl + E

Excel recognizes the pattern.

Useful for:

* Extracting names
* Creating initials
* Standardizing formats
* Splitting text
* Combining text patterns

* * *

# 7. 05_Text_Operations

## CONCAT

    =CONCAT(A2," ",B2)

## TEXTJOIN

    =TEXTJOIN(", ",TRUE,A2:C2)

## LEFT

    =LEFT(A2,5)

## RIGHT

    =RIGHT(A2,4)

## MID

    =MID(A2,3,5)

## LEN

    =LEN(A2)

## FIND

    =FIND("-",A2)

## SEARCH

    =SEARCH("pune",A2)

## SUBSTITUTE

    =SUBSTITUTE(A2,"old","new")

## REPLACE

    =REPLACE(A2,1,3,"ABC")

* * *

# 8. 06_Date_Time_Functions

## TODAY

    =TODAY()

## NOW

    =NOW()

## DATE

    =DATE(2024,5,15)

## YEAR

    =YEAR(A2)

## MONTH

    =MONTH(A2)

## DAY

    =DAY(A2)

## WEEKDAY

    =WEEKDAY(A2,2)

## WEEKNUM

    =WEEKNUM(A2,2)

## EDATE

Add months:

    =EDATE(A2,3)

## EOMONTH

End of month:

    =EOMONTH(A2,0)

End of next month:

    =EOMONTH(A2,1)

## DATEDIF

Years:

    =DATEDIF(A2,B2,"Y")

Months:

    =DATEDIF(A2,B2,"M")

Days:

    =DATEDIF(A2,B2,"D")

## NETWORKDAYS

    =NETWORKDAYS(A2,B2)

Useful for working-day calculations.

* * *

# 9. 07_Logical_Functions

## IF

    =IF(F2>=100000,"High","Low")

## Multiple conditions

    =IF(
        F2>=100000,
        "High",
        IF(F2>=50000,"Medium","Low")
    )

## IFS

    =IFS(
        F2>=100000,"High",
        F2>=50000,"Medium",
        TRUE,"Low"
    )

## AND

    =AND(F2>10000,G2>2)

## OR

    =OR(C2="Pune",C2="Mumbai")

## NOT

    =NOT(F2>10000)

## IFERROR

    =IFERROR(
        A2/B2,
        0
    )

## IFNA

    =IFNA(
        XLOOKUP(A2,H:H,I:I),
        "Not Found"
    )

* * *

# 10. 08_Lookup_Reference

## VLOOKUP

    =VLOOKUP(
        A2,
        H2:K100,
        3,
        FALSE
    )

Limitations:

* Lookup column must be first
* Primarily returns values to the right
* Column number can be fragile when structure changes

## XLOOKUP

    =XLOOKUP(
        A2,
        H:H,
        J:J,
        "Not Found"
    )

Advantages:

* Can look left or right
* Easier exact matching
* Built-in not-found result
* More flexible than VLOOKUP

## INDEX + MATCH

    =INDEX(
        J:J,
        MATCH(
            A2,
            H:H,
            0
        )
    )

## MATCH

    =MATCH(
        A2,
        H:H,
        0
    )

## OFFSET

    =OFFSET(
        A2,
        0,
        2
    )

OFFSET is useful for dynamic references but is volatile and can affect workbook performance when heavily used.

## INDIRECT

    =INDIRECT("A"&B2)

Useful for dynamically constructing references, but also volatile.

* * *

# 11. 09_Aggregation_Functions

## SUM

    =SUM(F2:F1000)

## AVERAGE

    =AVERAGE(F2:F1000)

## MEDIAN

    =MEDIAN(F2:F1000)

## MIN

    =MIN(F2:F1000)

## MAX

    =MAX(F2:F1000)

## COUNT

    =COUNT(F2:F1000)

## COUNTA

    =COUNTA(A2:A1000)

## COUNTBLANK

    =COUNTBLANK(A2:A1000)

## COUNTIF

    =COUNTIF(D:D,"Electronics")

## COUNTIFS

    =COUNTIFS(
        D:D,"Electronics",
        F:F,">2"
    )

## SUMIF

    =SUMIF(
        D:D,
        "Electronics",
        I:I
    )

## SUMIFS

    =SUMIFS(
        I:I,
        D:D,"Electronics",
        C:C,"Pune"
    )

## AVERAGEIF

    =AVERAGEIF(
        D:D,
        "Electronics",
        I:I
    )

## AVERAGEIFS

    =AVERAGEIFS(
        I:I,
        D:D,"Electronics",
        C:C,"Pune"
    )

* * *

# 12. 10_Conditional_Formatting

Conditional Formatting is primarily for visual detection and communication.

## Highlight duplicates

    Home
    → Conditional Formatting
    → Highlight Cells Rules
    → Duplicate Values

## Highlight blanks

Use formula:

    =ISBLANK(A2)

## Highlight values above threshold

    =A2>100000

## Top 10

    Conditional Formatting
    → Top/Bottom Rules
    → Top 10 Items

## Data Bars

Useful for comparing:

* Sales
* Quantity
* Profit
* Ratings

## Color Scales

Useful for:

* Performance matrices
* Heatmaps
* Category comparisons

## Icon Sets

Useful for:

* KPI direction
* Growth
* Performance status

* * *

# 13. 11_Data_Analysis

## Sort & Filter

    Data → Sort
    Data → Filter

## Advanced Filter

Useful when filtering based on complex criteria.

## Pivot Table

Basic workflow:

    Select dataset
     ↓
    Insert → PivotTable
     ↓
    Rows → Category
    Values → Sales

Example:

    Category       Total Sales
    Electronics    ₹...
    Furniture      ₹...
    Fashion        ₹...

## Pivot by month

Rows:

    Year
    Month

Values:

    Sales

## Pivot Chart

Use:

    Insert → PivotChart

## Slicers

Useful for interactive filtering by:

* Category
* City
* Product
* Year
* Customer segment

## What-If Analysis

### Goal Seek

Useful when you know the desired output and want Excel to determine the required input.

Example:

    Desired Profit = ₹100,000
    What selling price is required?

### Scenario Manager

Compare multiple assumptions.

* * *

# 14. 12_Power_Query

Power Query is especially useful for repeatable data-cleaning workflows.

## Typical workflow

    Get Data
     ↓
    Power Query Editor
     ↓
    Inspect
     ↓
    Remove errors
     ↓
    Remove duplicates
     ↓
    Change types
     ↓
    Replace values
     ↓
    Split/merge columns
     ↓
    Add columns
     ↓
    Group
     ↓
    Merge/Append
     ↓
    Close & Load

## Remove duplicates

    Home
    → Remove Rows
    → Remove Duplicates

## Remove blank rows

    Home
    → Remove Rows
    → Remove Blank Rows

## Change type

    Transform
    → Data Type

## Replace values

    Transform
    → Replace Values

## Split column

    Transform
    → Split Column
    → By Delimiter

## Merge columns

    Transform
    → Merge Columns

## Merge queries

Equivalent conceptually to joining datasets.

    Home
    → Merge Queries

## Append queries

Equivalent to stacking datasets vertically.

    Home
    → Append Queries

## Group By

Useful for:

* Total sales by category
* Average rating by product
* Count of customers by city

## Conditional Column

    Add Column
    → Conditional Column

* * *

# 15. 13_Frequent_Data_Cleaning_Scenarios

This folder contains complete business-style cleaning problems.

* * *

## Scenario 1 — Clean a Completely Messy Dataset

Workflow:

    Raw Data
     ↓
    Inspect rows/columns
     ↓
    Identify blanks
     ↓
    Check duplicates
     ↓
    Fix data types
     ↓
    Standardize text
     ↓
    Fix categories
     ↓
    Fix dates
     ↓
    Fix numeric values
     ↓
    Handle outliers
     ↓
    Validate
     ↓
    Clean Dataset

* * *

## Scenario 2 — Fill Rating Nulls With Category Median

Assume:

* Category = D
* Rating = I

    =IF(
        I2="",
        MEDIAN(
            FILTER(
                $I$2:$I$1000,
                ($D$2:$D$1000=D2)*
                ($I$2:$I$1000<>"")
            )
        ),
        I2
    )

Business reasoning:

Do not blindly use the overall median if categories have materially different rating distributions.

* * *

## Scenario 3 — Remove Duplicates

    Data
    → Remove Duplicates

First determine the business key.

Example:

    Order ID

may define uniqueness better than:

    Customer Name

* * *

## Scenario 4 — Standardize Category

Problem:

    electronics
    Electronics
     ELECTRONICS
    Electronic

Solution:

    =PROPER(TRIM(A2))

Then map `"Electronic"` to `"Electronics"` using XLOOKUP or a mapping table.

* * *

## Scenario 5 — Fix Numbers Stored as Text

Problem:

    "65,000"
    "₹50,000"
    "1000"

Solution:

    =VALUE(
        SUBSTITUTE(
            SUBSTITUTE(A2,",",""),
            "₹",""
        )
    )

* * *

## Scenario 6 — Separate Date and Time

If A2 contains:

    15-May-2024 14:30:25

Date:

    =INT(A2)

Time:

    =MOD(A2,1)

* * *

## Scenario 7 — Extract All Date Components

    =YEAR(A2)

    =MONTH(A2)

    =TEXT(A2,"mmmm")

    =DAY(A2)

    =WEEKDAY(A2,2)

    =TEXT(A2,"dddd")

    ="Q"&ROUNDUP(MONTH(A2)/3,0)

* * *

## Scenario 8 — Split Address

Example:

    Pune, Maharashtra, 411001

Use:

    Data
    → Text to Columns
    → Delimited
    → Comma

Or use:

    =TEXTSPLIT(A2,",")

* * *

## Scenario 9 — Remove Extra Spaces

    =TRIM(A2)

For hidden/non-printing characters:

    =TRIM(CLEAN(A2))

* * *

## Scenario 10 — Find Invalid Categories

Create a valid category list:

    Electronics
    Furniture
    Fashion

Then use Data Validation or:

    =COUNTIF(
        $M$2:$M$4,
        D2
    )

If result is 0, the category is not in the approved list.

* * *

## Scenario 11 — Create Performance Band

    =IFS(
        I2>=4.5,"Excellent",
        I2>=4,"Good",
        I2>=3,"Average",
        TRUE,"Poor"
    )

* * *

## Scenario 12 — Calculate Sales

    =Quantity*Unit_Price*(1-Discount)

Example:

    =F2*G2*(1-H2)

* * *

## Scenario 13 — Calculate Profit

If cost is available:

    =Sales-(Quantity*Cost)

* * *

## Scenario 14 — Top 5 Products

Use:

    Sort Largest to Smallest

or:

    =LARGE(
        Sales_Range,
        ROWS($A$1:A1)
    )

Modern Excel:

    =TAKE(
        SORTBY(
            A2:I1000,
            I2:I1000,
            -1
        ),
        5
    )

* * *

## Scenario 15 — Sales by Category

Using PivotTable:

    Rows:
    Category
    
    Values:
    Sum of Sales

Using SUMIF:

    =SUMIF(
        D:D,
        "Electronics",
        J:J
    )

* * *

## Scenario 16 — Average Rating by Category

    =AVERAGEIF(
        D:D,
        "Electronics",
        I:I
    )

* * *

## Scenario 17 — Count Orders by Category

    =COUNTIF(
        D:D,
        "Electronics"
    )

* * *

## Scenario 18 — Count Electronics Orders From Pune

    =COUNTIFS(
        D:D,
        "Electronics",
        C:C,
        "Pune"
    )

* * *

## Scenario 19 — Total Sales for Electronics in Pune

    =SUMIFS(
        J:J,
        D:D,
        "Electronics",
        C:C,
        "Pune"
    )

* * *

## Scenario 20 — Find Missing Lookup Information

    =XLOOKUP(
        A2,
        Product_ID_Range,
        Product_Name_Range,
        "Not Found"
    )

* * *

# 16. Data Cleaning Checklist

Before analysis:

    [ ] Correct number of rows
    [ ] Correct number of columns
    [ ] Correct headers
    [ ] Blank rows removed
    [ ] Blank columns removed
    [ ] Missing values identified
    [ ] Missing-value strategy selected
    [ ] Duplicates checked
    [ ] Business-key duplicates checked
    [ ] Text standardized
    [ ] Extra spaces removed
    [ ] Categories standardized
    [ ] Numbers converted correctly
    [ ] Currency symbols handled
    [ ] Dates converted correctly
    [ ] Date/time separated if required
    [ ] Invalid dates checked
    [ ] Outliers investigated
    [ ] Invalid categories checked
    [ ] Formulas validated
    [ ] Final dataset spot-checked

* * *

# 17. Excel Data Analyst Interview Questions

## Basics

1. Difference between relative and absolute references?
2. What is an Excel Table?
3. How do filters work?
4. What is the difference between COUNT, COUNTA and COUNTBLANK?
5. How do you find duplicates?
6. How do you handle missing values?

## Cleaning

7. How would you clean a messy dataset?
8. How do you remove extra spaces?
9. How do you standardize categories?
10. How do you convert numbers stored as text?
11. How do you fix inconsistent dates?
12. How do you identify outliers?
13. How would you fill missing values?
14. When would you use mean vs median?
15. How do you split an address column?

## Functions

16. VLOOKUP vs XLOOKUP?
17. XLOOKUP vs INDEX-MATCH?
18. SUMIF vs SUMIFS?
19. COUNTIF vs COUNTIFS?
20. IF vs IFS?
21. IFERROR vs IFNA?
22. What does TEXTJOIN do?
23. How do you extract year/month/day?
24. How do you calculate working days?

## Analysis

25. How do you calculate sales by category?
26. How do you find top 5 products?
27. How do you calculate percentage contribution?
28. How do you calculate month-over-month growth?
29. How do you create a PivotTable?
30. When would you use a PivotTable instead of formulas?

## Power Query

31. What is Power Query?
32. Why use Power Query instead of manual cleaning?
33. Merge vs Append?
34. How do you remove duplicates in Power Query?
35. How do you change data types?
36. How do you create a conditional column?
37. How do you combine files from a folder?
38. What does Refresh do?

* * *



* * *



* * *

# 20. Golden Rule

Do not memorize Excel formulas in isolation.

For every problem, ask:

**What is the business question?**

Then:

**What data do I have → What is wrong with it → How do I clean it → What formula/tool should I use → How do I validate the result → What business insight does it provide?**

That is the mindset that turns Excel knowledge into Data Analyst capability.
