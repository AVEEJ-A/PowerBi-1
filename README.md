Sample Superstore -- Power Query Data Cleaning & Date Normalization

📌 Project Overview

This project cleans and prepares the Sample Superstore dataset using
Power Query (M Language).

The transformation process focuses on:

Cleaning text fields using Text.Trim and Text.Clean

Correcting and standardizing data types

Safely converting date columns

Handling conversion errors without breaking the query

Creating Year, Month, and Quarter fields from Order Date

Creating readable Month and Quarter labels

Removing completely blank rows

Validating the final dataset before loading it into Power BI / Excel

Dataset

Source: samplesuperstore.csv

Rows: 10,194

Columns: 21

Primary date fields: Order Date, Ship Date

Main numeric fields: Sales, Quantity, Discount, Profit

📂 Suggested GitHub Structure

sample-superstore-power-query/
│
├── samplesuperstore.csv
├── README.md
└── PowerQuery/
    └── clean_normalize.m

🔧 Data Cleaning Process

1. Load CSV Data

The CSV is imported using UTF-8 encoding and comma delimiter.

Csv.Document(
    File.Contents("samplesuperstore.csv"),
    [Delimiter=",", Encoding=65001, QuoteStyle=QuoteStyle.Csv]
)

For GitHub/portable projects, avoid hard-coding a local Windows path
such as C:\Users\...\samplesuperstore.csv.

2. Promote Headers

The first row is converted into column headers.

Table.PromoteHeaders(Source, [PromoteAllScalars=true])

3. Data Type Correction

The following data types are applied:

Column           Data Type

Row ID           Whole Number
Order ID         Text
Order Date       Date
Ship Date        Date
Ship Mode        Text
Customer ID      Text
Customer Name    Text
Segment          Text
Country/Region   Text
City             Text
State/Province   Text
Postal Code      Text
Region           Text
Product ID       Text
Category         Text
Sub-Category     Text
Product Name     Text
Sales            Decimal Number
Quantity         Whole Number
Discount         Decimal Number
Profit           Decimal Number

🧹 Text Cleaning

Text columns are cleaned using:

Text.Trim
Text.Clean

This removes:

Leading spaces

Trailing spaces

Unwanted non-printing characters

Example:

"  Technology  "

becomes:

"Technology"

This is important for avoiding duplicate categories caused by invisible
or extra spaces.

📅 Date Normalization

The following date fields are standardized:

Order Date

Ship Date

The source dataset uses dates in MM-DD-YYYY format.

The query uses en-US culture during date conversion so that dates are
interpreted consistently.

Example:

01-04-2023

is interpreted as:

January 4, 2023

🗓️ Year, Month & Quarter Columns

The cleaned Order Date is used to create calendar attributes.

Year

Date.Year([Order Date])

Example:

2023

Month Number

Date.Month([Order Date])

Example:

1

Month Name

Date.MonthName([Order Date], "en-US")

Example:

January

Quarter Number

Date.QuarterOfYear([Order Date])

Example:

1

Quarter Name

"Q" & Text.From(Date.QuarterOfYear([Order Date]))

Example:

Q1

⚠️ Error Correction

A common Power Query problem occurs when a column contains an invalid
date or number.

Instead of allowing the entire transformation to fail, this project uses
safe conversions with try ... otherwise.

Example:

try Date.FromText(Text.Trim(_), "en-US") otherwise null

If an invalid date is found, Power Query returns null instead of
generating an error.

This makes the transformation more robust when the source CSV changes.

✅ Complete Power Query M Code

Use the following code in Power Query → Advanced Editor.

let
    // ============================================================
    // 1. SOURCE
    // ============================================================
    Source =
        Csv.Document(
            File.Contents("samplesuperstore.csv"),
            [
                Delimiter = ",",
                Columns = 21,
                Encoding = 65001,
                QuoteStyle = QuoteStyle.Csv
            ]
        ),

    // ============================================================
    // 2. PROMOTE HEADERS
    // ============================================================
    #"Promoted Headers" =
        Table.PromoteHeaders(
            Source,
            [PromoteAllScalars = true]
        ),

    // ============================================================
    // 3. REMOVE COMPLETELY BLANK ROWS
    // ============================================================
    #"Removed Blank Rows" =
        Table.SelectRows(
            #"Promoted Headers",
            each
                List.NonNullCount(
                    List.RemoveMatchingItems(
                        Record.FieldValues(_),
                        {"", null}
                    )
                ) > 0
        ),

    // ============================================================
    // 4. CLEAN ALL TEXT COLUMNS
    // ============================================================
    TextColumns =
        {
            "Order ID",
            "Ship Mode",
            "Customer ID",
            "Customer Name",
            "Segment",
            "Country/Region",
            "City",
            "State/Province",
            "Postal Code",
            "Region",
            "Product ID",
            "Category",
            "Sub-Category",
            "Product Name"
        },

    #"Cleaned Text" =
        Table.TransformColumns(
            #"Removed Blank Rows",
            List.Transform(
                TextColumns,
                each {
                    _,
                    (value) =>
                        if value = null
                        then null
                        else Text.Trim(Text.Clean(Text.From(value))),
                    type text
                }
            )
        ),

    // ============================================================
    // 5. SAFE DATE CONVERSION
    // ============================================================
    #"Normalized Dates" =
        Table.TransformColumns(
            #"Cleaned Text",
            {
                {
                    "Order Date",
                    each
                        try Date.FromText(
                            Text.Trim(Text.From(_)),
                            "en-US"
                        )
                        otherwise null,
                    type date
                },
                {
                    "Ship Date",
                    each
                        try Date.FromText(
                            Text.Trim(Text.From(_)),
                            "en-US"
                        )
                        otherwise null,
                    type date
                }
            }
        ),

    // ============================================================
    // 6. SAFE NUMERIC CONVERSION
    // ============================================================
    #"Normalized Numbers" =
        Table.TransformColumns(
            #"Normalized Dates",
            {
                {
                    "Row ID",
                    each try Int64.From(_) otherwise null,
                    Int64.Type
                },
                {
                    "Sales",
                    each try Number.From(_) otherwise null,
                    type number
                },
                {
                    "Quantity",
                    each try Int64.From(_) otherwise null,
                    Int64.Type
                },
                {
                    "Discount",
                    each try Number.From(_) otherwise null,
                    type number
                },
                {
                    "Profit",
                    each try Number.From(_) otherwise null,
                    type number
                }
            }
        ),

    // ============================================================
    // 7. ENSURE TEXT TYPES
    // ============================================================
    #"Changed Text Types" =
        Table.TransformColumnTypes(
            #"Normalized Numbers",
            {
                {"Order ID", type text},
                {"Ship Mode", type text},
                {"Customer ID", type text},
                {"Customer Name", type text},
                {"Segment", type text},
                {"Country/Region", type text},
                {"City", type text},
                {"State/Province", type text},
                {"Postal Code", type text},
                {"Region", type text},
                {"Product ID", type text},
                {"Category", type text},
                {"Sub-Category", type text},
                {"Product Name", type text}
            }
        ),

    // ============================================================
    // 8. ADD YEAR
    // ============================================================
    #"Added Year" =
        Table.AddColumn(
            #"Changed Text Types",
            "Year",
            each
                if [Order Date] = null
                then null
                else Date.Year([Order Date]),
            Int64.Type
        ),

    // ============================================================
    // 9. ADD MONTH NUMBER
    // ============================================================
    #"Added Month Number" =
        Table.AddColumn(
            #"Added Year",
            "Month Number",
            each
                if [Order Date] = null
                then null
                else Date.Month([Order Date]),
            Int64.Type
        ),

    // ============================================================
    // 10. ADD MONTH NAME
    // ============================================================
    #"Added Month Name" =
        Table.AddColumn(
            #"Added Month Number",
            "Month",
            each
                if [Order Date] = null
                then null
                else Date.MonthName([Order Date], "en-US"),
            type text
        ),

    // ============================================================
    // 11. ADD QUARTER NUMBER
    // ============================================================
    #"Added Quarter Number" =
        Table.AddColumn(
            #"Added Month Name",
            "Quarter Number",
            each
                if [Order Date] = null
                then null
                else Date.QuarterOfYear([Order Date]),
            Int64.Type
        ),

    // ============================================================
    // 12. ADD QUARTER NAME
    // ============================================================
    #"Added Quarter" =
        Table.AddColumn(
            #"Added Quarter Number",
            "Quarter",
            each
                if [Order Date] = null
                then null
                else "Q" & Text.From(Date.QuarterOfYear([Order Date])),
            type text
        ),

    // ============================================================
    // 13. FINAL COLUMN ORDER
    // ============================================================
    #"Reordered Columns" =
        Table.ReorderColumns(
            #"Added Quarter",
            {
                "Row ID",
                "Order ID",
                "Order Date",
                "Ship Date",
                "Year",
                "Month Number",
                "Month",
                "Quarter Number",
                "Quarter",
                "Ship Mode",
                "Customer ID",
                "Customer Name",
                "Segment",
                "Country/Region",
                "City",
                "State/Province",
                "Postal Code",
                "Region",
                "Product ID",
                "Category",
                "Sub-Category",
                "Product Name",
                "Sales",
                "Quantity",
                "Discount",
                "Profit"
            }
        )
in
    #"Reordered Columns"

🧪 Data Quality Checks

After applying the query, verify the following:

Date Validation

Check that:

Order Date is Date type

Ship Date is Date type

Invalid dates become null

Order Date is not later than Ship Date

Text Validation

Check that:

Category values do not contain unnecessary spaces

Customer names are consistently formatted

Product names do not contain leading/trailing spaces

Numeric Validation

Check that:

Sales is numeric

Quantity is a whole number

Discount is numeric

Profit is numeric

📊 Recommended Power BI Usage

The cleaned dataset can be used for:

Sales analysis

Profit analysis

Regional performance

Category analysis

Customer segmentation

Monthly sales trends

Quarterly sales trends

Year-over-year analysis

Shipping analysis

Recommended visualizations:

Analysis              Recommended Visual

Monthly Sales         Line Chart
Quarterly Sales       Column Chart
Category Sales        Bar Chart
Regional Profit       Map / Bar Chart
Segment Performance   Donut / Bar Chart
Sales vs Profit       Scatter Plot
Ship Mode Analysis    Column Chart

📌 Important Note for GitHub

The original query used a local path:

C:\Users\rajad\OneDrive\Desktop\EDA\samplesuperstore.csv

This path only works on the original computer.

For a GitHub project, use a relative path or configure the file path
through a Power BI parameter.

For example:

File.Contents("samplesuperstore.csv")

Depending on your Power BI environment, you may prefer a File Path
parameter for easier portability.

🚀 Project Objective

The objective of this transformation is to convert raw Superstore data
into a clean, consistent, analysis-ready dataset.

The final pipeline performs:

Raw CSV
   ↓
Promote Headers
   ↓
Remove Blank Rows
   ↓
Clean & Trim Text
   ↓
Normalize Dates
   ↓
Correct Numeric Types
   ↓
Handle Conversion Errors
   ↓
Create Year
   ↓
Create Month
   ↓
Create Quarter
   ↓
Reorder Columns
   ↓
Analysis-Ready Dataset
<img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/03f9268c-b5a4-4307-a50a-d7b97cb48a86" />

Author

K. Jeeva Saravanan
MCA | Computer Science
