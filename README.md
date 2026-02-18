# Excel Coffee Sales Dashboard Tutorial

## Overview
This tutorial demonstrates building an end-to-end Excel project: gathering data, transforming it, and creating a dynamic interactive coffee sales dashboard using pivot tables, pivot charts, slicers, and timelines.

## Final Dashboard Features
- **Line chart**: Total sales over time by coffee type (Arabica, Excelsa, Liberica, Robusta)
- **Bar chart**: Sales by country (US, Ireland, UK)
- **Bar chart**: Top 5 customers
- **Timeline**: Filter all visuals by date range
- **Slicers**: Filter by roast type (Dark, Light, Medium), size (0.2kg, 0.5kg, 1kg, 2.5kg), and loyalty card status

**Key Feature:** All visuals are interconnected - selecting any filter updates all charts simultaneously

> Example outputs of the completed tracker.

### Coffee Sales Dashboard (using Excel)

![Final Dashboard](imgs/Dashboard.png)

---

## PROJECT STRUCTURE

### Dataset Tables

#### Orders Table
**Pre-populated columns:**
- Order ID
- Order Date
- Customer ID
- Product ID
- Quantity

**Empty columns (to be filled via lookups):**
- Customer Name
- Email
- Country
- Coffee Type
- Roast Type
- Size
- Unit Price
- Sales
- Coffee Type Name (full name)
- Roast Type Name (full name)
- Loyalty Card

#### Customers Table
**Contains:**
- Customer ID (primary key)
- Customer Name
- Email
- Phone Number
- Address Line 1
- City
- Country
- Postcode
- Loyalty Card (Yes/No)

#### Products Table
**Contains:**
- Product ID (primary key)
- Coffee Type (abbreviated: Rob, Exc, Ara, Lib)
- Roast Type (abbreviated: M, L, D)
- Size (in kg)
- Unit Price
- Price per 100g
- Profit

---

## STEP 1: DATA GATHERING

### Using XLOOKUP for Customer Data

#### Customer Name Lookup

**Formula:**
```excel
=XLOOKUP(C2, Customers!$A$1:$A$1001, Customers!$B$1:$B$1001, "", 0)
```

**Breakdown:**
- `C2` = Lookup value (Customer ID)
- `Customers!$A$1:$A$1001` = Lookup array (Customer IDs in Customers table)
- `Customers!$B$1:$B$1001` = Return array (Customer Names)
- `""` = If not found, return blank
- `0` = Exact match

**Steps:**
1. Click in cell F2
2. Type `=XLOOKUP(`
3. Select C2 (Customer ID)
4. Comma
5. Go to Customers sheet
6. Click A1, then Ctrl+Shift+Down to select all Customer IDs
7. Press F4 to lock range (adds $ signs)
8. Comma
9. Press Ctrl+Home to return to A1
10. Click B1, Ctrl+Shift+Down to select all Customer Names
11. Press F4 to lock range
12. Comma, `""` for if not found
13. Comma, `0` for exact match
14. Close parenthesis, Enter
15. Double-click fill handle to copy down

**Locking Ranges:**
- Press F4 to toggle between reference types
- `$A$1` = Fully locked (column and row)
- `A$1` = Row locked only
- `$A1` = Column locked only

#### Email Lookup

**Formula:**
```excel
=XLOOKUP(C2, Customers!$A$1:$A$1001, Customers!$C$1:$C$1001, "", 0)
```

**Problem:** Missing emails return 0

**Solution - Add IF wrapper:**
```excel
=IF(XLOOKUP(C2, Customers!$A$1:$A$1001, Customers!$C$1:$C$1001, "", 0)=0, "", XLOOKUP(C2, Customers!$A$1:$A$1001, Customers!$C$1:$C$1001, "", 0))
```

**Logic:**
- IF XLOOKUP result = 0
- THEN return blank ""
- ELSE return XLOOKUP result

**Why needed:** Prevents zeros from displaying for missing emails

#### Country Lookup

**Formula:**
```excel
=XLOOKUP(C2, Customers!$A$1:$A$1001, Customers!$G$1:$G$1001, "", 0)
```

Same structure as Customer Name, but returns Country column (G)

**Selecting Full Column with Missing Data:**
1. Click into column start (e.g., C1)
2. Ctrl+Shift+Down (selects to first blank)
3. Ctrl+Shift+End (selects to actual end of data)
4. Hold Shift, press Left Arrow to deselect unwanted columns
5. Press F4 to lock

---

## STEP 2: USING INDEX MATCH FOR PRODUCT DATA

### Why INDEX MATCH Instead of XLOOKUP?

**Dynamic Formula:**
- One formula works for ALL product columns
- Drag right → formula adapts to column
- Drag down → formula copies to all rows

**XLOOKUP requires:**
- Separate formula for each column
- More repetition

### The INDEX MATCH Formula

**Complete formula in cell I2:**
```excel
=INDEX(Products!$A$1:$G$49, MATCH($D2, Products!$A$1:$A$49, 0), MATCH(I$1, Products!$A$1:$G$1, 0))

or 
=INDEX(productstbl[#All],MATCH(orders!$D2,products!$A:$A,0),MATCH(orders!I$1,products!$1:$1,0))
```

**Breakdown:**

#### INDEX Function
```excel
INDEX(array, row_num, column_num)
```

**Array:** `Products!$A$1:$G$49`
- Entire Products table data
- Locked with $ (doesn't change when copying)

#### MATCH for Row Number
```excel
MATCH($D2, Products!$A$1:$A$49, 0)
```

**Lookup value:** `$D2`
- Product ID from Orders table
- $ before D = Column locked (D doesn't change to E, F, G when dragging right)
- No $ before 2 = Row can change (2 becomes 3, 4, 5 when dragging down)

**Lookup array:** `Products!$A$1:$A$49`
- Product ID column in Products table
- Fully locked with $

**Match type:** `0` = Exact match

**Result:** Returns row number of matching Product ID

#### MATCH for Column Number
```excel
MATCH(I$1, Products!$A$1:$G$1, 0)
```

**Lookup value:** `I$1`
- Column header (Coffee Type, Roast Type, Size, Unit Price)
- No $ before I = Column can change (I becomes J, K, L when dragging right)
- $ before 1 = Row locked (always checks row 1 headers)

**Lookup array:** `Products!$A$1:$G$1`
- Header row in Products table
- Fully locked with $

**Match type:** `0` = Exact match

**Result:** Returns column number of matching header

### How Dynamic Behavior Works

**Original formula in I2:**
- Looks up Product ID from $D2
- Finds column named in I$1 (Coffee Type)
- Returns intersection value

**When copied to J2:**
- Still looks up Product ID from $D2 (column D locked)
- Now finds column named in J$1 (Roast Type)
- Returns new intersection value

**When copied to I3:**
- Looks up Product ID from $D3 (row changed)
- Finds column named in I$1 (row 1 locked)
- Returns correct product's coffee type

**Verify with F2 key:**
- Click cell, press F2
- Excel highlights referenced cells in color
- Confirms which cells formula uses

### Applying the Formula

1. Enter formula in I2
2. Press Enter
3. Drag right across columns J, K, L (Roast Type, Size, Unit Price)
4. All columns auto-populate correctly
5. Select I2:L2
6. Double-click fill handle
7. All product data populates down

---

## STEP 3: CALCULATING SALES

### Simple Multiplication Formula

**Formula in M2:**
```excel
=L2*E2
```

**Logic:**
- L2 = Unit Price
- E2 = Quantity
- Result = Total Sales for that order

**Apply:**
1. Enter formula in M2
2. Double-click fill handle
3. All sales values calculate

---

## STEP 4: CREATING FULL NAMES FOR CODES

### Coffee Type Name (Multiple Nested IFs)

**Problem:** Coffee Type column shows abbreviations (Rob, Exc, Ara, Lib)

**Solution:** Create full names using nested IF statements

**Formula in N2:**
```excel
=IF(I2="Rob", "Robusta", IF(I2="Exc", "Excelsa", IF(I2="Ara", "Arabica", IF(I2="Lib", "Liberica", ""))))
```

**Logic:**
- IF I2 = "Rob" THEN "Robusta"
- ELSE IF I2 = "Exc" THEN "Excelsa"
- ELSE IF I2 = "Ara" THEN "Arabica"
- ELSE IF I2 = "Lib" THEN "Liberica"
- ELSE "" (blank)

**Nested Structure:**
```
IF(condition1, result1,
  IF(condition2, result2,
    IF(condition3, result3,
      IF(condition4, result4, 
        default_result
      )
    )
  )
)
```

### Roast Type Name

**Formula in O2:**
```excel
=IF(J2="M", "Medium", IF(J2="L", "Light", IF(J2="D", "Dark", "")))
```

**Logic:**
- M → Medium
- L → Light
- D → Dark

---

## STEP 5: DATA FORMATTING

### Date Formatting (Custom Format)

**Problem:** Date formats vary by region
- European: Day-Month-Year (DD/MM/YYYY)
- American: Month-Day-Year (MM/DD/YYYY)

**Solution:** Use abbreviated month text

**Steps:**
1. Select Order Date column (B2, Ctrl+Shift+Down)
2. Press Ctrl+1 (Format Cells)
3. Category: Custom
4. Type: `dd-mmm-yyyy`
5. Click OK

**Result:** 05-Sep-2019 (universally clear format)

**Format Code Breakdown:**
- `dd` = Two-digit day
- `mmm` = Three-letter month abbreviation
- `yyyy` = Four-digit year

### Size Formatting (Adding Unit Label)

**Problem:** Size shows numbers without unit (0.5, 1, 2.5)

**Solution:** Add "kg" suffix with custom format

**Steps:**
1. Select Size column (K2, Ctrl+Shift+Down)
2. Ctrl+1 (Format Cells)
3. Category: Custom
4. Type: `0.0 "kg"`
5. Click OK

**Result:** 0.5 kg, 1.0 kg, 2.5 kg

**Format Code Breakdown:**
- `0.0` = Number with one decimal place
- `" kg"` = Text in quotes appears after number

### Currency Formatting

**Steps:**
1. Select Unit Price and Sales columns (L2:M2, Ctrl+Shift+Down)
2. Home > Number group > Currency dropdown
3. Select "$ English (United States)"

**Result:** $12.50, $25.00

---

## STEP 6: DATA QUALITY CHECKS

### Check for Duplicates

**Steps:**
1. Click in data (A1)
2. Ctrl+Shift+Right, Ctrl+Shift+Down (select all data)
3. Data tab > Remove Duplicates
4. Click OK
5. Excel reports: "No duplicate values found"

**Important:** Always verify data integrity before analysis

---

## STEP 7: CONVERT RANGE TO TABLE

### Why Convert to Table?

**Benefits:**
1. **Auto-expansion:** New columns/rows automatically included
2. **Structured references:** Formulas use table/column names
3. **Easy pivot refresh:** Pivot tables auto-update with new data
4. **Professional appearance:** Built-in formatting

### How to Convert

**Method 1 - Keyboard Shortcut:**
1. Click anywhere in data
2. Press Ctrl+T
3. Confirm range
4. Click OK

**Method 2 - Ribbon:**
1. Insert tab > Table
2. Confirm range
3. Click OK

**Verification:**
- Table Design tab appears
- Data has filter dropdowns
- Colored banding applies

### Table Setup

**Name the table:**
1. Table Design tab
2. Table Name box (top left)
3. Type: "Orders"

**Choose table style:**
1. Table Design tab
2. Table Styles gallery
3. Select lighter style for readability

---

## STEP 8: ADDING LOYALTY CARD COLUMN

### Demonstrating Table Auto-Expansion

**Scenario:** Need to add Loyalty Card data after table created

**Steps:**
1. Click in P1 (next to last column)
2. Type header: "Loyalty Card"
3. Press Enter
4. Column automatically becomes part of Orders table

**Populate with XLOOKUP:**
```excel
=XLOOKUP(C2, Customers!$A$1:$A$1001, Customers!$I$1:$I$1001, "", 0)
```

**Formula automatically fills down entire column** (table feature)

**Refresh Pivot Tables:**
1. Click any pivot table
2. PivotTable Analyze > Refresh
3. New column appears in field list

**Alternative Refresh:**
- Right-click pivot table > Refresh

**Why This Works:**
- Pivot table data source = "Orders" table
- Table auto-expanded to include column P
- Refresh brings in new column

**If Using Range Instead of Table:**
- Range would stop at column O
- Would need to manually update data source
- More error-prone and time-consuming

---

## STEP 9: CREATING PIVOT TABLES AND CHARTS

### Insert First Pivot Table

**Method 1 - Ribbon:**
1. Click in data
2. Insert > PivotTable
3. Table/Range: Auto-selects "Orders"
4. New Worksheet
5. Click OK

**Method 2 - Keyboard (Faster):**
1. Click in data
2. Press Alt+N+V+T
3. Press Enter

**Breakdown:**
- Alt = Activates ribbon shortcuts
- N = Insert tab
- V = PivotTable
- T = Table/Range option
- Enter = Confirm

**Practice:** Second method saves significant time

### Setting Up Total Sales Pivot Table

**Rename worksheet:** "Total Sales"

**Name pivot table:**
1. Click in pivot table
2. PivotTable Analyze tab
3. PivotTable Name box: "Total Sales"

#### Configure Fields

**Rows:**
1. Drag "Order Date" to Rows
2. Excel auto-groups by Years
3. Right-click on year > Group
4. Hold Ctrl, select both Years and Months
5. Click OK

**Result:** Hierarchical structure showing months within years

**Columns:**
1. Drag "Coffee Type Name" to Columns

**Values:**
1. Drag "Sales" to Values

#### Preferred Layout

**Steps:**
1. PivotTable Design (or Analyze) tab
2. Report Layout > Show in Tabular Form
3. Grand Totals > Off for Rows and Columns
4. Subtotals > Do Not Show Subtotals

**Why:**
- Cleaner appearance
- Better for charting
- More dashboard-friendly

#### Format Sales Values

**Steps:**
1. Click on Sales in pivot table
2. Value Field Settings
3. Number Format
4. Category: Number
5. Use 1000 Separator: Checked
6. Decimal Places: 0
7. Click OK twice

**Result:** 50000 displays as 50,000

---

## STEP 10: CREATING TOTAL SALES LINE CHART

### Insert Pivot Chart

**Steps:**
1. Click in pivot table
2. Insert > Line Chart > Line with Markers
3. Chart appears

### Remove Field Buttons

**Problem:** Default chart shows filter buttons

**Solution:**
1. Click on chart
2. PivotChart Analyze tab
3. Field Buttons > Hide All

**Result:** Clean chart without clutter

### Format Chart Background

**Steps:**
1. Double-click on chart area
2. Format pane opens
3. Fill > Solid Fill
4. Color: Custom RGB (60, 20, 100) - Purple
5. Adjust transparency for lighter shade

### Format Text Color

**Steps:**
1. Click on chart
2. Home tab > Font Color
3. More Colors > Custom
4. RGB: 60, 20, 100 (dark purple)

**Applies to:**
- Axis labels
- Chart title
- Legend

### Format Axis Lines

**Steps:**
1. Click on axis
2. Format Axis pane
3. Line > Solid Line
4. Color: White
5. Width: Increase slightly

### Add Axis Title

**Steps:**
1. Chart Design tab
2. Add Chart Element > Axis Titles
3. Primary Vertical
4. Type: "USD"

### Add Chart Title

**Steps:**
1. Chart Design > Add Chart Element
2. Chart Title > Above Chart
3. Type: "Total Sales Over Time"

### Customize Line Colors

**Purpose:** Differentiate coffee types visually

**Steps for each coffee type:**
1. Click on line to select series
2. Format Data Series pane
3. Line > Solid Line
4. Choose distinct color:
   - Liberica: Yellow
   - Excelsa: Dark brown
   - Arabica: Bright blue
   - Robusta: Red

**Result:** Each coffee type has unique, identifiable color

---

## STEP 11: CREATING AND FORMATTING TIMELINE

### Insert Timeline

**Steps:**
1. Click on pivot chart or pivot table
2. PivotChart Analyze (or PivotTable Analyze) tab
3. Insert Timeline
4. Excel auto-detects date field: Order Date
5. Check "Order Date"
6. Click OK

**Result:** Timeline filter appears

**Test functionality:**
- Drag across different time periods
- Chart updates to show only selected dates
- Clear filter to show all data

### Create Custom Timeline Style

**Problem:** Default style doesn't match dashboard theme

**Solution:** Create custom purple style

**Steps:**
1. Timeline tab > Timeline Styles
2. New Timeline Style
3. Name: "Purple Timeline Style"

#### Format "Whole Timeline"
- Font: Calibri Body, Size 11, White color
- Border: Dark purple (60, 20, 100), inside and outside
- Fill: Dark purple (60, 20, 100)

#### Format "Header"
- Font: Calibri Body, Bold, Size 11, White
- Border: White

#### Format "Selection Label"
- Font: Calibri Body, Bold, White

#### Format "Time Level"
- Font: Calibri Body, Bold, White

#### Format "Period Labels"
- Font: Calibri Body, Bold, White

#### Format "Selected Time Block"
- Border: White
- Fill: Bright purple (lighter than background)

#### Format "Unselected Time Block"
- Fill: Very light gray
- Creates clear visual distinction

**Apply style:**
1. Click on timeline
2. Select "Purple Timeline Style"

**Result:** Timeline matches dashboard color scheme

---

## STEP 12: CREATING AND FORMATTING SLICERS

### Insert Slicers

**Initial slicers needed:**
- Roast Type Name
- Size

**Steps:**
1. Click pivot chart or table
2. PivotChart Analyze > Insert Slicer
3. Check: Roast Type Name, Size
4. Click OK

**Note:** Loyalty Card not available yet (will add after refresh)

### Create Custom Slicer Style

**Steps:**
1. Slicer tab > Slicer Styles
2. New Slicer Style
3. Name: "Purple Slicer"

#### Format "Whole Slicer"
- Fill: Dark purple (60, 20, 100)
- Font: Calibri Body, Regular, White

#### Format "Header"
- Font: Calibri Body, Bold, White

#### Format "Selected Item with Data"
- Font: Calibri Body, Bold, White
- Border: White, All sides

#### Format "Selected Item without Data"
- Font: Calibri Body, Bold, White
- Border: White, All sides

#### Format "Unselected Item with Data"
- Font: Calibri Body, Regular
- Color: Light gray
- Border: White
- Text: Strikethrough

#### Format "Unselected Item without Data"
- Font: Calibri Body, Regular
- Color: Light gray
- Border: White
- Text: Strikethrough

**Set as default:**
1. Right-click on custom style
2. "Set as Default"

**Apply to slicers:**
1. Click each slicer
2. Select "Purple Slicer" style

### Adjust Slicer Layout

**Roast Type Name - 3 columns:**
1. Right-click slicer
2. Size and Properties
3. Number of columns: 3
4. Resize slicer to show all in one row

**Size - 2 columns:**
1. Right-click slicer
2. Size and Properties
3. Number of columns: 2

### Add Loyalty Card Slicer

**After refreshing pivot table with new column:**

**Steps:**
1. Click pivot chart
2. PivotChart Analyze > Insert Slicer
3. Check: Loyalty Card
4. Click OK
5. Purple Slicer style auto-applies (if set as default)

**Test all slicers:**
- Dark/Light/Medium roasts
- Different sizes
- Yes/No loyalty card
- All should filter the chart

---

## STEP 13: CREATING SALES BY COUNTRY CHART

### Duplicate Worksheet

**Why duplicate instead of new pivot:**
- Maintains connection to existing slicers/timeline
- Easier to connect later
- Consistent data source

**Method:**
1. Hold Ctrl
2. Click and drag "Total Sales" tab
3. Release when you see small + icon
4. Duplicate appears

**Rename:** "Country Bar Chart"

### Delete Unnecessary Visuals

**Keep:** Pivot table only

**Delete:**
- Line chart
- Timeline
- All slicers

### Reconfigure Pivot Table

**Remove existing fields:**
- Order Date (from Rows)
- Coffee Type Name (from Columns)

**Add new fields:**

**Rows:**
1. Drag "Country" to Rows

**Values:**
1. "Sales" already in Values (from duplication)

**Result:** Simple table showing total sales per country

### Insert Bar Chart

**Steps:**
1. Click in pivot table
2. Insert > Bar Chart > Clustered Bar
3. Chart appears

### Sort Bars by Value

**Goal:** Highest sales country at top

**Steps:**
1. Right-click on "Country" in chart
2. Sort > More Sort Options
3. Sort by: Sum of Sales
4. Order: Ascending (for bar charts, this puts highest at top)
5. Click OK

**Result:** US (highest) at top, UK (lowest) at bottom

### Remove Field Buttons

**Steps:**
1. Click chart
2. PivotChart Analyze > Field Buttons > Hide All

### Remove Legend

**Steps:**
1. Click on legend
2. Press Delete

(Legend not needed - country names on axis)

### Add Chart Title

**Steps:**
1. Click chart title
2. Type: "Sales by Country"

### Format Chart

**Chart area:**
1. Double-click chart background
2. Fill > Solid Fill
3. Color: Purple (matching theme)
4. Transparency: Adjust for lighter shade

**Axis:**
1. Click axis
2. Format > Line > Solid Line
3. Color: White
4. Width: Slightly thicker

**Font colors:**
1. Select chart
2. Home > Font Color > Dark Purple

### Customize Bar Colors

**Give each country unique color:**

**Initial setup:**
1. Click on bars (all selected)
2. Fill > Solid Fill
3. Color: Base green
4. Border: Solid line, White, 2 points width

**Individual colors:**
1. Click on bar twice (selects single bar)
2. Fill > More Colors
3. Choose color:
   - US: Darker green
   - Ireland: Medium green  
   - UK: Lighter green

**Result:** Visual differentiation between countries

### Add Data Labels

**Steps:**
1. Chart Design > Add Chart Element
2. Data Labels > Outside End
3. Labels appear at end of bars

### Format Data Labels as Currency

**Steps:**
1. Click in pivot table
2. Right-click on "Sum of Sales" in Values
3. Value Field Settings
4. Number Format
5. Category: Currency
6. Symbol: $ English (United States)
7. Click OK twice

**Result:** Labels show as $79,566

---

## STEP 14: CREATING TOP 5 CUSTOMERS CHART

### Duplicate Country Chart Worksheet

**Steps:**
1. Hold Ctrl
2. Drag "Country Bar Chart" tab
3. Release at +
4. Duplicate appears

**Rename:** "Top 5 Customers"

### Modify Pivot Table

**Remove:** Country from Rows

**Add:** Customer Name to Rows

**Problem:** Shows ALL customers (100+)

### Filter to Top 5

**Steps:**
1. Click dropdown on "Customer Name" in pivot
2. Value Filters > Top 10
3. Change "10" to "5"
4. Filter by: Sum of Sales
5. Click OK

**Result:** Only 5 customers showing (highest sales)

### Sort Top 5

**Steps:**
1. Right-click "Customer Name" in chart
2. Sort > More Sort Options
3. Ascending by Sum of Sales
4. Click OK

**Result:** Highest customer at top

### Update Chart Title

**Steps:**
1. Click chart title
2. Type: "Top 5 Customers"

**Result:** Bar chart showing top 5 customers by sales value

---

## STEP 15: BUILDING THE DASHBOARD

### Create Dashboard Worksheet

**Steps:**
1. Click + at bottom (new sheet)
2. Rename: "Dashboard"

### Set Up Grid

**Column A:**
1. Select column A
2. Set width: 1
3. Creates narrow left margin

**Row 1:**
1. Select row 1  
2. Set height: 5
3. Creates narrow top margin

### Create Title Bar

**Steps:**
1. Insert > Shapes > Rectangle
2. Hold Alt while dragging (snaps to cell borders)
3. Drag from column B to column Z
4. Covers top of dashboard

**Format title bar:**
1. Fill: Dark purple
2. Outline: Dark purple (or no outline)
3. Click inside shape
4. Type: "Coffee Sales Dashboard"
5. Font: White, Large size
6. Align: Center, Middle

**Tip:** Holding Alt makes shapes snap to cell borders precisely

### Move Visuals to Dashboard

**From Total Sales worksheet:**
1. Hold Ctrl
2. Select: Timeline, Line Chart, 3 Slicers
3. Ctrl+X (cut)
4. Click on Dashboard
5. Ctrl+V (paste)

**From Country Bar Chart:**
1. Select bar chart
2. Ctrl+X
3. Dashboard tab
4. Ctrl+V

**From Top 5 Customers:**
1. Select bar chart
2. Ctrl+X
3. Dashboard tab
4. Ctrl+V

### Arrange Dashboard Elements

**Layout (using Alt for snapping):**

**Top:**
- Title bar spanning full width

**Below title:**
- Timeline spanning most of width

**Left side:**
- Total Sales Over Time line chart (large)
- Sales by Country bar chart (medium)
- Top 5 Customers bar chart (medium)

**Right side:**
- Roast Type Name slicer
- Size slicer
- Loyalty Card slicer

**Tips:**
- Hold Alt while moving to snap to cells
- Resize while holding Alt for perfect alignment
- Leave small margins between elements
- Balance visual weight across dashboard

---

## STEP 16: CONNECTING FILTERS

### Why Connect?

**Current problem:**
- Timeline only filters Total Sales chart
- Slicers only filter Total Sales chart
- Country and Top 5 charts don't update

**Goal:** All filters affect all visuals

### Connect Timeline

**Steps:**
1. Click on timeline
2. Timeline tab > Report Connections
3. Check boxes for:
   - Total Sales (already checked)
   - Country Bar Chart
   - Top 5 Customers
4. Click OK

**Test:**
- Drag timeline to select 2020-2021
- All three charts update
- Clear filter

### Connect Roast Type Name Slicer

**Steps:**
1. Click slicer
2. Slicer tab > Report Connections
3. Check:
   - Country Bar Chart
   - Top 5 Customers
4. Click OK

**Test:**
- Click "Dark"
- All charts show only dark roast sales

### Connect Size Slicer

**Steps:**
1. Click slicer
2. Slicer tab > Report Connections
3. Check both additional charts
4. Click OK

**Test:**
- Click "1.0 kg"
- All charts filter to 1kg packages

### Connect Loyalty Card Slicer

**Steps:**
1. Click slicer
2. Slicer tab > Report Connections
3. Check both additional charts
4. Click OK

**Test:**
- Click "Yes"
- All charts show only loyalty card customers

### Verify All Connections

**Full test:**
1. Select 2020 on timeline
2. Select "Light" roast
3. Select "0.5 kg" size
4. Click "No" for loyalty card
5. All charts should update simultaneously

**Clear all filters and verify return to full data**

---

## STEP 17: FINAL DASHBOARD POLISH

### Remove Gridlines

**Steps:**
1. View tab
2. Uncheck "Gridlines"

**Result:** Clean background, professional appearance

### Hide Excel Interface Elements

**For presentation mode:**

**Steps:**
1. File > Options
2. Advanced
3. Display section

**Uncheck:**
- Show formula bar
- Show horizontal scroll bar
- Show vertical scroll bar
- Show sheet tabs
- Show row and column headers

4. Click OK

**Result:**
- Only dashboard visible
- No Excel navigation elements
- Users can only interact with slicers/timeline

**Pro tip:** Double-click on ribbon to auto-hide

### Optional: Keep Scroll Bars

**Why:**
- Users with small screens may need to scroll
- Horizontal scroll especially important
- Mouse wheel won't scroll horizontally without scrollbar

**Recommendation:**
- Keep horizontal and vertical scroll bars enabled
- Allows flexibility for different screen sizes

**Re-enable if needed:**
1. File > Options > Advanced
2. Display section
3. Check: Show horizontal scroll bar
4. Check: Show vertical scroll bar
5. Click OK

---

## COMPLETE FORMULA REFERENCE

### XLOOKUP Formulas

**Customer Name:**
```excel
=XLOOKUP(C2, Customers!$A$1:$A$1001, Customers!$B$1:$B$1001, "", 0)
```

**Email (with IF to handle zeros):**
```excel
=IF(XLOOKUP(C2, Customers!$A$1:$A$1001, Customers!$C$1:$C$1001, "", 0)=0, "", XLOOKUP(C2, Customers!$A$1:$A$1001, Customers!$C$1:$C$1001, "", 0))
```

**Country:**
```excel
=XLOOKUP(C2, Customers!$A$1:$A$1001, Customers!$G$1:$G$1001, "", 0)
```

**Loyalty Card:**
```excel
=XLOOKUP(C2, Customers!$A$1:$A$1001, Customers!$I$1:$I$1001, "", 0)
```

### INDEX MATCH Formula

**Dynamic product data lookup:**
```excel
=INDEX(Products!$A$1:$G$49, MATCH($D2, Products!$A$1:$A$49, 0), MATCH(I$1, Products!$A$1:$G$1, 0))
```

**Component breakdown:**
- Array: `Products!$A$1:$G$49` (entire products table)
- Row: `MATCH($D2, Products!$A$1:$A$49, 0)` (matches Product ID)
- Column: `MATCH(I$1, Products!$A$1:$G$1, 0)` (matches column header)

### Simple Formulas

**Sales calculation:**
```excel
=L2*E2
```
(Unit Price × Quantity)

**Coffee Type Name (Nested IFs):**
```excel
=IF(I2="Rob", "Robusta", IF(I2="Exc", "Excelsa", IF(I2="Ara", "Arabica", IF(I2="Lib", "Liberica", ""))))
```

**Roast Type Name:**
```excel
=IF(J2="M", "Medium", IF(J2="L", "Light", IF(J2="D", "Dark", "")))
```

---

## CUSTOM NUMBER FORMATS

**Date format:**
```
dd-mmm-yyyy
```
Result: 05-Sep-2019

**Size with unit:**
```
0.0 "kg"
```
Result: 1.0 kg

**Currency:**
- Category: Currency
- Symbol: $ English (United States)
- Decimals: 2
Result: $12.50

---

## KEYBOARD SHORTCUTS USED

**General:**
- **Ctrl+T**: Convert range to table
- **Ctrl+1**: Format Cells dialog
- **F4**: Toggle absolute/relative references ($)
- **F2**: Edit cell / Show formula references
- **Ctrl+Home**: Go to cell A1
- **Ctrl+Shift+Down**: Select from current cell to last filled cell in column
- **Ctrl+Shift+Right**: Select from current cell to last filled cell in row
- **Ctrl+Shift+End**: Select to end of used range
- **Ctrl+C**: Copy
- **Ctrl+V**: Paste
- **Ctrl+X**: Cut
- **Ctrl+Y**: Redo
- **Ctrl+Z**: Undo

**Pivot Table:**
- **Alt+N+V+T, Enter**: Insert PivotTable (fast method)

**Worksheet:**
- **Ctrl+Page Down**: Next worksheet
- **Ctrl+Page Up**: Previous worksheet

**Dashboard:**
- **Alt**: Hold while resizing/moving shapes to snap to cell borders

---

## PROJECT BEST PRACTICES DEMONSTRATED

### Data Management
1. **Use tables** for auto-expansion and structured references
2. **Name tables meaningfully** (Orders, Customers, Products)
3. **Lock cell references** appropriately ($D2 vs $D$2 vs D$2)
4. **Check for duplicates** before analysis
5. **Use consistent formatting** (dates, currency, units)

### Formula Strategy
1. **XLOOKUP for simple lookups** (newer, cleaner syntax)
2. **INDEX MATCH for dynamic lookups** (one formula for multiple columns)
3. **IF statements for data transformation** (abbreviations to full names)
4. **Error handling** (IF wrapper to prevent zeros)

### Pivot Table Design
1. **Tabular layout** for charts
2. **Remove subtotals/grand totals** for clean appearance
3. **Name pivot tables** for easier reference
4. **Group dates** meaningfully (years and months)

### Visualization
1. **Consistent color scheme** throughout (purple theme)
2. **Remove chart clutter** (field buttons, unnecessary legends)
3. **Add meaningful titles** and axis labels
4. **Differentiate data series** with distinct colors
5. **Format numbers** for readability (currency, thousands separator)

### Dashboard Design
1. **Create dedicated dashboard sheet** (separate from data)
2. **Use shapes for title** (professional header)
3. **Snap elements to grid** (hold Alt for precision)
4. **Balance visual weight** (large main chart, smaller supporting charts)
5. **Group related filters** (slicers together)

### User Experience
1. **Connect all filters** to all charts (consistent filtering)
2. **Remove gridlines** for clean look
3. **Hide interface elements** for presentation mode
4. **Keep scroll bars** for accessibility
5. **Test all interactions** before sharing

---

## TROUBLESHOOTING COMMON ISSUES

### XLOOKUP Returns 0 for Blank Cells
**Solution:** Wrap in IF statement
```excel
=IF(XLOOKUP(...)=0, "", XLOOKUP(...))
```

### INDEX MATCH Not Working When Copied
**Issue:** Dollar signs in wrong places

**Check:**
- Row match: `$D2` (column locked, row free)
- Column match: `I$1` (column free, row locked)

### Pivot Table Not Showing New Columns
**Solution:** Refresh the pivot table
1. Right-click pivot > Refresh
2. Or: PivotTable Analyze > Refresh

### Timeline/Slicer Not Filtering All Charts
**Solution:** Check report connections
1. Click timeline/slicer
2. Timeline/Slicer tab > Report Connections
3. Ensure all charts checked

### Chart Shows Wrong Sort Order
**For bar charts:** Use Ascending to put highest at top
**For column charts:** Use Descending to put highest at left

### Dates Not Grouping in Pivot
**Solution:** Ensure dates are actual date format (not text)
1. Select dates in source data
2. Ctrl+1 > Date format

---

## CUSTOMIZATION IDEAS

### Extend the Dashboard
1. **Add more metrics:**
   - Average order value
   - Profit margins
   - Customer retention rate

2. **Add more visualizations:**
   - Sales by roast type (pie chart)
   - Monthly growth rate (line chart)
   - Regional comparison (map chart)

3. **Add more slicers:**
   - Year filter
   - Month filter
   - Customer segment

4. **Enhance interactivity:**
   - Buttons to clear all filters
   - Toggle between different views
   - Drill-down capabilities

### Automate Data Updates
1. **Use Power Query** to refresh data from external sources
2. **Create refresh macro** (VBA) for one-click update
3. **Schedule automatic refreshes** (if using Excel Online/SharePoint)

### Professional Enhancements
1. **Add company branding** (logo, colors)
2. **Include data quality indicators** (last refresh date, data completeness)
3. **Add instructions** (text box with usage guide)
4. **Create multiple dashboard views** (executive summary vs detailed analysis)

---

*This coffee sales dashboard demonstrates end-to-end Excel project skills: data gathering with lookups, transformation with formulas, analysis with pivot tables, and professional visualization with an interactive dashboard.*
