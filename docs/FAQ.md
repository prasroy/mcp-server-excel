# Frequently Asked Questions - ExcelMcp

## What type of Excel work can I do with ExcelMcp?

ExcelMcp gives AI assistants (GitHub Copilot, Claude, ChatGPT) and scripts full control over Microsoft Excel through natural language or CLI commands. Here is a summary of what you can do, grouped by task type:

---

### 📊 Data Entry & Manipulation
- Read and write cell values, ranges, rows, and columns
- Copy, move, insert, or delete rows/columns/cells
- Find and replace values across a workbook
- Sort data by one or multiple columns

**Example prompt:** *"Put this table in A1:C4 - Name, Age, City / Alice, 30, Seattle / Bob, 25, Portland"*

---

### 📐 Formulas & Calculations
- Set or retrieve Excel formulas (e.g., `=SUM(A1:A10)`)
- Control calculation mode (Automatic, Manual, Semi-Automatic)
- Trigger a full workbook, sheet, or range recalculation

**Example prompt:** *"Add a formula column in D that multiplies Quantity by Unit Price"*

---

### 🎨 Formatting & Styling
- Apply font, color, borders, alignment, and cell orientation
- Apply built-in Excel cell styles
- Set number formats (currency, percentage, dates, custom)
- Auto-fit column widths and row heights
- Merge or unmerge cells

**Example prompt:** *"Make the header row bold with a dark background and auto-fit all column widths"*

---

### 🗂️ Worksheets & Workbooks
- Create, rename, copy, move, show/hide, or delete worksheets
- Set worksheet tab colors
- Open and close workbooks (including IRM/AIP-protected files)
- Create new `.xlsx` or `.xlsm` workbooks

**Example prompt:** *"Create a sheet called 'Summary' and copy the data from 'RawData' into it"*

---

### 📋 Excel Tables (ListObjects)
- Convert a range to an Excel Table with a chosen style
- Add/remove/rename columns, append rows, resize tables
- Apply filters (criteria or value-based) and multi-level sorts
- Add a Totals row with aggregate functions
- Load a table to the Power Pivot Data Model

**Example prompt:** *"Convert A1:D50 to a Table with a blue style and add a Totals row"*

---

### 🔄 Power Query & Data Import
- Create, update, rename, refresh, delete Power Query queries
- Write and manage M code with automatic formatting
- Set load destinations: worksheet, Data Model, both, or connection-only
- Evaluate M code ad-hoc without saving a permanent query

**Example prompt:** *"Import products.csv using Power Query and load it to the Data Model"*

---

### 📊 PivotTables & Analysis
- Create PivotTables from a range, Excel Table, or Data Model
- Add/remove row, column, value, and filter fields
- Set aggregation functions (Sum, Average, Count, Min, Max, etc.)
- Add calculated fields or OLAP calculated members
- Sort and filter PivotTable fields
- Create slicers for interactive filtering

**Example prompt:** *"Create a PivotTable showing total sales by Product and Region, then add a Region slicer"*

---

### 📉 Charts & Visualizations
- Create charts from a data range or PivotTable
- Set chart type, title, axis titles, legend, and style
- Add/remove/configure data series
- Configure data labels, axis scales, and gridlines
- Add trendlines (Linear, Exponential, Moving Average, etc.)
- Position and size charts on a worksheet

**Example prompt:** *"Create a bar chart from the Sales table and place it to the right of the data"*

---

### 🧮 Data Model & DAX (Power Pivot)
- Create, update, or delete DAX measures with automatic formatting
- Manage table relationships (create, update active/inactive, delete)
- List tables, columns, and measures in the Data Model
- Evaluate DAX EVALUATE queries for ad-hoc analysis
- Run DMV queries for metadata discovery

**Example prompt:** *"Create a measure called 'Total Revenue' as SUM(Sales[Amount]) in the Sales table"*

---

### 🔌 Data Connections
- Create, test, refresh, or delete OLEDB/ODBC connections
- Update connection strings and command text
- Load connection data to a worksheet

---

### 🏷️ Named Ranges & Parameters
- Create, read, write, update, or delete named ranges
- Ideal for parameter-driven workbooks: update a named cell → Power Query auto-refreshes

**Example prompt:** *"Set the named range 'ReportYear' to 2025"*

---

### 📝 VBA Macros
- List VBA components and procedures in a workbook
- Import new VBA modules from code or a file
- Update existing VBA code
- Run procedures with optional parameters
- Export modules for version control

**Example prompt:** *"Run the macro 'UpdatePrices' in the workbook"*

---

### 🎚️ Slicers & Conditional Formatting
- Add interactive slicers to PivotTables and Excel Tables
- Apply conditional formatting rules (cell value, formulas, color scales, data bars, icon sets)
- Clear conditional formatting from ranges

**Example prompt:** *"Highlight all values over $500 in the Revenue column in green"*

---

### 📸 Screenshots & Visual Verification
- Capture a specific cell range as a PNG image
- Capture an entire worksheet as a PNG image
- Images are returned directly to the AI for visual verification

**Example prompt:** *"Take a screenshot of the chart on Sheet1 so I can review it"*

---

### 🪧 Window Management
- Show or hide the Excel window
- Arrange the window (left-half, right-half, full-screen, center, etc.)
- Display live progress messages in Excel's status bar

**Example prompt:** *"Show me Excel side-by-side while you build this dashboard"*

---

## Quick Capability Reference

| Category | # of Operations |
|----------|----------------|
| File Operations | 6 |
| Worksheets | 16 |
| Ranges (values, formulas, formatting) | 46 |
| Excel Tables | 27 |
| PivotTables | 30 |
| Charts | 29 |
| Power Query | 12 |
| Data Model / DAX | 19 |
| Connections | 9 |
| Named Ranges | 6 |
| VBA Macros | 6 |
| Slicers | 8 |
| Conditional Formatting | 2 |
| Screenshot | 2 |
| Calculation Mode | 3 |
| Window Management | 9 |
| **Total** | **230** |

📚 **[Full Feature Reference →](../FEATURES.md)** - Detailed documentation of all 230 operations

---

## What ExcelMcp Does NOT Support

- ❌ **Server-side / headless Excel** — Requires a real Windows machine with Excel installed
- ❌ **Linux or macOS** — COM interop is Windows-only
- ❌ **High-volume batch processing** — Consider ClosedXML or EPPlus for that use case
- ❌ **DAX calculated columns** — Use Excel's UI instead

---

## More Resources

- **[README](../README.md)** — Overview and quick start
- **[Installation Guide](INSTALLATION.md)** — Setup for all AI assistants
- **[FEATURES.md](../FEATURES.md)** — Complete list of all 230 operations
- **[MCP Server Guide](../src/ExcelMcp.McpServer/README.md)** — Tool documentation
- **[CLI Guide](../src/ExcelMcp.CLI/README.md)** — Command-line reference
