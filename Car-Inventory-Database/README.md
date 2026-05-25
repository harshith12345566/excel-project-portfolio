# 🚗 Car Inventory Management & Analysis System

This project is a beginner-to-intermediate Excel-based car inventory management and vehicle analysis system created using Microsoft Excel.

The project simulates a real-world vehicle inventory database where different car brands, models, drivers, warranty coverage, and mileage information are managed and analyzed using Excel formulas, lookup operations, and business logic.

The dataset was imported from a raw text file and transformed into a structured business inventory spreadsheet for analysis and reporting.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 📌 Project Overview

The project manages inventory data for multiple automobile manufacturers including:

• Ford
• Toyota
• Honda
• Hyundai
• Chrysler
• General Motors

The spreadsheet stores detailed information about vehicles such as:

• Car ID
• Manufacturer Codes
• Manufacturer Names
• Vehicle Model Codes
• Vehicle Models
• Manufacturing Year
• Vehicle Age
• Mileage Information
• Miles Per Year
• Driver Assignment
• Warranty Miles
• Warranty Coverage Status
• New Car IDs

The project also includes lookup operations and data transformation logic for converting abbreviated inventory codes into readable business information.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 📂 Data Import Process

The project involved importing external vehicle inventory data from a text file into Microsoft Excel.

The raw `.txt` dataset was cleaned, structured, and transformed into a usable inventory management spreadsheet.

Operations performed after importing the text file included:

• Organizing raw inventory data
• Converting manufacturer codes into readable names
• Mapping vehicle model codes
• Formatting imported data into tabular structure
• Cleaning inconsistent values
• Performing lookup operations
• Generating calculated fields
• Analyzing warranty coverage and mileage data

This project helped me understand how Excel can be used for:

* data importing
* data cleaning
* business data transformation
* inventory analysis workflows

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# ✨ Features

• Vehicle inventory management
• Manufacturer and model code conversion
• Mileage tracking
• Warranty coverage analysis
• Driver assignment tracking
• Vehicle age analysis
• Miles-per-year calculations
• Lookup-based data mapping
• Business-style inventory dataset
• Spreadsheet automation using formulas

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 🧠 Excel Concepts Practiced

• IF Formula
• Lookup Operations
• Relative Referencing
• Absolute Referencing
• Text File Import
• Data Cleaning
• Data Transformation
• Spreadsheet Formatting
• Inventory Analysis
• Business Reporting

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 🔗 Lookup Operations Performed

The spreadsheet uses lookup-style operations to convert abbreviated codes into meaningful business values.

Examples include:

Manufacturer Code Conversion:

• FD → Ford
• GM → General Motors
• TY → Toyota
• HO → Honda
• HY → Hyundai
• CR → Chrysler

Vehicle Model Code Conversion:

• MTG → Mustang
• FCS → Focus
• CAM → Camry
• CIV → Civic
• ELA → Elantra
• PTC → PT Cruiser
• SLV → Silverado

This helped transform raw inventory codes into readable vehicle information.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 📊 Operations & Calculations Performed

## Vehicle Age Calculation

Vehicle manufacturing years were analyzed to determine vehicle age.

Example Concept:

Current Year - Manufacture Year

Used for:
• Inventory aging analysis
• Vehicle lifecycle tracking

---

## Miles Per Year Calculation

Mileage efficiency was analyzed by calculating miles driven per year.

Example Concept:

Miles / Vehicle Age

Used for:
• Vehicle usage analysis
• Driver usage comparison
• Wear and tear estimation

---

## Warranty Coverage Analysis

Warranty eligibility was analyzed based on mileage limits.

Example Logic:

IF(Miles <= Warranty_Miles, "Covered", "Not Covered")

Used to determine:
• Vehicles still under warranty
• Warranty expiration analysis

---

## New Car ID Transformation

New vehicle IDs were generated using:
• Manufacturer codes
• Model codes
• Vehicle numbers

This helped standardize inventory identifiers.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 🔗 Relative vs Absolute Referencing

## Relative Referencing

Relative references automatically adjust when formulas are copied across rows.

Example:

Miles / Age

Used to calculate mileage analysis for multiple vehicles automatically.

---

## Absolute Referencing

Absolute references keep specific cells fixed during formula copying.

Example:

=$F$2 / G2

The `$` symbol prevents Excel from changing the referenced cell during calculations.

Used for:
• Lookup tables
• Standardized calculations
• Fixed business references

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 📊 Formulas & Excel Operations Used

## IF Formula

Used for warranty coverage analysis.

Example:

=IF(Miles <= Warranty_Miles, "Covered", "Not Covered")

---

## Lookup Operations

Used to map:
• Manufacturer codes
• Vehicle model codes

into readable business names.

---

## Division Operations

Used for:

Miles Per Year calculations

Example:

=Miles / Age

---

## Concatenation / ID Formatting

Used to generate:
• Standardized Car IDs
• Vehicle identifiers

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# ⌨️ Excel Shortcuts Used

Ctrl + C → Copy
Ctrl + V → Paste
Ctrl + Z → Undo
Ctrl + S → Save Workbook
Ctrl + D → Fill Down
Ctrl + R → Fill Right
F2 → Edit Selected Cell
Ctrl + Arrow Keys → Fast Navigation
Ctrl + Shift + L → Apply Filters
Ctrl + T → Convert Range to Table
Alt + = → AutoSum

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 📈 Analytics Included

• Vehicle age analysis
• Mileage analysis
• Miles-per-year analysis
• Warranty coverage tracking
• Manufacturer inventory comparison
• Driver vehicle assignment analysis
• Vehicle model categorization

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 🎯 What I Learned

Through this project, I learned:

• How inventory databases are structured
• How to import external text datasets into Excel
• How raw business data is cleaned and structured
• How lookup operations transform coded inventory data
• How IF formulas automate business logic
• How vehicle analytics can be performed using spreadsheets
• How Excel can simulate real-world inventory management systems

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 🚀 Future Improvements

• Interactive inventory dashboard
• Vehicle maintenance tracking
• Dynamic filtering system
• Pivot table analytics
• Inventory visualization charts
• Automated reporting system

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 🛠️ Tools Used

• Microsoft Excel

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 📷 Project Screenshots

• Vehicle Inventory Dataset
• Lookup Operations
• Warranty Analysis
• Mileage Calculations

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 📚 Learning Journey

This project is part of my Excel and Data Analytics learning journey where I’m practicing:

• Spreadsheet Automation
• Inventory Management
• Data Cleaning
• Lookup Operations
• Business Reporting
• Data Analysis

through practical hands-on projects and business simulations.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 🌟 About This Project

This project focuses on learning practical inventory management and business data analysis using Excel by simulating a real-world vehicle inventory system.

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# 🙌 Feedback

Suggestions, improvements, and feedback are always welcome!

