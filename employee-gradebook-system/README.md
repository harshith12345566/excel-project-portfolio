````markdown id="m7x2qp"
# 📘 Employee Gradebook System

This project is a beginner-level Excel gradebook system created to practice spreadsheet automation, performance analysis, and data visualization using Microsoft Excel.

The idea behind this project was to simulate how a company or organization might evaluate employee performance through different tests and assessments.

---

# 📌 Project Overview

In this project, employees are evaluated using multiple tests such as:

- Safety Test
- Company Philosophy Test
- Financial Skills Test
- Drug Test

Using Excel formulas and formatting tools, the spreadsheet automatically:
- calculates percentages
- analyzes performance
- identifies pass/fail conditions
- highlights employee performance visually
- generates summary analytics

This project helped me understand how Excel can be used beyond simple calculations and applied to real-world business reporting.

---

# ✨ Features

- Employee test score management
- Automatic percentage calculations
- Pass/Fail analysis using formulas
- Conditional formatting for performance visualization
- Maximum, Minimum, and Average score calculations
- Employee performance charts
- Structured spreadsheet formatting
- Business-style reporting layout

---

# 🧠 Excel Concepts Practiced

| Concept | Purpose |
|---|---|
| IF Formula | Pass/Fail conditions |
| OR Formula | Employee evaluation logic |
| Percentage Calculations | Performance analysis |
| Conditional Formatting | Highlighting scores visually |
| MAX / MIN / AVERAGE | Summary analytics |
| Relative References | Formula automation |
| Absolute References | Fixed value calculations |
| Charts | Data visualization |
| Spreadsheet Formatting | Improving readability |

---

# 🔗 Relative vs Absolute Referencing

## Relative Referencing

Relative references automatically adjust when formulas are copied across rows or columns.

### Example

```excel
=Employee_Score / Total_Marks
```

This was useful for calculating percentages for multiple employees quickly.

---

## Absolute Referencing

Absolute references keep a specific row or column fixed during formula copying.

### Example

```excel
=C6 / C$2
```

The `$` symbol prevents Excel from changing the reference while copying formulas.

This was especially useful when dividing scores by a fixed total marks value.

---

# 📊 Formulas Used

## Percentage Formula

```excel
=Employee_Score / Total_Marks
```

Used to calculate employee test percentages.

---

## IF Formula

```excel
=IF(Test_Percentage >= 50%, "Pass", "Fail")
```

Used to determine whether an employee passed or failed.

---

## OR Formula

```excel
=OR(Test1 < 50%, Test2 < 50%, Test3 < 50%, Test4 < 50%)
```

Used to identify employees who failed any test.

---

## MAX Formula

```excel
=MAX(Test_Score_Range)
```

Finds the highest employee score.

---

## MIN Formula

```excel
=MIN(Test_Score_Range)
```

Finds the lowest employee score.

---

## AVERAGE Formula

```excel
=AVERAGE(Test_Score_Range)
```

Calculates average employee performance.

---

# ⌨️ Excel Shortcuts Used

| Shortcut | Function |
|---|---|
| `Ctrl + C` | Copy |
| `Ctrl + V` | Paste |
| `Ctrl + Z` | Undo |
| `Ctrl + S` | Save workbook |
| `Ctrl + D` | Fill down |
| `Ctrl + R` | Fill right |
| `F2` | Edit selected cell |
| `Alt + =` | AutoSum |
| `Ctrl + Shift + %` | Percentage format |
| `Ctrl + 1` | Open format cells |
| `Ctrl + Arrow Keys` | Fast navigation |

---

# 📈 Analytics Included

The spreadsheet includes:
- Employee performance comparison
- Highest and lowest scores
- Average employee performance
- Pass/Fail analysis
- Performance charts
- Conditional formatting indicators

---

# 🎯 What I Learned

Through this project, I learned:
- How to automate calculations using formulas
- How relative and absolute references work
- How to use conditional formatting effectively
- How to organize large spreadsheets
- How to visualize data using charts
- How business-style reporting systems are structured

---

# 🚀 Future Improvements

Some improvements I plan to add later:
- Pivot tables
- Interactive dashboard
- Department-wise employee analysis
- Dynamic filtering
- Advanced employee analytics

---

# 🛠️ Tools Used

- Microsoft Excel

---

# 📷 Project Screenshots

## Gradebook Overview
(Add screenshot here)

## Charts & Performance Analytics
(Add screenshot here)

---

# 📚 Learning Journey

This project is part of my journey of learning:
- Excel
- Data Analysis
- Business Reporting
- SQL
- Python
- Data Visualization

through practical hands-on projects.

---

# 🌟 About This Project

This is a beginner-level project focused on learning spreadsheet automation and performance analysis using Excel in a practical and business-oriented way.

---

# 🙌 Feedback

Suggestions, improvements, and feedback are always welcome!
````
