# Power BI Dashboard Setup Guide
## Student Early Warning System - Interactive Dashboard

---

### Why Power BI Cannot Be Auto-Generated
Power BI `.pbix` files are proprietary binary formats that **can only be created using Power BI Desktop** (free download from Microsoft). They cannot be programmatically generated outside of the Power BI ecosystem.

### Quick Setup (15 minutes)

---

## Step 1: Install Power BI Desktop
- Download free from: https://powerbi.microsoft.com/desktop/
- Available for Windows only (use Power BI Service on Mac via browser)

## Step 2: Import Data
1. Open Power BI Desktop
2. Click **"Get Data"** → **"Text/CSV"**
3. Select `Student_performance_data _.csv`
4. Click **"Transform Data"** to open Power Query Editor

## Step 3: Add Calculated Columns (in Power Query)
Click **"Add Column"** → **"Custom Column"** for each:

### Grade Letter
```
= if [GradeClass] = 0 then "A" 
  else if [GradeClass] = 1 then "B" 
  else if [GradeClass] = 2 then "C" 
  else if [GradeClass] = 3 then "D" 
  else "F"
```

### Risk Score
```
= (1 - [GPA] / 4) * 35 + 
  ([Absences] / 30) * 25 + 
  (1 - [StudyTimeWeekly] / 20) * 20 + 
  (1 - [ParentalSupport] / 4) * 10 + 
  ([GradeClass] / 4) * 10
```

### Risk Category
```
= if [RiskScore] <= 30 then "Low"
  else if [RiskScore] <= 55 then "Medium"  
  else if [RiskScore] <= 75 then "High"
  else "Critical"
```

### Pass/Fail Status
```
= if [GradeClass] = 4 then "Fail" else "Pass"
```

Click **"Close & Apply"**

## Step 4: Create DAX Measures
In the **Modeling** tab, create these measures:

```dax
Total Students = COUNTROWS('Student_performance_data _')

Average GPA = AVERAGE('Student_performance_data _'[GPA])

Fail Count = CALCULATE(COUNTROWS('Student_performance_data _'), 'Student_performance_data _'[GradeClass] = 4)

Fail Rate = DIVIDE([Fail Count], [Total Students], 0)

Pass Rate = 1 - [Fail Rate]

At Risk Count = CALCULATE(COUNTROWS('Student_performance_data _'), 'Student_performance_data _'[RiskScore] > 55)

Avg Risk Score = AVERAGE('Student_performance_data _'[RiskScore])
```

## Step 5: Create Dashboard Pages

### Page 1: Academic Overview
| Visual | Type | Fields |
|--------|------|--------|
| KPI Cards (4x) | Card | Total Students, Average GPA, Pass Rate, Fail Rate |
| Grade Distribution | Clustered Bar Chart | Axis: GradeLetter, Values: Count of StudentID |
| Pass vs Fail | Donut Chart | Legend: PassFailStatus, Values: Count |
| GPA by Education | Clustered Bar Chart | Axis: ParentalEducation, Values: Average GPA |

### Page 2: Risk Factor Analysis
| Visual | Type | Fields |
|--------|------|--------|
| Study Time vs GPA | Scatter Plot | X: StudyTimeWeekly, Y: GPA, Color: GradeLetter |
| Absences vs GPA | Scatter Plot | X: Absences, Y: GPA, Color: RiskCategory |
| Correlation Matrix | Table | Factor names + correlation values |
| Tutoring Impact | Grouped Bar | Axis: Tutoring, Values: Average GPA |

### Page 3: Performance Risk Index  
| Visual | Type | Fields |
|--------|------|--------|
| Risk Distribution | Pie/Donut | Legend: RiskCategory, Values: Count |
| Risk Heatmap | Matrix | Rows: RiskCategory, Values: Count, Avg GPA |
| At-Risk Students | Table | StudentID, GPA, Absences, RiskScore (filtered >55) |
| Risk Score Gauge | Gauge | Value: Avg Risk Score, Max: 100 |

### Page 4: Intervention Simulator
| Visual | Type | Fields |
|--------|------|--------|
| What-If Parameter | Slicer | Study Time Increase (1-10) |
| What-If Parameter | Slicer | Absence Reduction (1-10) |
| Projected Improvement | Card | Calculated projected values |
| Current vs Projected | Clustered Bar | Before/After comparison |

### Page 5: Ethics & Safeguards
| Visual | Type | Fields |
|--------|------|--------|
| Factor Weights | Donut Chart | Static data: Factor weights |
| Fairness by Gender | Grouped Bar | Axis: Gender, Values: Fail Rate |
| Fairness by Ethnicity | Grouped Bar | Axis: Ethnicity, Values: Fail Rate |
| Policy Text | Text Box | Privacy and ethics text |

## Step 6: Add What-If Parameters (Page 4)
1. Go to **Modeling** → **New Parameter**
2. Create "Study Time Increase": Min=0, Max=10, Increment=1
3. Create "Absence Reduction": Min=0, Max=10, Increment=1
4. Create measure:
```dax
Projected GPA Lift = [Study Time Increase Value] * 0.04 + [Absence Reduction Value] * 0.03
Projected Fail Rate = MAX(0, [Fail Rate] - [Projected GPA Lift] / 4)
```

## Step 7: Apply Theme
1. Go to **View** → **Themes** → **Browse for themes**
2. Use the included `powerbi_theme.json` file for consistent styling

## Step 8: Save & Share
1. Save as `.pbix` file
2. Publish to Power BI Service for web access
3. Share dashboard link with stakeholders

---
*Dashboard designed to complement the Excel Early Warning Dashboard*
