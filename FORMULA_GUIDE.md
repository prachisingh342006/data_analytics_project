# FORMULA QUICK REFERENCE GUIDE
## Student Early Warning Dashboard

---

## 🎯 MOST IMPORTANT FORMULAS TO UNDERSTAND

### 1. RISK SCORE FORMULA (Performance Risk Index, Column H)

**Purpose:** Calculate a 0-100 risk score for each student

**Formula Breakdown:**
```excel
=IF(GPA<2, 35, IF(GPA<2.5, 20, 0))           // GPA Component (35% weight)
  └─ If GPA < 2.0 → Add 35 points (critical)
  └─ If GPA 2.0-2.49 → Add 20 points (warning)
  └─ If GPA ≥ 2.5 → Add 0 points (ok)

+IF(Absences>15, 25, IF(Absences>10, 15, 0))  // Absence Component (25% weight)
  └─ If >15 days absent → Add 25 points (critical)
  └─ If 11-15 days → Add 15 points (warning)
  └─ If ≤10 days → Add 0 points (ok)

+IF(StudyTime<10, 20, IF(StudyTime<15, 10, 0)) // Study Time Component (20% weight)
  └─ If <10 hrs/week → Add 20 points (critical)
  └─ If 10-14 hrs/week → Add 10 points (warning)
  └─ If ≥15 hrs/week → Add 0 points (ok)

+IF(AND(Tutoring=0, Support<2), 10, 0)        // Support Component (10% weight)
  └─ If NO tutoring AND low support → Add 10 points
  └─ Otherwise → Add 0 points

+IF(GradeClass>=3, 10, 0)                     // Grade Component (10% weight)
  └─ If D or F grade → Add 10 points
  └─ If A, B, or C → Add 0 points
```

**Example for Student 1001:**
- GPA = 2.93 → 0 points (≥2.5)
- Absences = 7 → 0 points (≤10)
- Study Time = 19.8 → 0 points (≥15)
- Tutoring = 1, Support = 2 → 0 points (has support)
- Grade = 2 (C) → 0 points (<3)
- **Total Risk Score = 0** → Low Risk

**Example for At-Risk Student:**
- GPA = 1.5 → 35 points
- Absences = 20 → 25 points
- Study Time = 5 → 20 points
- Tutoring = 0, Support = 1 → 10 points
- Grade = 4 (F) → 10 points
- **Total Risk Score = 100** → Critical Risk

---

### 2. RISK LEVEL CLASSIFICATION

**Purpose:** Convert numeric risk score to category

```excel
=IF(RiskScore>=60, "Critical",
   IF(RiskScore>=40, "High",
      IF(RiskScore>=20, "Medium", "Low")))
```

**Thresholds:**
- 60-100 = Critical (immediate intervention)
- 40-59 = High (weekly monitoring)
- 20-39 = Medium (bi-weekly check-in)
- 0-19 = Low (standard support)

---

### 3. CORRELATION ANALYSIS

**Purpose:** Measure relationship strength between factors and GPA

```excel
=CORREL(FactorRange, GPARange)
```

**Example:**
```excel
=CORREL(F6:F2398, O6:O2398)  // Study Time vs GPA
```

**Interpretation:**
- +0.7 to +1.0 = Strong positive (more study → higher GPA)
- +0.3 to +0.7 = Moderate positive
- -0.3 to +0.3 = Weak/no relationship
- -0.3 to -0.7 = Moderate negative
- -0.7 to -1.0 = Strong negative (more absences → lower GPA)

---

### 4. CONDITIONAL AVERAGING (AVERAGEIFS)

**Purpose:** Calculate average GPA for students meeting specific criteria

**Study Time Analysis:**
```excel
=AVERAGEIFS(
   O6:O2398,           // Average this: GPA column
   F6:F2398, ">=10",   // Where: Study time ≥ 10 hours
   F6:F2398, "<15"     // And: Study time < 15 hours
)
```

**Attendance Analysis:**
```excel
=AVERAGEIFS(
   O6:O2398,           // Average GPA
   G6:G2398, ">=0",    // Where absences ≥ 0
   G6:G2398, "<5"      // And absences < 5
)
```

---

### 5. PASS/FAIL RATE CALCULATIONS

**Pass Rate (Grades A-D):**
```excel
=COUNTIFS(P6:P2398, "<4") / COUNTA(P6:P2398) * 100
  └─ Count students with grade 0-3
  └─ Divide by total students
  └─ Multiply by 100 for percentage
```

**Failure Rate (Grade F):**
```excel
=COUNTIF(P6:P2398, 4) / COUNTA(P6:P2398) * 100
  └─ Count students with grade 4 (F)
  └─ Divide by total students
  └─ Multiply by 100 for percentage
```

**Double-check:** Pass Rate + Fail Rate should = 100%

---

### 6. INTERVENTION IMPACT SIMULATION

**GPA Projection Formula:**
```excel
=CurrentGPA + SUM(AllImpacts)

Where each impact is calculated as:
Impact = ParameterChange × CorrelationCoefficient
```

**Example:**
```excel
Current GPA = 2.75

Impact from +5 hrs study time = 5 × 0.15 = +0.75
Impact from +10% attendance = 10 × 0.02 = +0.20
Impact from +25% tutoring = 25 × 0.01 = +0.25
Impact from +1 parent support = 1 × 0.08 = +0.08

Projected GPA = 2.75 + 0.75 + 0.20 + 0.25 + 0.08 = 4.03
```

**Note:** GPA capped at 4.0 in practice

---

### 7. COST-BENEFIT ANALYSIS

**Total Intervention Cost:**
```excel
=AtRiskStudentCount × AVERAGE(CostPerIntervention)
```

**ROI Calculation:**
```excel
ROI = (PreventedFailures × TuitionPerStudent) - TotalInterventionCost

Where PreventedFailures = CurrentFailures - ProjectedFailures
```

**Example:**
```
Current failures: 200 students
Projected failures: 150 students (after intervention)
Prevented failures: 50 students

ROI = (50 × $40,000) - (200 × $2,000)
    = $2,000,000 - $400,000
    = $1,600,000 net benefit
```

---

### 8. FAIRNESS/BIAS MONITORING

**Variance Calculation:**
```excel
=ABS(GroupAverage - PopulationAverage) / PopulationAverage
```

**Example:**
```
Overall avg risk score: 35.0
Male avg risk score: 36.5
Female avg risk score: 33.8

Male variance = ABS(36.5 - 35.0) / 35.0 = 0.043 = 4.3% ✓ Fair
Female variance = ABS(33.8 - 35.0) / 35.0 = 0.034 = 3.4% ✓ Fair
```

**Threshold:** <10% variance = Fair, ≥10% = Review needed

---

## 🔧 HOW TO MODIFY THE SIMULATOR (PAGE 4)

### Step-by-Step Instructions:

1. **Go to "Intervention Simulator" sheet**

2. **Locate the yellow highlighted cells in Column C (rows 7-10):**
   - Row 7: Study Time increase (default: 5)
   - Row 8: Attendance improvement (default: 10)
   - Row 9: Tutoring enrollment (default: 25)
   - Row 10: Parental engagement (default: 1)

3. **Change any value:**
   - Example: Change study time from 5 to 10 hours
   - Type `10` in cell C7
   - Press Enter

4. **Watch automatic updates:**
   - Column D (New Value) recalculates
   - Column E (Impact on GPA) updates
   - Projected Outcomes section (rows 16-19) updates
   - Cost-Benefit Analysis (rows 24-30) updates

5. **Try different scenarios:**
   - **Aggressive:** 10, 20, 40, 2
   - **Moderate:** 5, 10, 25, 1
   - **Conservative:** 3, 5, 15, 0.5
   - **Zero intervention:** 0, 0, 0, 0 (baseline)

6. **Compare results:**
   - Check "Target Met?" column (F)
   - Review Net Benefit (D29)
   - Examine ROI Ratio (D30)

---

## 📊 UNDERSTANDING THE DATA

### Grade Class Mapping:
```
0 = A (Excellent)   → GPA 3.5-4.0   → PASS
1 = B (Good)        → GPA 3.0-3.49  → PASS
2 = C (Average)     → GPA 2.0-2.99  → PASS
3 = D (Below Avg)   → GPA 1.0-1.99  → PASS (but at-risk)
4 = F (Fail)        → GPA 0.0-0.99  → FAIL
```

### Key Thresholds:
- **At-Risk:** GPA < 2.0 OR Grade Class ≥ 3
- **High Absences:** > 15 days
- **Low Study Time:** < 10 hours/week
- **No Support:** Tutoring = 0 AND Parental Support < 2

---

## 🎯 COMMON TASKS

### Task 1: Find Critical Risk Students
1. Go to "Performance Risk Index" sheet
2. Click on Risk Score column (H)
3. Sort descending (highest to lowest)
4. Filter Risk Level column (I) for "Critical"
5. Review StudentID (A) and Recommended Action (K)

### Task 2: Analyze Study Time Impact
1. Go to "Risk Factor Analysis" sheet
2. Find "Study Time vs Performance" table
3. Compare Avg GPA across different hour ranges
4. Note the correlation coefficient at top of sheet

### Task 3: Calculate Budget Needed
1. Go to "Performance Risk Index" sheet
2. Scroll to "Risk Distribution Summary"
3. Check "Total Cost" column (F)
4. Sum all risk levels for total budget

### Task 4: Test "What If" Scenario
1. Go to "Intervention Simulator" sheet
2. Modify yellow cells in column C
3. Check if "Target Met?" shows ✓ YES
4. Review Net Benefit and ROI Ratio
5. Adjust until goals are met at acceptable cost

### Task 5: Check for Bias
1. Go to "Ethics & Safeguards" sheet
2. Find "Algorithmic Bias Monitoring" section
3. Review Fairness Metric column (D)
4. Ensure Status shows "✓ Fair" for all groups
5. If "⚠ Review" appears, investigate disparities

---

## 🚨 TROUBLESHOOTING

### Issue: #DIV/0! Error
**Cause:** Division by zero (no students in category)
**Fix:** Check that data range is correct and not empty

### Issue: #VALUE! Error
**Cause:** Non-numeric data in calculation
**Fix:** Verify data types in source columns

### Issue: Correlations showing as 0
**Cause:** Incorrect cell ranges
**Fix:** Ensure ranges start at same row and have equal length

### Issue: Percentages showing as decimals
**Cause:** Number format not set
**Fix:** Select cells → Format → Percentage

### Issue: Formulas not updating
**Cause:** Manual calculation mode
**Fix:** Press F9 or set to automatic (Formulas → Calculation Options → Automatic)

---

## 💡 PRO TIPS

1. **Always keep a backup:** Save a copy before making major changes

2. **Test with small changes first:** Modify one parameter at a time to see its effect

3. **Use scenario comparison:** Keep notes on different parameter combinations

4. **Cross-reference metrics:** If GPA improves, pass rate should improve too

5. **Watch for unrealistic projections:** GPA can't exceed 4.0, percentages can't exceed 100%

6. **Document your assumptions:** Note why you chose specific intervention values

7. **Review correlations regularly:** Real-world data may differ from assumptions

8. **Consult with stakeholders:** Get input from advisors on realistic intervention levels

9. **Track actual outcomes:** Compare predictions to real results for validation

10. **Update formulas as needed:** Adjust correlation coefficients based on observed data

---

## 📞 FORMULA SUPPORT

**Need help with a specific formula?**

1. Click on the cell with the formula
2. Look at the formula bar at the top
3. Press F2 to see cell references highlighted
4. Use this guide to understand each component

**Still stuck?**
- Contact: analytics@university.edu
- Include: Sheet name, cell reference, and what you're trying to calculate

---

**Last Updated:** February 18, 2026  
**Version:** 1.0  
**Compatibility:** Excel 2016 or later, Google Sheets (most formulas)
