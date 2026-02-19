import dash
from dash import dcc, html, dash_table, Input, Output, callback
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import pandas as pd
import numpy as np
import os

# ── GitHub Repository ────────────────────────────────────────────────
GITHUB_REPO = "https://github.com/prachisingh342006/data_analytics_project"

# ── Load & Prepare Data ──────────────────────────────────────────────
DATA_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                         "Student_performance_data _.csv")
df = pd.read_csv(DATA_PATH)

grade_map = {0: "A", 1: "B", 2: "C", 3: "D", 4: "F"}
support_map = {0: "None", 1: "Low", 2: "Moderate", 3: "High", 4: "Very High"}
edu_map = {0: "None", 1: "High School", 2: "Some College", 3: "Bachelor's", 4: "Higher"}

df["GradeLetter"] = df["GradeClass"].map(grade_map)
df["SupportLabel"] = df["ParentalSupport"].map(support_map)
df["EducationLabel"] = df["ParentalEducation"].map(edu_map)
df["PassFail"] = df["GradeClass"].apply(lambda x: "Fail" if x == 4 else "Pass")
df["TutoringLabel"] = df["Tutoring"].map({0: "No Tutoring", 1: "With Tutoring"})
df["GenderLabel"] = df["Gender"].map({0: "Female", 1: "Male"})

max_abs = df["Absences"].max()
max_study = df["StudyTimeWeekly"].max()
df["RiskScore"] = (
    (1 - df["GPA"] / 4.0) * 35 +
    (df["Absences"] / max_abs) * 25 +
    (1 - df["StudyTimeWeekly"] / max_study) * 20 +
    (1 - df["ParentalSupport"] / 4.0) * 10 +
    (df["GradeClass"] / 4.0) * 10
).round(2)
df["RiskCategory"] = pd.cut(df["RiskScore"], bins=[0, 30, 55, 75, 100],
                             labels=["Low", "Medium", "High", "Critical"],
                             include_lowest=True)

# ── Colors ────────────────────────────────────────────────────────────
COLORS = {
    "bg": "#0f1923",
    "card": "#1a2733",
    "card_border": "#2a3f52",
    "text": "#e0e6ed",
    "muted": "#8899aa",
    "accent": "#4fc3f7",
    "blue": "#2196f3",
    "green": "#66bb6a",
    "yellow": "#fdd835",
    "orange": "#ffa726",
    "red": "#ef5350",
    "purple": "#ab47bc",
}
RISK_COLORS = {"Low": COLORS["green"], "Medium": COLORS["yellow"],
               "High": COLORS["orange"], "Critical": COLORS["red"]}
GRADE_COLORS = {"A": "#4caf50", "B": "#8bc34a", "C": "#ffc107",
                "D": "#ff9800", "F": "#f44336"}

# ── App ───────────────────────────────────────────────────────────────
app = dash.Dash(__name__, suppress_callback_exceptions=True,
                meta_tags=[{"name": "viewport",
                            "content": "width=device-width, initial-scale=1"}])
app.title = "Student Early Warning Dashboard"
server = app.server  # Expose Flask server for Vercel deployment

# ── Shared Styles ─────────────────────────────────────────────────────
CARD = {
    "backgroundColor": COLORS["card"],
    "borderRadius": "12px",
    "border": f"1px solid {COLORS['card_border']}",
    "padding": "20px",
    "marginBottom": "16px",
}
KPI_BOX = {
    "backgroundColor": COLORS["card"],
    "borderRadius": "12px",
    "border": f"1px solid {COLORS['card_border']}",
    "padding": "20px",
    "textAlign": "center",
    "flex": "1",
    "minWidth": "160px",
}

def kpi_card(label, value, color=COLORS["accent"]):
    return html.Div([
        html.P(label, style={"color": COLORS["muted"], "fontSize": "13px",
                              "marginBottom": "6px", "fontWeight": "500"}),
        html.H3(value, style={"color": color, "margin": "0", "fontSize": "28px",
                               "fontWeight": "700"}),
    ], style=KPI_BOX)

# ── TAB STYLES ────────────────────────────────────────────────────────
tab_style = {
    "backgroundColor": COLORS["card"],
    "color": COLORS["muted"],
    "border": f"1px solid {COLORS['card_border']}",
    "borderRadius": "8px 8px 0 0",
    "padding": "12px 20px",
    "fontWeight": "500",
    "fontSize": "14px",
}
tab_selected_style = {
    **tab_style,
    "backgroundColor": COLORS["blue"],
    "color": "#fff",
    "border": f"1px solid {COLORS['blue']}",
}

# ── Helper: chart layout ─────────────────────────────────────────────
def chart_layout(title=""):
    return dict(
        template="plotly_dark",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
        title=dict(text=title, font=dict(size=16, color=COLORS["text"])),
        font=dict(color=COLORS["muted"], size=12),
        margin=dict(l=50, r=30, t=50, b=50),
        legend=dict(bgcolor="rgba(0,0,0,0)"),
    )

# ════════════════════════════════════════════════════════════════════════
#  PAGE 1 – Academic Overview
# ════════════════════════════════════════════════════════════════════════
def page_academic():
    total = len(df)
    avg_gpa = round(df["GPA"].mean(), 2)
    pass_rate = round((df["GradeClass"] < 4).mean() * 100, 1)
    fail_rate = round((df["GradeClass"] == 4).mean() * 100, 1)
    avg_study = round(df["StudyTimeWeekly"].mean(), 1)
    avg_abs = round(df["Absences"].mean(), 1)

    # Grade distribution bar
    grade_counts = df["GradeLetter"].value_counts().reindex(["A","B","C","D","F"])
    fig_grade = go.Figure(go.Bar(
        x=grade_counts.index, y=grade_counts.values,
        marker_color=[GRADE_COLORS[g] for g in grade_counts.index],
        text=grade_counts.values, textposition="outside",
    ))
    fig_grade.update_layout(**chart_layout("Grade Distribution"))

    # Pass/Fail pie
    pf = df["PassFail"].value_counts()
    fig_pf = go.Figure(go.Pie(
        labels=pf.index, values=pf.values,
        marker_colors=[COLORS["green"], COLORS["red"]],
        hole=0.5, textinfo="label+percent",
    ))
    fig_pf.update_layout(**chart_layout("Pass vs Fail Rate"))

    # GPA by education
    gpa_edu = df.groupby("EducationLabel", observed=True)["GPA"].mean().reindex(
        ["None","High School","Some College","Bachelor's","Higher"])
    fig_edu = go.Figure(go.Bar(
        x=gpa_edu.index, y=gpa_edu.values.round(2),
        marker_color=COLORS["orange"],
        text=gpa_edu.values.round(2), textposition="outside",
    ))
    fig_edu.update_layout(**chart_layout("Avg GPA by Parental Education"))

    return html.Div([
        # KPI row
        html.Div([
            kpi_card("Total Students", f"{total:,}", COLORS["accent"]),
            kpi_card("Average GPA", str(avg_gpa), COLORS["blue"]),
            kpi_card("Pass Rate", f"{pass_rate}%", COLORS["green"]),
            kpi_card("Fail Rate", f"{fail_rate}%", COLORS["red"]),
            kpi_card("Avg Study Hrs/Wk", str(avg_study), COLORS["purple"]),
            kpi_card("Avg Absences", str(avg_abs), COLORS["orange"]),
        ], style={"display": "flex", "gap": "12px", "flexWrap": "wrap",
                  "marginBottom": "20px"}),
        # Charts row
        html.Div([
            html.Div([dcc.Graph(figure=fig_grade)], style={**CARD, "flex": "1"}),
            html.Div([dcc.Graph(figure=fig_pf)], style={**CARD, "flex": "1"}),
        ], style={"display": "flex", "gap": "16px"}),
        html.Div([
            html.Div([dcc.Graph(figure=fig_edu)], style={**CARD, "flex": "1"}),
        ]),
        # GitHub Repository Link
        html.Div([
            html.Div([
                html.Span("📂 ", style={"fontSize":"18px"}),
                html.Span("Source Code & Documentation: ", style={"color":COLORS["muted"],"fontWeight":"500"}),
                html.A(GITHUB_REPO, href=GITHUB_REPO, target="_blank",
                       style={"color":COLORS["accent"],"textDecoration":"none","fontWeight":"600"}),
                html.Span(" — built with Plotly Dash & Python",
                          style={"color":COLORS["muted"],"marginLeft":"8px","fontSize":"13px"}),
            ], style={"display":"flex","alignItems":"center","flexWrap":"wrap","gap":"4px"}),
        ], style={**CARD, "padding":"14px 20px","borderLeft":f"4px solid {COLORS['blue']}"}),
    ])

# ════════════════════════════════════════════════════════════════════════
#  PAGE 2 – Risk Factor Analysis
# ════════════════════════════════════════════════════════════════════════
def page_risk_factors():
    # Ensure categorical columns are strings for color mapping
    df_plot = df.copy()
    df_plot["GradeLetter"] = df_plot["GradeLetter"].astype(str)
    df_plot["RiskCategory"] = df_plot["RiskCategory"].astype(str)

    # Scatter: Study Time vs GPA
    fig_scatter1 = go.Figure()
    for grade in ["A", "B", "C", "D", "F"]:
        mask = df_plot["GradeLetter"] == grade
        if mask.any():
            fig_scatter1.add_trace(go.Scatter(
                x=df_plot.loc[mask, "StudyTimeWeekly"],
                y=df_plot.loc[mask, "GPA"],
                mode="markers",
                name=grade,
                marker=dict(color=GRADE_COLORS.get(grade, "#888"), opacity=0.6, size=6),
            ))
    fig_scatter1.update_layout(**chart_layout("Study Time vs GPA"),
                                xaxis_title="Study Hours/Week", yaxis_title="GPA")

    # Scatter: Absences vs GPA
    fig_scatter2 = go.Figure()
    for risk in ["Low", "Medium", "High", "Critical"]:
        mask = df_plot["RiskCategory"] == risk
        if mask.any():
            fig_scatter2.add_trace(go.Scatter(
                x=df_plot.loc[mask, "Absences"],
                y=df_plot.loc[mask, "GPA"],
                mode="markers",
                name=risk,
                marker=dict(color=RISK_COLORS.get(risk, "#888"), opacity=0.6, size=6),
            ))
    fig_scatter2.update_layout(**chart_layout("Absences vs GPA"),
                                xaxis_title="Number of Absences", yaxis_title="GPA")

    # Correlation bar
    corr_factors = ["StudyTimeWeekly", "Absences", "ParentalSupport",
                    "ParentalEducation", "Tutoring"]
    corr_vals = [df[f].corr(df["GPA"]) for f in corr_factors]
    corr_labels = ["Study Time", "Absences", "Parental Support",
                   "Parental Education", "Tutoring"]
    fig_corr = go.Figure(go.Bar(
        x=corr_vals, y=corr_labels, orientation="h",
        marker_color=[COLORS["green"] if v > 0 else COLORS["red"] for v in corr_vals],
        text=[f"{v:.3f}" for v in corr_vals], textposition="outside",
    ))
    fig_corr.update_layout(**chart_layout("Correlation with GPA"))

    # Tutoring impact
    tut_gpa = df.groupby("TutoringLabel", observed=True)["GPA"].mean()
    fig_tut = go.Figure(go.Bar(
        x=tut_gpa.index, y=tut_gpa.values.round(2),
        marker_color=[COLORS["red"], COLORS["green"]],
        text=tut_gpa.values.round(2), textposition="outside",
    ))
    fig_tut.update_layout(**chart_layout("Tutoring Impact on GPA"))

    # Parental Support
    sup_gpa = df.groupby("SupportLabel", observed=True)["GPA"].mean().reindex(
        ["None","Low","Moderate","High","Very High"])
    fig_sup = go.Figure(go.Bar(
        x=sup_gpa.index, y=sup_gpa.values.round(2),
        marker_color=COLORS["accent"],
        text=sup_gpa.values.round(2), textposition="outside",
    ))
    fig_sup.update_layout(**chart_layout("Avg GPA by Parental Support"))

    return html.Div([
        html.Div([
            html.Div([dcc.Graph(figure=fig_scatter1)], style={**CARD, "flex":"1"}),
            html.Div([dcc.Graph(figure=fig_scatter2)], style={**CARD, "flex":"1"}),
        ], style={"display":"flex","gap":"16px"}),
        html.Div([
            html.Div([dcc.Graph(figure=fig_corr)], style={**CARD, "flex":"1"}),
            html.Div([dcc.Graph(figure=fig_tut)], style={**CARD, "flex":"1"}),
        ], style={"display":"flex","gap":"16px"}),
        html.Div([dcc.Graph(figure=fig_sup)], style=CARD),
    ])

# ════════════════════════════════════════════════════════════════════════
#  PAGE 3 – Performance Risk Index
# ════════════════════════════════════════════════════════════════════════
def page_risk_index():
    # Risk distribution
    risk_order = ["Low","Medium","High","Critical"]
    risk_counts = df["RiskCategory"].value_counts().reindex(risk_order)

    fig_pie = go.Figure(go.Pie(
        labels=risk_counts.index, values=risk_counts.values,
        marker_colors=[RISK_COLORS[r] for r in risk_counts.index],
        hole=0.45, textinfo="label+percent+value",
    ))
    fig_pie.update_layout(**chart_layout("Student Risk Distribution"))

    fig_bar = go.Figure(go.Bar(
        x=risk_counts.index, y=risk_counts.values,
        marker_color=[RISK_COLORS[r] for r in risk_counts.index],
        text=risk_counts.values, textposition="outside",
    ))
    fig_bar.update_layout(**chart_layout("Students per Risk Category"))

    # Risk summary table
    risk_summary = df.groupby("RiskCategory", observed=False).agg(
        Count=("StudentID","count"),
        AvgGPA=("GPA","mean"),
        AvgAbsences=("Absences","mean"),
        AvgRiskScore=("RiskScore","mean"),
    ).reindex(risk_order).reset_index()
    risk_summary["AvgGPA"] = risk_summary["AvgGPA"].round(2)
    risk_summary["AvgAbsences"] = risk_summary["AvgAbsences"].round(1)
    risk_summary["AvgRiskScore"] = risk_summary["AvgRiskScore"].round(1)
    risk_summary.columns = ["Risk Category","Count","Avg GPA","Avg Absences","Avg Risk Score"]

    # Top 20 at-risk
    top20 = df.nlargest(20, "RiskScore")[
        ["StudentID","GPA","GradeLetter","Absences","StudyTimeWeekly","RiskScore","RiskCategory"]
    ].copy()
    top20["GPA"] = top20["GPA"].round(2)
    top20["StudyTimeWeekly"] = top20["StudyTimeWeekly"].round(1)
    top20.columns = ["Student ID","GPA","Grade","Absences","Study Hrs/Wk","Risk Score","Risk Level"]

    return html.Div([
        html.Div([
            html.Div([dcc.Graph(figure=fig_pie)], style={**CARD,"flex":"1"}),
            html.Div([dcc.Graph(figure=fig_bar)], style={**CARD,"flex":"1"}),
        ], style={"display":"flex","gap":"16px"}),
        html.Div([
            html.H4("Risk Category Summary", style={"color": COLORS["text"], "marginBottom":"10px"}),
            dash_table.DataTable(
                data=risk_summary.to_dict("records"),
                columns=[{"name":c,"id":c} for c in risk_summary.columns],
                style_header={"backgroundColor": COLORS["blue"], "color":"#fff",
                              "fontWeight":"bold", "textAlign":"center"},
                style_cell={"backgroundColor": COLORS["card"], "color": COLORS["text"],
                            "textAlign":"center", "padding":"10px",
                            "border": f"1px solid {COLORS['card_border']}"},
                style_data_conditional=[
                    {"if":{"filter_query":'{Risk Category} = "Critical"'},
                     "backgroundColor":"rgba(239,83,80,0.2)","color":COLORS["red"]},
                    {"if":{"filter_query":'{Risk Category} = "High"'},
                     "backgroundColor":"rgba(255,167,38,0.2)","color":COLORS["orange"]},
                ],
            ),
        ], style=CARD),
        html.Div([
            html.H4("Top 20 At-Risk Students", style={"color": COLORS["text"], "marginBottom":"10px"}),
            dash_table.DataTable(
                data=top20.to_dict("records"),
                columns=[{"name":c,"id":c} for c in top20.columns],
                style_header={"backgroundColor":"#d32f2f","color":"#fff",
                              "fontWeight":"bold","textAlign":"center"},
                style_cell={"backgroundColor": COLORS["card"], "color": COLORS["text"],
                            "textAlign":"center","padding":"8px",
                            "border": f"1px solid {COLORS['card_border']}"},
                style_data_conditional=[
                    {"if":{"filter_query":'{Risk Level} = "Critical"'},
                     "backgroundColor":"rgba(239,83,80,0.15)","color":COLORS["red"]},
                ],
                page_size=10,
            ),
        ], style=CARD),
    ])

# ════════════════════════════════════════════════════════════════════════
#  PAGE 4 – Intervention Simulator (Interactive Sliders)
# ════════════════════════════════════════════════════════════════════════
def page_intervention():
    slider_style = {"marginBottom": "20px"}
    label_style = {"color": COLORS["text"], "fontWeight": "500", "marginBottom": "6px"}

    return html.Div([
        html.Div([
            html.H4("Adjust Intervention Parameters", style={"color": COLORS["accent"], "marginBottom": "16px"}),
            html.Label("Study Time Increase (hrs/week)", style=label_style),
            dcc.Slider(id="slider-study", min=0, max=10, step=1, value=5,
                       marks={i: str(i) for i in range(11)},
                       tooltip={"placement":"bottom"}),
            html.Br(),
            html.Label("Absence Reduction (fewer absences)", style=label_style),
            dcc.Slider(id="slider-absence", min=0, max=15, step=1, value=5,
                       marks={i: str(i) for i in range(0,16,3)},
                       tooltip={"placement":"bottom"}),
            html.Br(),
            html.Label("Tutoring Enrollment (% of at-risk)", style=label_style),
            dcc.Slider(id="slider-tutor", min=0, max=100, step=5, value=20,
                       marks={i: f"{i}%" for i in range(0,101,20)},
                       tooltip={"placement":"bottom"}),
            html.Br(),
            html.Label("Cost per Tutor / Semester ($)", style=label_style),
            dcc.Input(id="input-cost", type="number", value=500, min=0, step=50,
                      style={"backgroundColor": COLORS["card"], "color": COLORS["text"],
                             "border": f"1px solid {COLORS['card_border']}",
                             "borderRadius": "6px", "padding": "8px", "width": "150px"}),
        ], style={**CARD, "maxWidth": "600px"}),

        html.Div(id="intervention-results"),
    ])

@callback(
    Output("intervention-results", "children"),
    Input("slider-study", "value"),
    Input("slider-absence", "value"),
    Input("slider-tutor", "value"),
    Input("input-cost", "value"),
)
def update_intervention(study_inc, abs_red, tutor_pct, cost_per):
    total = len(df)
    fail_count = int((df["GradeClass"] == 4).sum())
    fail_rate = fail_count / total
    at_risk = int((df["RiskScore"] > 55).sum())
    avg_gpa_risk = df.loc[df["RiskScore"] > 55, "GPA"].mean()

    gpa_lift = study_inc * 0.04 + abs_red * 0.03
    new_gpa = avg_gpa_risk + gpa_lift
    saved = int(round(fail_count * (gpa_lift / (4 - avg_gpa_risk)))) if avg_gpa_risk < 4 else 0
    saved = min(saved, fail_count)
    new_fail = max(0, fail_count - saved)
    new_fail_rate = new_fail / total
    reduction = fail_rate - new_fail_rate
    target_met = reduction >= fail_rate * 0.2

    tutor_students = int(round(at_risk * tutor_pct / 100))
    total_cost = tutor_students * (cost_per or 500)
    cost_per_saved = int(round(total_cost / saved)) if saved > 0 else 0

    # Comparison chart
    fig = make_subplots(rows=1, cols=2, subplot_titles=["Fail Count", "Fail Rate (%)"],
                        specs=[[{"type":"bar"},{"type":"bar"}]])
    fig.add_trace(go.Bar(x=["Current","Projected"], y=[fail_count, new_fail],
                         marker_color=[COLORS["red"], COLORS["green"]],
                         text=[fail_count, new_fail], textposition="outside"), row=1, col=1)
    fig.add_trace(go.Bar(x=["Current","Projected"],
                         y=[round(fail_rate*100,1), round(new_fail_rate*100,1)],
                         marker_color=[COLORS["red"], COLORS["green"]],
                         text=[f"{fail_rate*100:.1f}%", f"{new_fail_rate*100:.1f}%"],
                         textposition="outside"), row=1, col=2)
    fig.update_layout(**chart_layout(""), showlegend=False, height=350)

    status_color = COLORS["green"] if target_met else COLORS["red"]
    status_text = "✅ TARGET MET — 20% Failure Reduction Achieved!" if target_met else \
                  "❌ TARGET NOT MET — Increase intervention parameters"

    return html.Div([
        html.Div([
            kpi_card("Students Saved", str(saved), COLORS["green"]),
            kpi_card("New Fail Rate", f"{new_fail_rate*100:.1f}%", COLORS["accent"]),
            kpi_card("Fail Rate Reduction", f"{reduction*100:.1f}pp", COLORS["blue"]),
            kpi_card("Program Cost", f"${total_cost:,}", COLORS["orange"]),
            kpi_card("Cost/Student Saved", f"${cost_per_saved:,}", COLORS["purple"]),
        ], style={"display":"flex","gap":"12px","flexWrap":"wrap","marginBottom":"16px"}),
        html.Div(status_text, style={
            "backgroundColor": COLORS["card"], "color": status_color,
            "padding": "16px", "borderRadius": "10px", "textAlign": "center",
            "fontSize": "18px", "fontWeight": "700",
            "border": f"2px solid {status_color}", "marginBottom": "16px",
        }),
        html.Div([dcc.Graph(figure=fig)], style=CARD),
    ])

# ════════════════════════════════════════════════════════════════════════
#  PAGE 5 – Ethics & Safeguards
# ════════════════════════════════════════════════════════════════════════
def page_ethics():
    # Factor weights pie
    fig_wt = go.Figure(go.Pie(
        labels=["GPA (35%)","Absences (25%)","Study Time (20%)",
                "Parental Support (10%)","Grade Class (10%)"],
        values=[35,25,20,10,10],
        marker_colors=[COLORS["blue"], COLORS["orange"], COLORS["accent"],
                       COLORS["yellow"], COLORS["purple"]],
        hole=0.45, textinfo="label+percent",
    ))
    fig_wt.update_layout(**chart_layout("Risk Score Factor Weights"))

    # Fairness: Fail rate by gender
    fair_gender = df.groupby("GenderLabel", as_index=False).agg(
        Total=("GradeClass", "count"),
        FailCount=("GradeClass", lambda x: (x == 4).sum())
    )
    fair_gender["FailRate"] = fair_gender["FailCount"] / fair_gender["Total"]
    fig_gen = go.Figure(go.Bar(
        x=fair_gender["GenderLabel"], y=(fair_gender["FailRate"]*100).round(1),
        marker_color=[COLORS["accent"], COLORS["purple"]],
        text=(fair_gender["FailRate"]*100).round(1).astype(str)+"%", textposition="outside",
    ))
    fig_gen.update_layout(**chart_layout("Fail Rate by Gender"))

    # Fairness: Fail rate by ethnicity
    fair_eth = df.groupby("Ethnicity", as_index=False).agg(
        Total=("GradeClass", "count"),
        FailCount=("GradeClass", lambda x: (x == 4).sum())
    )
    fair_eth["FailRate"] = fair_eth["FailCount"] / fair_eth["Total"]
    fair_eth["Ethnicity"] = "Ethnicity " + fair_eth["Ethnicity"].astype(str)
    fig_eth = go.Figure(go.Bar(
        x=fair_eth["Ethnicity"], y=(fair_eth["FailRate"]*100).round(1),
        marker_color=COLORS["blue"],
        text=(fair_eth["FailRate"]*100).round(1).astype(str)+"%", textposition="outside",
    ))
    fig_eth.update_layout(**chart_layout("Fail Rate by Ethnicity"))

    transparency_data = [
        {"Factor":"GPA","Weight":"35%","Reason":"Primary academic indicator",
         "Bias Risk":"Low","Mitigation":"Validated by registrar"},
        {"Factor":"Absences","Weight":"25%","Reason":"Predictor of disengagement",
         "Bias Risk":"Medium","Mitigation":"Context review before action"},
        {"Factor":"Study Time","Weight":"20%","Reason":"Effort indicator",
         "Bias Risk":"Medium","Mitigation":"Cross-reference with grades"},
        {"Factor":"Parental Support","Weight":"10%","Reason":"Environmental factor",
         "Bias Risk":"High","Mitigation":"Never used alone"},
        {"Factor":"Grade Class","Weight":"10%","Reason":"Current standing",
         "Bias Risk":"Low","Mitigation":"Confirmation only"},
    ]
    policies = [
        "🔒 **Data Minimization** — Only academically relevant data used. No names or addresses.",
        "🎯 **Purpose Limitation** — Data used solely for early intervention, not disciplinary decisions.",
        "🔑 **Access Control** — Dashboard restricted to authorized academic advisors only.",
        "📋 **Consent & Transparency** — Students informed about data usage. Opt-out available.",
        "⏱️ **Retention Policy** — Scores recalculated each semester. History purged after 2 years.",
        "⚖️ **Right to Challenge** — Students may contest categorization via academic affairs.",
        "🔍 **Bias Auditing** — Annual review across demographics for disparate impact.",
        "🆔 **De-identification** — Reports use student IDs only.",
    ]

    return html.Div([
        html.Div([
            html.Div([dcc.Graph(figure=fig_wt)], style={**CARD,"flex":"1"}),
            html.Div([
                html.H4("Labeling Transparency Matrix",
                         style={"color": COLORS["text"],"marginBottom":"10px"}),
                dash_table.DataTable(
                    data=transparency_data,
                    columns=[{"name":c,"id":c} for c in transparency_data[0].keys()],
                    style_header={"backgroundColor": COLORS["blue"],"color":"#fff",
                                  "fontWeight":"bold","textAlign":"center"},
                    style_cell={"backgroundColor": COLORS["card"],"color": COLORS["text"],
                                "textAlign":"center","padding":"8px",
                                "border": f"1px solid {COLORS['card_border']}",
                                "whiteSpace":"normal","minWidth":"80px"},
                ),
            ], style={**CARD,"flex":"1.5"}),
        ], style={"display":"flex","gap":"16px"}),
        html.Div([
            html.Div([dcc.Graph(figure=fig_gen)], style={**CARD,"flex":"1"}),
            html.Div([dcc.Graph(figure=fig_eth)], style={**CARD,"flex":"1"}),
        ], style={"display":"flex","gap":"16px"}),
        html.Div([
            html.H4("Student Data Privacy Policy",
                     style={"color": COLORS["text"],"marginBottom":"12px"}),
            html.Div([dcc.Markdown(p, style={"color": COLORS["text"], "marginBottom":"8px",
                                              "fontSize":"14px"}) for p in policies]),
        ], style=CARD),
    ])

# ════════════════════════════════════════════════════════════════════════
#  MAIN LAYOUT
# ════════════════════════════════════════════════════════════════════════
app.layout = html.Div([
    # Header
    html.Div([
        html.Div([
            html.H1("🎓 Student Early Warning Dashboard",
                    style={"margin":"0","fontSize":"26px","fontWeight":"700"}),
            html.P("Reducing university failure rates by 20% through data-driven early intervention",
                   style={"margin":"4px 0 0","fontSize":"14px","color": COLORS["muted"]}),
        ], style={"flex":"1"}),
        html.A(
            html.Div([
                html.Img(src="https://cdn.jsdelivr.net/gh/devicons/devicon/icons/github/github-original.svg",
                         style={"width":"22px","height":"22px","filter":"invert(1)","marginRight":"8px","verticalAlign":"middle"}),
                html.Span("View on GitHub", style={"verticalAlign":"middle"}),
            ], style={"display":"flex","alignItems":"center"}),
            href=GITHUB_REPO, target="_blank",
            style={"color":"#4fc3f7","textDecoration":"none","fontSize":"14px",
                   "fontWeight":"600","backgroundColor":"rgba(33,150,243,0.15)",
                   "padding":"10px 20px","borderRadius":"8px",
                   "border":f"1px solid {COLORS['blue']}","transition":"all 0.2s"},
        ),
    ], style={"backgroundColor": COLORS["card"], "padding": "20px 30px",
              "borderBottom": f"2px solid {COLORS['blue']}",
              "color": COLORS["text"],
              "display":"flex","alignItems":"center","justifyContent":"space-between"}),

    # Tabs
    html.Div([
        dcc.Tabs(id="tabs", value="tab-1", children=[
            dcc.Tab(label="📊 Academic Overview", value="tab-1",
                    style=tab_style, selected_style=tab_selected_style),
            dcc.Tab(label="🔍 Risk Factor Analysis", value="tab-2",
                    style=tab_style, selected_style=tab_selected_style),
            dcc.Tab(label="⚠️ Risk Index", value="tab-3",
                    style=tab_style, selected_style=tab_selected_style),
            dcc.Tab(label="🎯 Intervention Simulator", value="tab-4",
                    style=tab_style, selected_style=tab_selected_style),
            dcc.Tab(label="🛡️ Ethics & Safeguards", value="tab-5",
                    style=tab_style, selected_style=tab_selected_style),
        ], style={"marginBottom": "0"}),
    ], style={"padding": "16px 30px 0"}),

    # Page content
    html.Div(id="tab-content", style={"padding": "20px 30px"}),

], style={"backgroundColor": COLORS["bg"], "minHeight": "100vh", "fontFamily": "Segoe UI, Roboto, sans-serif"})

@callback(Output("tab-content", "children"), Input("tabs", "value"))
def render_tab(tab):
    if tab == "tab-1": return page_academic()
    if tab == "tab-2": return page_risk_factors()
    if tab == "tab-3": return page_risk_index()
    if tab == "tab-4": return page_intervention()
    if tab == "tab-5": return page_ethics()

# ════════════════════════════════════════════════════════════════════════
if __name__ == "__main__":
    print("\n" + "="*60)
    print("  🎓 Student Early Warning Dashboard")
    print("  Open browser: http://127.0.0.1:8050")
    print("="*60 + "\n")
    app.run(debug=True, port=8050)
