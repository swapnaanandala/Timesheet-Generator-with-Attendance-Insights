import pandas as pd
import numpy as np
from datetime import datetime
import os

# ----------------------------
# Utility functions
# ----------------------------

def to_time(x):
    if pd.isna(x) or x == "":
        return None
    try:
        return datetime.strptime(str(x).strip(), "%H:%M").time()
    except:
        return None

def to_date(x):
    if pd.isna(x) or x == "":
        return None
    try:
        return pd.to_datetime(str(x)).date()
    except:
        return None

def hours_between(t1, t2):
    if t1 is None or t2 is None:
        return None
    dt1 = datetime.combine(datetime.today(), t1)
    dt2 = datetime.combine(datetime.today(), t2)
    delta = dt2 - dt1
    return delta.total_seconds() / 3600.0

# ----------------------------
# Core computation
# ----------------------------

def compute_timesheet(df):
    df = df.copy()
    df["date"] = df["date"].apply(to_date)
    df["check_in"] = df["check_in"].apply(to_time)
    df["check_out"] = df["check_out"].apply(to_time)
    df["shift_start"] = df["shift_start"].apply(to_time)
    df["shift_end"] = df["shift_end"].apply(to_time)

    df["break_minutes"] = pd.to_numeric(df["break_minutes"], errors="coerce").fillna(0.0)
    df["expected_hours"] = pd.to_numeric(df["expected_hours"], errors="coerce").fillna(8.0)

    # Worked hours
    df["worked_hours_raw"] = df.apply(lambda r: hours_between(r["check_in"], r["check_out"]), axis=1)
    df["worked_hours"] = df["worked_hours_raw"] - (df["break_minutes"]/60.0)
    df.loc[df["worked_hours"] < 0, "worked_hours"] = np.nan

    # Late arrival / Early exit
    df["late_hours"] = df.apply(
        lambda r: max(0, hours_between(r["shift_start"], r["check_in"])) if r["shift_start"] and r["check_in"] else 0,
        axis=1
    )
    df["early_exit_hours"] = df.apply(
        lambda r: max(0, hours_between(r["check_out"], r["shift_end"])) if r["check_out"] and r["shift_end"] else 0,
        axis=1
    )

    # Overtime & Under hours
    df["overtime_hours"] = (df["worked_hours"] - df["expected_hours"]).clip(lower=0).fillna(0.0)
    df["under_hours"] = (df["expected_hours"] - df["worked_hours"]).clip(lower=0).fillna(0.0)

    # Flags
    df["missing_punch"] = df["worked_hours_raw"].isna()
    df["absent"] = ((df["worked_hours"].fillna(0) == 0) & (df["leave_type"].fillna("").str.lower().isin(["", "unplanned"])))
    df["compliance_alert"] = (
        (df["overtime_hours"] > 2) |
        (df["late_hours"] > 1) |
        (df["missing_punch"]) |
        (df["early_exit_hours"] > 1.0)
    )

    return df

def summarize_month(df):
    agg = df.groupby(["employee_id","employee_name"], as_index=False).agg(
        days_worked = ("worked_hours", lambda x: (x.fillna(0) > 0).sum()),
        total_hours = ("worked_hours","sum"),
        expected_hours_total = ("expected_hours","sum"),
        overtime_total = ("overtime_hours","sum"),
        late_count = ("late_hours", lambda x: (x > 0.01).sum()),
        early_exit_count = ("early_exit_hours", lambda x: (x > 0.01).sum()),
        missing_punches = ("missing_punch","sum"),
        absences = ("absent","sum"),
        compliance_alerts = ("compliance_alert","sum")
    )
    agg["utilization_pct"] = (agg["total_hours"] / agg["expected_hours_total"]) * 100.0
    agg["utilization_pct"] = agg["utilization_pct"].round(1)
    return agg.sort_values("employee_id")

def insights(df, summary):
    return {
        "top_late": summary.sort_values("late_count", ascending=False).head(5),
        "top_overtime": summary.sort_values("overtime_total", ascending=False).head(5),
        "top_missing": summary.sort_values("missing_punches", ascending=False).head(5),
        "top_absent": summary.sort_values("absences", ascending=False).head(5)
    }

def export_excel(daily, summary, ins, path="timesheets_report.xlsx"):
    with pd.ExcelWriter(path, engine="xlsxwriter") as writer:
        daily.to_excel(writer, sheet_name="Timesheet_Daily", index=False)
        summary.to_excel(writer, sheet_name="Summary_By_Employee", index=False)
        ins["top_late"].to_excel(writer, sheet_name="Top_Latecomers", index=False)
        ins["top_overtime"].to_excel(writer, sheet_name="Top_Overtime", index=False)
        ins["top_missing"].to_excel(writer, sheet_name="Top_MissingPunches", index=False)
        ins["top_absent"].to_excel(writer, sheet_name="Top_Absentees", index=False)
    print(f"✅ Report saved to {path}")

# ----------------------------
# Example usage
# ----------------------------

if __name__ == "__main__":
    # Example input CSV format
    # employee_id,employee_name,date,check_in,check_out,break_minutes,shift_start,shift_end,expected_hours,work_type,leave_type
    df_raw = pd.read_csv("attendance.csv")
    
    daily = compute_timesheet(df_raw)
    summary = summarize_month(daily)
    ins = insights(daily, summary)
    
    export_excel(daily, summary, ins, path="timesheets_report.xlsx")
