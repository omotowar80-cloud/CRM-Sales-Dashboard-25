import argparse
import pandas as pd
import matplotlib.pyplot as plt
from pathlib import Path
import shutil
import sys
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.metrics import classification_report

# -----------------------------
# Helpers
# -----------------------------
def project_root(start: Path = None) -> Path:
    here = start or (Path(__file__).resolve() if "__file__" in globals() else Path.cwd())
    return here.parent if here.name == "scripts" else here

def ensure_excel_available(src_excel: Path, raw_dir: Path, canonical_name="CRM_Sales_Dashboard_25.xlsx") -> Path:
    raw_dir.mkdir(parents=True, exist_ok=True)
    dest = raw_dir / canonical_name
    if not dest.exists():
        if src_excel.exists():
            shutil.copy(src_excel, dest)
            print(f"[INFO] Copied Excel to {dest}")
        else:
            raise FileNotFoundError(f"Could not find source Excel at {src_excel}")
    return dest

def detect_sheet_names(xls: pd.ExcelFile) -> tuple[str, str]:
    sheets = [s.lower() for s in xls.sheet_names]
    print(f"[INFO] Available sheets: {xls.sheet_names}")

    def find_sheet(keywords):
        for s in xls.sheet_names:
            lname = s.lower()
            if any(k in lname for k in keywords):
                return s
        return None

    deals = find_sheet(["deal", "pipeline", "opportunit"]) or xls.sheet_names[0]
    teams = find_sheet(["team", "sales"]) or xls.sheet_names[-1]
    print(f"[INFO] Using deals sheet: {deals}")
    print(f"[INFO] Using teams sheet: {teams}")
    return deals, teams

# -----------------------------
# Processing
# -----------------------------
def process_and_merge(deals_df, teams_df):
    if "SalesRep" in deals_df.columns and "SalesRep" in teams_df.columns:
        merged = deals_df.merge(teams_df, on="SalesRep", how="left")
    else:
        print("[WARN] Could not merge on SalesRep — saving deals only")
        merged = deals_df
    return merged

def generate_summary(df: pd.DataFrame) -> dict:
    summary = {}
    if "Amount" in df.columns:
        summary["Total Revenue"] = df["Amount"].sum()
        summary["Average Deal Size"] = df["Amount"].mean()
    if "Stage" in df.columns:
        summary["Deals by Stage"] = df["Stage"].value_counts().to_dict()
    if "Closed" in df.columns:
        win_rate = df["Closed"].mean() * 100
        summary["Win Rate (%)"] = round(win_rate, 2)
    return summary

def save_summary_report(summary: dict, reports_dir: Path):
    reports_dir.mkdir(parents=True, exist_ok=True)
    out = reports_dir / "summary_report.txt"
    with open(out, "w") as f:
        for k, v in summary.items():
            f.write(f"{k}: {v}\n")
    print(f"[INFO] Summary report saved to {out}")

def generate_charts(df: pd.DataFrame, reports_dir: Path):
    reports_dir.mkdir(parents=True, exist_ok=True)

    if "Amount" in df.columns and "Stage" in df.columns:
        plt.figure()
        df.groupby("Stage")["Amount"].sum().plot(kind="bar")
        plt.title("Revenue by Stage")
        plt.ylabel("Revenue")
        plt.savefig(reports_dir / "revenue_by_stage.png")
        plt.close()

    if "SalesRep" in df.columns and "Amount" in df.columns:
        plt.figure()
        df.groupby("SalesRep")["Amount"].sum().sort_values(ascending=False).head(10).plot(kind="bar")
        plt.title("Top 10 Sales Reps by Revenue")
        plt.ylabel("Revenue")
        plt.savefig(reports_dir / "top_sales_reps.png")
        plt.close()

    print(f"[INFO] Charts saved to {reports_dir}")

def train_model(df: pd.DataFrame, reports_dir: Path):
    if "Closed" not in df.columns or "Amount" not in df.columns:
        print("[WARN] Not enough columns for model training.")
        return

    X = df[["Amount"]].fillna(0)  # Simple predictor
    y = df["Closed"].astype(int)

    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.2, random_state=42)
    model = LogisticRegression()
    model.fit(X_train, y_train)

    y_pred = model.predict(X_test)
    report = classification_report(y_test, y_pred)

    reports_dir.mkdir(parents=True, exist_ok=True)
    with open(reports_dir / "model_report.txt", "w") as f:
        f.write(report)

    print(f"[INFO] Model report saved to {reports_dir / 'model_report.txt'}")

# -----------------------------
# Main pipeline
# -----------------------------
def main():
    parser = argparse.ArgumentParser(description="CRM Data Pipeline")
    parser.add_argument("--excel", type=str, default=None, help="Path to Excel source file")
    parser.add_argument("--deals-sheet", type=str, default=None, help="Name of deals sheet")
    parser.add_argument("--teams-sheet", type=str, default=None, help="Name of sales team sheet")
    args = parser.parse_args()

    root = project_root()
    raw_dir = root / "data" / "raw"
    processed_dir = root / "data" / "processed"
    reports_dir = root / "reports"

    processed_dir.mkdir(parents=True, exist_ok=True)
    reports_dir.mkdir(parents=True, exist_ok=True)

    # Ensure Excel exists in raw
    if args.excel:
        src_excel = Path(args.excel)
    else:
        src_excel = raw_dir / "CRM_Sales_Dashboard_25.xlsx"
    excel_path = ensure_excel_available(src_excel, raw_dir)

    # Open workbook
    xls = pd.ExcelFile(excel_path)

    # Detect or use provided sheet names
    deals_sheet, teams_sheet = args.deals_sheet, args.teams_sheet
    if not deals_sheet or not teams_sheet:
        auto_deals, auto_teams = detect_sheet_names(xls)
        deals_sheet = deals_sheet or auto_deals
        teams_sheet = teams_sheet or auto_teams

    # Load data
    deals_df = pd.read_excel(xls, sheet_name=deals_sheet)
    teams_df = pd.read_excel(xls, sheet_name=teams_sheet)

    print(f"[INFO] Deals shape: {deals_df.shape}")
    print(f"[INFO] Teams shape: {teams_df.shape}")

    # Process + merge
    merged = process_and_merge(deals_df, teams_df)
    out_csv = processed_dir / "CRM_Sales_Dashboard_Merged_Enhanced.csv"
    merged.to_csv(out_csv, index=False)
    print(f"[INFO] Processed data saved to {out_csv}")

    # Generate reports
    summary = generate_summary(merged)
    save_summary_report(summary, reports_dir)
    generate_charts(merged, reports_dir)
    train_model(merged, reports_dir)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        sys.exit(1)
