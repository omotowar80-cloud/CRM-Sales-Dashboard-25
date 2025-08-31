import argparse
import pandas as pd
from pathlib import Path
import shutil
import sys

# -----------------------------
# Helpers
# -----------------------------
def project_root(start: Path = None) -> Path:
    """Return the project root (handles __file__ not defined)."""
    here = start or (Path(__file__).resolve() if "__file__" in globals() else Path.cwd())
    return here.parent if here.name == "scripts" else here

def ensure_excel_available(src_excel: Path, raw_dir: Path, canonical_name="CRM_Sales_Dashboard_25.xlsx") -> Path:
    """Ensure the Excel file exists in data/raw. If missing, copy it."""
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
    """Detect deals/opportunities and teams sheets by keywords."""
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
    processed_dir.mkdir(parents=True, exist_ok=True)

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

    # Merge & clean (simplified example)
    if "SalesRep" in deals_df.columns and "SalesRep" in teams_df.columns:
        merged = deals_df.merge(teams_df, on="SalesRep", how="left")
    else:
        print("[WARN] Could not merge on SalesRep — saving deals only")
        merged = deals_df

    # Save processed CSV
    out_csv = processed_dir / "CRM_Sales_Dashboard_Merged_Enhanced.csv"
    merged.to_csv(out_csv, index=False)
    print(f"[INFO] Processed data saved to {out_csv}")

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"[ERROR] {e}", file=sys.stderr)
        sys.exit(1)

python pipeline.py --excel "/mnt/data/Copy of CRM Sales Dashboard '25 (1).xlsx"
