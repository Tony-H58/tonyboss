#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import pandas as pd
import numpy as np
from datetime import datetime
import os

# File paths
MONTHLY_QF_DIR = r'E:\88. Claude\11_inspdata_monthly\insprecord_QF'
OUTPUT_DIR = r'E:\88. Claude\14_qc performance'

# Column indices (0-based)
COL_SOURCE = 0      # 工廠 vs 品管
COL_DATE = 1        # Inspection date
COL_STYLE = 3       # Style/款號
COL_COLOR = 4       # Color/顏色
COL_QC = 22         # QC name (only in factory records)

# For FACTORY records (工廠):
# - Col 25 = Pass/Fail, Col 33 = C級%, Col 35 = 百碼瑕疵點數
FACTORY_PASS_FAIL = 24     # Col 25 (0-indexed) - Pass/Fail
FACTORY_C_GRADE_PCT = 32   # Col 33 (0-indexed) - C級%
FACTORY_DEFECT_PER_HUNDRED = 34  # Col 35 (0-indexed) - 百碼瑕疵點數

# For QC records (品管):
# - Col 25 = Pass/Fail, Col 26 = receipt qty, Col 27 = sample qty
# - Col 33 = C級%, Col 35 = 百碼瑕疵點數
QC_PASS_FAIL = 24          # Col 25 (0-indexed) - Pass/Fail
QC_RECEIPT_QTY = 25        # Col 26 (0-indexed) - Receipt qty
QC_SAMPLE_QTY = 26         # Col 27 (0-indexed) - Sample qty
QC_C_GRADE_PCT = 32        # Col 33 (0-indexed) - C級%
QC_DEFECT_PER_HUNDRED = 34  # Col 35 (0-indexed) - 百碼瑕疵點數

def load_data(year_month='26-1~26-6'):
    """Load data from monthly insprecord_QF folder based on year_month period"""
    print(f"Loading insprecord_QF data for period: {year_month}...")

    # Parse year_month format: "26-1~26-6"
    parts = year_month.split('~')
    start_ym = parts[0].strip()  # "26-1"
    end_ym = parts[1].strip() if len(parts) > 1 else start_ym  # "26-6"

    # Convert month string to numbers for comparison
    start_yy, start_mm = start_ym.split('-')
    end_yy, end_mm = end_ym.split('-')
    start_yymmdd = f"{start_yy}0{start_mm}02"  # "260102"
    end_yymmdd = f"{end_yy}0{end_mm}31"  # "260631" (use 31 as max possible)

    # Find all matching files in MONTHLY_QF_DIR
    import glob
    pattern = os.path.join(MONTHLY_QF_DIR, 'insprecord_QF_*.xlsx')
    matching_files = sorted(glob.glob(pattern))

    if not matching_files:
        raise FileNotFoundError(f"No insprecord_QF files found in {MONTHLY_QF_DIR}")

    # Filter files that overlap with the requested period
    # File format: insprecord_QF_YYMMDD-YYMMDD.xlsx
    dfs = []
    for filepath in matching_files:
        filename = os.path.basename(filepath)
        # Extract date range from filename
        date_part = filename.replace('insprecord_QF_', '').replace('.xlsx', '')
        file_start, file_end = date_part.split('-')

        # Check if file overlaps with requested period
        if file_end >= start_yymmdd and file_start <= end_yymmdd:
            print(f"  Loading: {filename}")
            df = pd.read_excel(filepath, sheet_name=0)
            dfs.append(df)

    if not dfs:
        raise FileNotFoundError(f"No files found matching period {year_month}")

    all_df = pd.concat(dfs, ignore_index=True)
    print(f"  Total rows: {len(all_df)}")

    return all_df

def separate_qc_factory(df):
    """Separate QC and Factory records"""
    col1 = df.iloc[:, COL_SOURCE].fillna('').astype(str)

    # Get unique values to determine actual category values
    unique_vals = col1.unique()

    # Try to identify which is which
    # Usually '品管' = QC/Inspection, '工廠' = Factory
    qc_val = None
    factory_val = None

    for val in unique_vals:
        if val and len(val) > 0:
            # Check bytes to identify
            try:
                byte_rep = val.encode('utf-8')
                # 品管 = b'\xe5\x93\x81\xe7\xae\xa1'
                # 工廠 = b'\xe5\xb7\xa5\xe5\xbb\xa0'
                if byte_rep == b'\xe5\x93\x81\xe7\xae\xa1':
                    qc_val = val
                elif byte_rep == b'\xe5\xb7\xa5\xe5\xbb\xa0':
                    factory_val = val
            except:
                pass

    # Fallback: just use first two unique values
    if qc_val is None or factory_val is None:
        non_empty = [v for v in unique_vals if v and len(v) > 0]
        if len(non_empty) >= 2:
            # Assume first is QC, second is Factory
            qc_val = non_empty[0]
            factory_val = non_empty[1]

    qc_df = df[col1 == qc_val].copy() if qc_val else df.iloc[0:0]
    factory_df = df[col1 == factory_val].copy() if factory_val else df.iloc[0:0]

    print(f"\nData separation:")
    print(f"  QC records ({repr(qc_val)}): {len(qc_df)}")
    print(f"  Factory records ({repr(factory_val)}): {len(factory_df)}")

    return qc_df, factory_df

def calculate_qc_kpi(qc_records, factory_records):
    """
    Calculate QC KPI based on QC (品管) records:
    A. Work Volume (10%): Total sample qty & PO count
    B. Execution Rate (20%): Total sample qty / Total receipt qty
    C. Work Efficiency (30%): Total sample qty / Total inspection days
    D. Accuracy (30%): Compare with factory records for same style+color

    Note: QC records have QC names and complete data; Factory records are for accuracy comparison
    """

    # Use QC records which have QC names and complete data
    work_df = qc_records[qc_records.iloc[:, COL_QC].notna()].copy()

    results = []

    for qc_name in work_df.iloc[:, COL_QC].unique():
        qc_person_df = work_df[work_df.iloc[:, COL_QC] == qc_name]

        # A. Work Volume: Total sample qty and PO count
        total_sample_qty = pd.to_numeric(qc_person_df.iloc[:, QC_SAMPLE_QTY], errors='coerce').sum()
        po_count = len(qc_person_df)  # Count of PO/inspection records

        # B. Execution Rate: Total sample qty / Total receipt qty
        total_receipt_qty = pd.to_numeric(qc_person_df.iloc[:, QC_RECEIPT_QTY], errors='coerce').sum()
        execution_rate = (total_sample_qty / total_receipt_qty * 100) if total_receipt_qty > 0 else 0

        # C. Work Efficiency: Total sample qty / Number of inspection days
        dates = pd.to_datetime(qc_person_df.iloc[:, COL_DATE], errors='coerce').dropna()
        if len(dates) > 0:
            # Count distinct inspection days (not date range)
            inspection_days = dates.nunique()  # Number of different dates
            daily_efficiency = total_sample_qty / max(inspection_days, 1)
        else:
            daily_efficiency = 0

        # D. Accuracy: Count discrepancies with factory records
        debug_name = qc_name if qc_name == 'NgocBich' else None
        discrepancy_count = calculate_accuracy_discrepancy(qc_person_df, factory_records, qc_name=debug_name)

        # Get location (use most common location if multiple)
        locations = qc_person_df.iloc[:, 7].fillna('').astype(str).unique()
        location = [l for l in locations if l and l != ''][0] if any(l and l != '' for l in locations) else 'Unknown'

        results.append({
            'QC': qc_name,
            'Location': location,
            'A_SampleQty': round(total_sample_qty, 0),
            'A_POCount': po_count,
            'B_ExecutionRate': round(execution_rate, 2),
            'C_DailyEfficiency': round(daily_efficiency, 2),
            'D_DiscrepancyCount': discrepancy_count
        })

    return pd.DataFrame(results)

def calculate_accuracy_discrepancy(qc_person_df, factory_df, qc_name=None):
    """
    Calculate discrepancy count for QC vs Factory:
    (1) QC Pass but Factory Fail = 1
    (2) Factory C級% - QC C級% > 10% = 1
    """

    if len(factory_df) == 0:
        return 0, {}

    # Create pairing key: style + color
    qc_person_df = qc_person_df.copy()
    qc_person_df['pair_key'] = qc_person_df.iloc[:, COL_STYLE].astype(str) + '_' + qc_person_df.iloc[:, COL_COLOR].astype(str)

    factory_df_copy = factory_df.copy()
    factory_df_copy['pair_key'] = factory_df_copy.iloc[:, COL_STYLE].astype(str) + '_' + factory_df_copy.iloc[:, COL_COLOR].astype(str)

    # Find matching pairs
    common_keys = set(qc_person_df['pair_key'].unique()) & set(factory_df_copy['pair_key'].unique())

    if len(common_keys) == 0:
        return 0

    discrepancy_count = 0
    detail_count = {'PassFail': 0, 'C_Grade': 0}

    for key in common_keys:
        qc_rows = qc_person_df[qc_person_df['pair_key'] == key]
        factory_rows = factory_df_copy[factory_df_copy['pair_key'] == key]

        if len(qc_rows) == 0 or len(factory_rows) == 0:
            continue

        for _, qc_row in qc_rows.iterrows():
            for _, factory_row in factory_rows.iterrows():
                has_discrepancy = False
                reason = None

                # (1) QC Pass but Factory Fail
                qc_pass = str(qc_row.iloc[QC_PASS_FAIL]).strip() if pd.notna(qc_row.iloc[QC_PASS_FAIL]) else ""
                factory_pass = str(factory_row.iloc[FACTORY_PASS_FAIL]).strip() if pd.notna(factory_row.iloc[FACTORY_PASS_FAIL]) else ""

                if qc_pass == "Pass" and factory_pass == "Fail":
                    has_discrepancy = True
                    reason = 'PassFail'
                    detail_count['PassFail'] += 1

                # (2) Factory C級% - QC C級% > 10%
                if not has_discrepancy:
                    try:
                        qc_c_pct = float(qc_row.iloc[QC_C_GRADE_PCT]) if pd.notna(qc_row.iloc[QC_C_GRADE_PCT]) else 0
                        factory_c_pct = float(factory_row.iloc[FACTORY_C_GRADE_PCT]) if pd.notna(factory_row.iloc[FACTORY_C_GRADE_PCT]) else 0

                        if factory_c_pct - qc_c_pct > 10:
                            has_discrepancy = True
                            reason = 'C_Grade'
                            detail_count['C_Grade'] += 1
                    except:
                        pass

                if has_discrepancy:
                    discrepancy_count += 1

    if qc_name:
        print(f"\n{qc_name} - Discrepancy breakdown: Pass/Fail={detail_count['PassFail']}, C級%={detail_count['C_Grade']}, Total={discrepancy_count}")

    return discrepancy_count

def calculate_scores(kpi_df):
    """Convert KPI values to 0-100 scores and calculate overall"""

    # Scoring functions
    def score_sample_qty(qty):
        # More sample = better (scale to 0-100)
        return min(100, qty / 50)  # 50+ codes = 100 points

    def score_po_count(count):
        # More POs = more work (scale to 0-100)
        return min(100, count / 5)  # 5+ POs = 100 points

    def score_rate(rate):
        # Execution rate 0-100%, clamp to 0-100 score
        return min(100, max(0, rate))

    def score_efficiency(eff):
        # Daily efficiency (codes/day)
        return min(100, eff)  # 100+ codes/day = 100 points

    def score_discrepancy(discrepancy_df):
        # Calculate score based on 10-point scale
        # Min discrepancy = 100 points, each step down by 10 points
        min_count = discrepancy_df['D_DiscrepancyCount'].min()
        max_count = discrepancy_df['D_DiscrepancyCount'].max()

        if min_count == max_count:
            # All same count = all 100 points
            return pd.Series(100, index=discrepancy_df.index)

        # Distribute to 10 equal intervals
        interval = (max_count - min_count) / 10
        scores = []
        for count in discrepancy_df['D_DiscrepancyCount']:
            steps = (count - min_count) / interval if interval > 0 else 0
            score = max(0, 100 - steps * 10)
            scores.append(score)

        return pd.Series(scores, index=discrepancy_df.index)

    # Apply scoring
    kpi_df['A_SampleScore'] = kpi_df['A_SampleQty'].apply(score_sample_qty)
    kpi_df['A_POScore'] = kpi_df['A_POCount'].apply(score_po_count)
    kpi_df['B_Score'] = kpi_df['B_ExecutionRate'].apply(score_rate)
    kpi_df['C_Score'] = kpi_df['C_DailyEfficiency'].apply(score_efficiency)
    kpi_df['D_Score'] = score_discrepancy(kpi_df)

    # Weighted average (A:10% + B:20% + C:30% + D:30% = 90分)
    # A includes both sample qty and PO count
    # All scores are 0-100, weights sum to 90 maximum
    kpi_df['Overall_Score'] = (
        ((kpi_df['A_SampleScore'] + kpi_df['A_POScore']) / 2) * 0.10 +
        kpi_df['B_Score'] * 0.30 +
        kpi_df['C_Score'] * 0.30 +
        kpi_df['D_Score'] * 0.30
    )

    return kpi_df

def generate_html(kpi_df):
    """Generate HTML report"""

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Define region order: 中國, 台灣, 印尼, 越南
    region_order = {'CHN': 0, 'TWN': 1, 'IND': 2, 'VIN': 3}
    kpi_df['region_sort'] = kpi_df['Location'].map(lambda x: region_order.get(x, 999))

    # Sort by region first (按指定順序), then by overall score (高低)
    kpi_df = kpi_df.sort_values(by=['region_sort', 'Overall_Score'], ascending=[True, False])

    html = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <title>QC Performance KPI Report</title>
    <style>
        * {{ margin: 0; padding: 0; }}
        body {{ font-family: Arial, sans-serif; background: #f5f5f5; padding: 20px; }}
        .container {{ max-width: 1400px; margin: 0 auto; background: white; padding: 30px; border-radius: 8px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }}
        h1 {{ color: #333; margin-bottom: 10px; }}
        .info {{ color: #666; font-size: 14px; margin-bottom: 20px; }}
        table {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
        th {{ background-color: #90EE90; color: #333; padding: 12px; text-align: left; font-weight: bold; border: 1px solid #ddd; }}
        td {{ padding: 12px; border: 1px solid #ddd; text-align: center; }}
        td:first-child {{ text-align: left; }}
        tr:nth-child(even) {{ background-color: #f9f9f9; }}
        tr:hover {{ background-color: #f0f0f0; }}
        .metric {{ font-weight: bold; color: #666; font-size: 12px; }}
        .score-excellent {{ background-color: #c6efce; color: #006100; font-weight: bold; }}
        .score-good {{ background-color: #ffc7ce; color: #9c0006; font-weight: bold; }}
        .footer {{ margin-top: 20px; color: #999; font-size: 12px; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>QC Performance Evaluation Report</h1>
        <div class="info">Report Generated: {timestamp} | Report Period: 2026-01 ~ 2026-06</div>

        <h2>QC Performance Evaluation (滿分100分)</h2>
        <table>
            <tr>
                <th>QC Name</th>
                <th>Region</th>
                <th>A. Work Volume (10%)<br/><span class="metric">Sample Qty + PO</span></th>
                <th>B. Execution Rate (30%)<br/><span class="metric">Sample%</span></th>
                <th>C. Daily Efficiency (30%)<br/><span class="metric">Sample/Day</span></th>
                <th>D. Discrepancy (30%)<br/><span class="metric">Fail+C%+Point</span></th>
                <th>Overall<br/>Score</th>
            </tr>
"""

    for idx, row in kpi_df.iterrows():
        overall_score = row['Overall_Score']
        score_class = 'score-excellent' if overall_score >= 75 else 'score-good'

        html += f"""            <tr>
                <td><strong>{row['QC']}</strong></td>
                <td>{row['Location']}</td>
                <td>{row['A_SampleQty']:.0f}碼 / {row['A_POCount']:.0f}筆</td>
                <td>{row['B_ExecutionRate']:.1f}%</td>
                <td>{row['C_DailyEfficiency']:.0f}</td>
                <td>{row['D_DiscrepancyCount']:.0f}筆</td>
                <td class="{score_class}">{overall_score:.1f}</td>
            </tr>
"""

    html += """        </table>
        <div class="footer">
            <p>Note: This report compares QC inspection records with factory records.</p>
            <p>Scores are calculated based on: Volume (A), Sample rate (B), Daily efficiency (C), and Accuracy match with factory (D).</p>
        </div>
    </div>
</body>
</html>
"""

    return html

def main(year_month=None):
    print("="*60)
    print("QC Performance KPI Analysis System")
    print("="*60)

    # Ask for statistics period if not provided
    if year_month is None:
        year_month = input("\n請輸入統計期間 (例如: 26-1~26-6): ").strip()
        if not year_month:
            year_month = '26-1~26-6'  # Default

    print(f"\n統計期間: {year_month}")

    # Load data
    all_df = load_data(year_month)

    # Separate QC and Factory records
    qc_records, factory_records = separate_qc_factory(all_df)

    # Calculate KPI
    print("\nCalculating KPI metrics...")
    kpi_df = calculate_qc_kpi(qc_records, factory_records)

    # Calculate scores
    print("Calculating scores...")
    kpi_df = calculate_scores(kpi_df)

    print(f"\nKPI Summary ({len(kpi_df)} QC staff):")
    print(kpi_df[['QC', 'Location', 'A_SampleQty', 'A_POCount', 'B_ExecutionRate', 'C_DailyEfficiency', 'D_DiscrepancyCount', 'Overall_Score']].to_string())

    # Generate HTML
    print("\nGenerating HTML report...")
    html = generate_html(kpi_df)

    # Save HTML to 14_qc performance
    os.makedirs(OUTPUT_DIR, exist_ok=True)
    output_file = os.path.join(OUTPUT_DIR, 'qc_kpi_report.html')
    with open(output_file, 'w', encoding='utf-8') as f:
        f.write(html)

    # Copy to 00_html file/14_qc performance
    try:
        import shutil
        backup_dir = r"E:\88. Claude\00_html file\14_qc performance"
        os.makedirs(backup_dir, exist_ok=True)
        backup_path = os.path.join(backup_dir, 'qc_kpi_report.html')
        shutil.copy2(output_file, backup_path)
        print(f"\nOK - Report saved:")
        print(f"  {output_file}")
        print(f"  {backup_path}")
    except Exception as e:
        print(f"Copy failed: {e}")

if __name__ == '__main__':
    import sys
    year_month = sys.argv[1] if len(sys.argv) > 1 else None
    main(year_month)
