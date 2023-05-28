import pandas as pd
from openpyxl import load_workbook
from config import threshold_contraction, threshold_upsell
from utils import get_month_range, get_monthly_revenue

def create_mrr_status_table(input_file, output_file):
    # Load input data
    df = pd.read_excel(input_file, sheet_name="MRR Per Client", index_col=0)

    # Initialize MRR Status DataFrame
    mrr_status_df = pd.DataFrame(index=df.index)
    month_range = get_month_range(df)

    # Calculate MRR Status
    for i in range(1, len(month_range)):
        pm = month_range[i - 1]
        cm = month_range[i]

        pm_revenue = df[pm]
        cm_revenue = df[cm]

        mrr_status = []

        for logo in df.index:
            pmba = pm_revenue.loc[logo]
            cmba = cm_revenue.loc[logo]

            if pmba == 0 and cmba > 0:
                status = "New Logo"
            elif pmba > 0 and cmba == 0:
                status = "Churned"
            elif pmba == cmba or (pmba > cmba and pmba * (1 - threshold_contraction) <= cmba) or (cmba > pmba and cmba * (1 - threshold_upsell) <= pmba):
                status = "Stable"
            elif cmba > pmba and cmba * (1 - threshold_upsell) > pmba:
                status = "Upsell"
            elif pmba > cmba and pmba * (1 - threshold_contraction) > cmba:
                status = "Contraction"
            else:
                status = "Unknown"

            mrr_status.append(status)

        mrr_status_df[cm] = mrr_status

    # Save MRR Status DataFrame to output file
    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a') as writer:
        mrr_status_df.to_excel(writer, sheet_name="MRR Status")