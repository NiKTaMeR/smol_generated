import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from utils import get_month_range, get_unique_clients, get_client_monthly_revenue

def create_mrr_per_client_table(input_file, output_file):
    # Read input data
    df = pd.read_excel(input_file)

    # Get unique clients and month range
    unique_clients = get_unique_clients(df)
    month_range = get_month_range(df)

    # Create MRR Per Client DataFrame
    mrr_per_client_data = {'Logo': unique_clients, 'Cohort Month': [get_client_monthly_revenue(df, client).index.min().strftime('%m/%Y') for client in unique_clients]}
    mrr_per_client_df = pd.DataFrame(mrr_per_client_data)

    # Add monthly revenue columns
    for month in month_range:
        month_str = month.strftime('%m/%Y')
        mrr_per_client_df[month_str] = [get_client_monthly_revenue(df, client).get(month, 0) for client in unique_clients]

    # Write MRR Per Client DataFrame to output file
    with pd.ExcelWriter(output_file, engine='openpyxl', mode='a') as writer:
        mrr_per_client_df.to_excel(writer, index=False, sheet_name='MRR Per Client')