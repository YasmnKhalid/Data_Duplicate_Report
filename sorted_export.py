# Ensure pandas and SQLAlchemy are installed in your environment
# pip install pandas sqlalchemy psycopg2-binary openpyxl
import pandas as pd
from sqlalchemy import create_engine

# Database connection details
db_user = '_YOUR_DB_USER_'  # Replace with your actual DB user
db_pass = '_YOUR_DB_PASS_'  # Replace with your actual DB password
db_host = '_YOUR_DB_HOST_'  # Replace with your actual DB host
db_port = '_YOUR_DB_PORT_'  # Replace with your actual DB port
db_name = '_YOUR_DB_NAME_'  # Replace with your actual DB name

# SQLAlchemy engine
engine = create_engine(f'postgresql+psycopg2://{db_user}:{db_pass}@{db_host}:{db_port}/{db_name}')

# Define tables and key columns (update these as per business logic)
table_keys = {
    'Table_name': ['key_column1', 'key_column2'],  # Replace with actual table and key columns
    'Another_table': ['another_key_column1', 'another_key_column2'],  # Add more tables as needed   
}

schema = 'public'  # Replace with your actual schema if needed


# Excel writer to store all duplicate rows and summary
with pd.ExcelWriter('duplicate_sorted_report.xlsx') as writer:
    dq_summary = []

    for table, key_columns in table_keys.items():
        print(f"🔍 Checking table: {table} by keys {key_columns}")
        full_table = f'"{schema}"."{table}"'
        try:
            # Read up to 10,000 rows (adjust as needed)
            df = pd.read_sql(f'SELECT * FROM {full_table} LIMIT 10000', engine)
            num_duplicates = df.duplicated(subset=key_columns).sum()

            # Extract all duplicate rows (keep=False includes all appearances)
            duplicates_df = df[df.duplicated(subset=key_columns, keep=False)]

            # Sort duplicates by key columns
            if not duplicates_df.empty:
                duplicates_df = duplicates_df.sort_values(by=key_columns)
                # Save to Excel sheet (Excel sheet name max 31 chars)
                duplicates_df.to_excel(writer, sheet_name=table[:31], index=False)

            dq_summary.append({
                'table': table,
                'num_rows_checked': len(df),
                'key_columns': ', '.join(key_columns),
                'num_duplicates': num_duplicates
            })

        except Exception as e:
            print(f"❌ Error processing {table}: {e}")
            dq_summary.append({
                'table': table,
                'num_rows_checked': 'ERROR',
                'key_columns': ', '.join(key_columns),
                'num_duplicates': 'ERROR',
                'error_message': str(e)
            })

    # Save summary sheet
    summary_df = pd.DataFrame(dq_summary)
    summary_df.to_excel(writer, sheet_name='Summary', index=False)

print("✅ Duplicate sorted report saved as duplicate_sorted_report.xlsx")
