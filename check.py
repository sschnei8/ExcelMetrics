import duckdb
import os
from tabulate import tabulate # Import tabulate

# --- (Keep the connection and query execution part the same as Method 2) ---

database_path = ':memory:'
excel_file_path = '/Users/iamsam/Downloads/for_sam.xlsx' # Use your actual path

print(f"Connecting to database: {database_path}")
print(f"Reading Excel file: {excel_file_path}")

if not os.path.exists(excel_file_path):
     print(f"\nError: Excel file not found at '{excel_file_path}'")
     exit()

try:
    with duckdb.connect(database=database_path) as con:
        con.execute(
            "CREATE OR REPLACE TABLE XL_TBL AS SELECT * FROM read_xlsx(?, sheet = 'Sheet1', header = true)",
            [excel_file_path]
        )
        print("Table 'XL_TBL' created from Excel.")

        sql_query = """
            SELECT
                modality,
                "Question category",
                COUNT(*) AS count_per_group
            FROM XL_TBL
            GROUP BY ALL
            ORDER BY modality, "Question category";
        """
        print("\nExecuting query for counts per modality and category...")

        result = con.execute(sql_query)
        headers = [desc[0] for desc in result.description]
        data = result.fetchall()

        # --- Display using tabulate ---
        print("\n--- Count per modality, category (Tabulate Format) ---")
        if not data:
             print("(No results found)")
        else:
            # Simple usage: headers list and data list-of-lists/tuples
            print(tabulate(data, headers=headers, tablefmt="grid")) # "grid", "pipe", "simple", etc.

            # Example with different format:
            # print("\n--- Alternative Format ---")
            # print(tabulate(data, headers=headers, tablefmt="pipe"))

        print("------------------------------------------------------")


except duckdb.Error as e:
    print(f"\nAn error occurred with DuckDB: {e}")
except ImportError:
     print("\nError: tabulate library not found. Please install it: uv pip install tabulate")
except Exception as e:
    print(f"\nAn unexpected error occurred: {e}")

print("\nAnalysis complete.")