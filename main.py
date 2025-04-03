import duckdb

# Connect to an in-memory DuckDB database
con = duckdb.connect(database=':memory:')

# Path to your Excel file
excel_file_path = '/Users/iamsam/Downloads/for_sam.xlsx'  



# Create a table from an Excel file
con.execute(f"CREATE TABLE XL_TBL AS SELECT * FROM read_xlsx('{excel_file_path}', sheet = 'Sheet1')");

#select_all = con.execute(f"DESCRIBE XL_TBL").fetchall()
#print(select_all)

# Count per modality, category 
mod_cat = con.execute(f"""SELECT modality, "Question category", COUNT(*) FROM XL_TBL GROUP BY modality, "Question category" """).fetchall()
print(f"Count per modality, category: {mod_cat}")

# Count per modality, category 
all_out = con.execute(f"""SELECT * FROM XL_TBL LIMIT 2""").fetchall()
print(all_out)


# for the modality, how many in each category:
#1. Manikin-based simulator
#2. Part-task trainer
#3. VR, AR, MR or screen-based simulator
#4. Cadaver or live tissue
#how many papers asked questions in the different categories:
#1. Realism
#2. Educational Value
#3. Usability
#4. General evaluation
#average number of questions per paper
#average number of participants per paper, and average number of participants for each modality category
#question style:
#1. rating scale
#2. rating scale with open-ended
#3. open-ended
#4. interview
