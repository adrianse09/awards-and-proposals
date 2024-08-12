import pandas as pd

# Define the file path
file_path = r"C:\Users\abuitronboada\Box\CombineNode\Jupyter\combined.xlsx"

# Load the Excel file
df_combined = pd.read_excel(file_path)

# Display the first few rows of the dataframe to understand its structure
print(df_combined.head())
import pandas as pd

# Define the file path
file_path = r"C:\Users\abuitronboada\Box\CombineNode\Jupyter\combined.xlsx"

# Load each sheet into a dataframe
df_awards = pd.read_excel(file_path, sheet_name='Awards')  # Adjust 'Sheet1' to the actual sheet name for awards
df_proposals = pd.read_excel(file_path, sheet_name='Proposals')  # Adjust 'Sheet2' to the actual sheet name for proposals

# Display the first few rows of each dataframe to confirm correct loading
print("Awards Data:")
print(df_awards.head())
print("\nProposals Data:")
print(df_proposals.head())

from sqlalchemy import create_engine

# Create an SQLAlchemy engine instance
engine = create_engine('sqlite:///project_data.db')  # This creates a SQLite database named 'project_data.db' in your current directory

# Import data into the database
df_awards.to_sql('awards', con=engine, if_exists='replace', index=False)
df_proposals.to_sql('proposals', con=engine, if_exists='replace', index=False)

from sqlalchemy import text  # Import the text function

# Example of querying the database
with engine.connect() as connection:
    # Define the query as a text object
    query = text("SELECT * FROM awards LIMIT 5;")
    result = connection.execute(query)
    for row in result:
        print(row)

from sqlalchemy import text

# Adjusted SQL query for Awards
query_awards = text("""
SELECT SUM("Award Amount") as Total_Awarded, COUNT("Project Number") as Total_Awards
FROM awards
WHERE "Award Notice Date" >= '2023-07-01'
""")

# Adjusted SQL query for Proposals
query_proposals = text("""
SELECT SUM("Total Amount Proposed") as Total_Proposed, COUNT("Proposal Number") as Total_Proposals
FROM proposals
WHERE "Submitted to Sponsor" >= '2023-07-01'
""")

# Execute the queries
with engine.connect() as connection:
    awards_data = connection.execute(query_awards).fetchone()
    proposals_data = connection.execute(query_proposals).fetchone()

print("Awards Data:", awards_data)
print("Proposals Data:", proposals_data)

import matplotlib.pyplot as plt

# Data preparation
categories = ['Total Awarded', 'Total Awards', 'Total Proposed', 'Total Proposals']
values = [awards_data[0], awards_data[1], proposals_data[0], proposals_data[1]]

# Create a bar chart
plt.figure(figsize=(10, 6))
plt.bar(categories, values, color=['blue', 'green', 'red', 'purple'])
plt.title('Cumulative Data from July 1')
plt.xlabel('Categories')
plt.ylabel('Values')
plt.show()
