import pandas as pd
import os
import random
from datetime import datetime, timedelta

# Configuration
INPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'input')
NUM_STUDENTS = 20
NUM_DAYS = 5

# Ensure input directory exists
os.makedirs(INPUT_DIR, exist_ok=True)

# Generate dummy students
students = []
for i in range(NUM_STUDENTS):
    students.append({
        'name': f'Student_{i+1}',
        'id': f'S{1000+i}'
    })

# Generate files for each day
start_date = datetime(2023, 10, 1)

for day in range(NUM_DAYS):
    current_date = start_date + timedelta(days=day)
    date_str = current_date.strftime('%Y-%m-%d')
    filename = f"{date_str}_comment_sheet.xlsx"
    filepath = os.path.join(INPUT_DIR, filename)

    # Randomly select students who submitted (simulate absentees)
    num_submissions = random.randint(15, NUM_STUDENTS)
    submitting_students = random.sample(students, num_submissions)

    data = []
    for s in submitting_students:
        # Create a row with empty columns A-D, and data in E, F, G
        # Pandas dataframe construction
        row = {
            'A': '', 'B': '', 'C': '', 'D': '',
            'E': s['name'],         # Full Name
            'F': s['id'],           # Student ID
            'G': f"Comment from {s['name']} on {date_str}" # Comment
        }
        data.append(row)

    df = pd.DataFrame(data)
    
    # Rename columns to map to Excel columns (A, B, C... work differently in pandas read/write, 
    # but for simplicity we will write with header=False to simulate the form structure if needed.
    # However, usually forms have headers. Let's assume row 1 is header.
    # The prompt says "E列 Full Name", so E is the 5th column.
    
    # Let's create a DataFrame with explicit column positions
    # Columns: A, B, C, D, Full Name, Student ID, Comment
    # Column indices: 0, 1, 2, 3, 4, 5, 6
    
    final_df = pd.DataFrame(columns=['ColA', 'ColB', 'ColC', 'ColD', 'フルネーム', 'Q00_学籍番号', 'Q01_コメントシート'])
    
    for i, s in enumerate(submitting_students):
        comment = f"Comment from {s['name']} on {date_str}"
        # We can simulate empty comment sometimes
        if random.random() < 0.1:
            comment = "" 
            
        final_df.loc[i] = ['', '', '', '', s['name'], s['id'], comment]

    # Save to Excel
    print(f"Generating {filepath} with {len(final_df)} comments.")
    final_df.to_excel(filepath, index=False)

print("Test data generation complete.")
