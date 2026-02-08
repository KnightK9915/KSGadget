import pandas as pd
import os

INPUT_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'input')
os.makedirs(INPUT_DIR, exist_ok=True)

# Test Case: Inconsistent IDs for "Sato Jo"
# File 1: 2023-10-01 (ID: 25BB0214)
# File 2: 2023-10-02 (ID: 25BB0215) - Typo in ID
# We expect ONE row for "Sato Jo" in the output, and the ID should probably be 25BB0215 (latest)

data1 = pd.DataFrame({
    'ColA': ['', ''], 'ColB': ['', ''], 'ColC': ['', ''], 'ColD': ['', ''],
    'Name': ['Sato Jo', 'Tanaka'],
    'ID': ['25BB0214', '111111'],
    'Comment': ['Comment 1', 'Comment T1']
})
data1.to_excel(os.path.join(INPUT_DIR, '2023-10-01_submit.xlsx'), header=True, index=False)

data2 = pd.DataFrame({
    'ColA': ['', ''], 'ColB': ['', ''], 'ColC': ['', ''], 'ColD': ['', ''],
    'Name': ['Sato Jo', 'Tanaka'],
    'ID': ['25BB0215', '111111'], # CHANGED ID for Sato Jo
    'Comment': ['Comment 2', 'Comment T2']
})
data2.to_excel(os.path.join(INPUT_DIR, '2023-10-02_submit.xlsx'), header=True, index=False)

# Add a student who only submitted on day 1 (to test "未回答")
data3 = pd.DataFrame({
    'ColA': [''], 'ColB': [''], 'ColC': [''], 'ColD': [''],
    'Name': ['Suzuki'],
    'ID': ['333333'],
    'Comment': ['Comment S1']
})
# Append to day 1 file
with pd.ExcelWriter(os.path.join(INPUT_DIR, '2023-10-01_submit.xlsx'), mode='a', if_sheet_exists='overlay') as writer:
    # This is complex to append row without messing up.
    # Easire to just rewrite day 1 with 3 students.
    pass

# Rewrite Day 1 with 3 students
data1_v2 = pd.DataFrame({
    'ColA': ['', '', ''], 'ColB': ['', '', ''], 'ColC': ['', '', ''], 'ColD': ['', '', ''],
    'Name': ['Sato Jo', 'Tanaka', 'Suzuki'],
    'ID': ['25BB0214', '111111', '333333'],
    'Comment': ['Comment 1', 'Comment T1', 'Comment S1']
})
data1_v2.to_excel(os.path.join(INPUT_DIR, '2023-10-01_submit.xlsx'), header=True, index=False)

print("Test data created.")
