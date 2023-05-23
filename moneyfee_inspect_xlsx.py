import sys
import pandas as pd

if len(sys.argv) < 2: 
    sys.exit()

moneyfee_table = pd.read_excel(sys.argv[1], engine='openpyxl')

print("Shape of import", moneyfee_table.shape)
print(moneyfee_table)

for i,row in moneyfee_table.iterrows():
    print(i,"money +-:",row["amount"])