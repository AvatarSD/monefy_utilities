import sys
import pandas as pd
import re
import enum

if len(sys.argv) < 2: 
    print("One arg required")
    sys.exit(-1)


moneyfee_table = pd.read_excel(sys.argv[1], engine='openpyxl')
print("Shape of import", moneyfee_table.shape)

class dir_t(enum.Enum):
    DIRECT = 0
    INVERT = 1

transaction = 0
buy = 0
sell = 0

# 1. Add additional column "handled"
moneyfee_table['handled'] = False

# 2. Iterate over rows of dataframe
for i,row in moneyfee_table.iterrows():
    print(f'{i+1}\{len(moneyfee_table)}           \r')
    if row["handled"] == True: continue

    # 1.1 Check category
    category = row["category"]
    old_acc = row['account']
    amount = float(str(row['amount']).replace(" ","").replace("\xa0",""))
    currency = row['currency.1']

    # 1.2 Transaction
    categoty_match = re.match(r"(To|From) '(.*?)'", category)
    if categoty_match: 
        # 1.2.1 meet first "To \'%s\'" or "From\'%s\'"
        dir = dir_t.INVERT if categoty_match.group(0) == 'From' else dir_t.INVERT
        new_acc = categoty_match.group(1)
    
        # 1.2.2 iterate over unhandled
        for dest_i, dest_row in moneyfee_table.loc[i:].iterrows():
            if dest_row['handled'] is True: continue
            
            # 1.2.2.1 find "To \'%s\'" or "From\'%s\'" with s
            dest_category = dest_row["category"]
            dest_categoty_match = re.match(r"(To|From) '(.*?)'", dest_category)
            if not dest_categoty_match : continue
            
            dest_dir = dir_t.INVERT if dest_categoty_match.group(0) == 'From' else dir_t.INVERT
            dest_new_acc = dest_row['account']
            dest_old_acc = dest_categoty_match.group(1)
            dest_amount = float(str(dest_row['amount']).replace(" ","").replace("\xa0",""))

            # 1.2.2.2 same -amount value in "amount" field
            # 1.2.2.3 same account from 'category' column (%s in From or To) 
            if (dest_dir is dir or old_acc is not dest_old_acc or new_acc is not dest_new_acc or amount is not -dest_amount): 
                if len(dest_row)-1 is dest_i :
                    print(f"Error, transaction not founded: {dest_dir} is {dir} or {old_acc} is not {dest_old_acc} or {new_acc} is not {dest_new_acc} or {amount} is not {-dest_amount}\r\n")
                continue
            
            print(f"Row {i} + {dest_i} is transaction: dir:{str(dir)}, {dest_old_acc} --> {dest_new_acc}, amount: {amount} {currency}\r\n")
            
            # 1.2.3 mark dest handled
            dest_row['handled'] = True
            
            # 1.2.4 save transfer
            transaction = transaction + 1

            # 1.2.5 break, continue
            break
        
    # 1.3 Regular
    else:
        # 1.3.2 save transaction
        if amount > 0:
            sell = sell+1
        else:
            buy = buy+1

    # 1.4 mark handled
    row['handled'] = True
print('>>>>\r\n')

# 3. Check unhandled
ungandled = moneyfee_table.loc[(moneyfee_table['handled'] == False)]
# ungandled.to_excel("unhandled.xlsl")
print(ungandled)

# 4. Save new list in specific format
print(f'Solds: {sell}, Buys: {buy}, Transaction: {transaction}, Err: {len(ungandled)}')
print(f'>> Total: {sell+buy+transaction*2+len(ungandled)}/{len(moneyfee_table)}')

# print(moneyfee_table)
