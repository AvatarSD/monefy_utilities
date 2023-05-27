import sys
import pandas as pd
import re
import enum
from datetime import datetime as dt


# Types
class dir_t(enum.Enum):
    DIRECT = 0
    INVERT = 1

class stats:
    transaction = 0
    buy = 0
    sell = 0

class money_move:
    def __init__(self):
        self.date = dt.fromtimestamp(0)    # Дата платежу
        self.date_income = str("")         # Дата нарахування
        self.date_begin = str("")          # Період нарахування (дата "з")
        self.date_end = str("")            # Період нарахування (дата "по")
        self.amount = 0                    # Сума в валюті рахунка*
        self.amount_company = str("")      # Сума в валюті компанії
        self.account_tx = str("")          # З рахунку - для оплати
        self.account_rx = str("")          # На рахунок - для надходжень 
        self.categoty = str("")            # Категорія
        self.categoty_sub = str("")        # Підкатегорія
        self.agent = str("")               # Контрагент
        self.project = str("")             # Проект
        self.tags = []                     # Теги
        self.comment = str("")             # Коментар		

    def to_str(self) -> str:
        return str("dat: ") + str(self.date) + str("\t\tamo: ") + str(self.amount) + str("\t\tamo_com: ") + str(self.amount_company) + str("\t\tatx: ") + str(self.account_tx) + str("\t\tarx: ") + str(self.account_rx) + str("\t\tcat: ") + str(self.categoty) + str("\t\tmsg: ") + str(self.comment) + str("\t\t")


# PARSE ARGS
if len(sys.argv) < 2: 
    print("One arg required")
    sys.exit(-1)


# 0. Init required vars
moneyfee_table = pd.read_excel(sys.argv[1], engine='openpyxl')
def_currency = str("UAH")
mm_table = []

print(moneyfee_table)
print("Shape of import", moneyfee_table.shape)

# 1. Add additional column "handled"
moneyfee_table['handled'] = False

# 2. Iterate over rows of dataframe
for i,row in moneyfee_table.iterrows():
    print(f'{i+1}\{len(moneyfee_table)}           ',end='\r')
    if row["handled"] == True: continue
    mm = money_move()

    # 1.1 Check category
    category = row["category"]
    old_acc = row['account']
    amount = float(str(row['amount']).replace(" ","").replace("\xa0",""))
    amount_conv = float(str(row['converted amount']).replace(" ","").replace("\xa0",""))
    currency = row['currency.1']

    # 1.2 transaction
    categoty_match = re.match(r"(To|From) '(.*?)'", category)
    if categoty_match: 
        # 1.2.1 meet first "To \'%s\'" or "From\'%s\'"
        dir = dir_t.DIRECT if categoty_match.group(1) == 'From' else dir_t.INVERT
        new_acc = categoty_match.group(2)
    
        # 1.2.2 iterate over unhandled
        for dest_i, dest_row in moneyfee_table.loc[i:].iterrows():
            if dest_row['handled'] is True: continue
            
            # 1.2.2.1 find "To \'%s\'" or "From\'%s\'" with s
            dest_category = dest_row["category"]
            dest_categoty_match = re.match(r"(To|From) '(.*?)'", dest_category)
            if not dest_categoty_match : continue

            dest_dir = dir_t.DIRECT if dest_categoty_match.group(1) == 'From' else dir_t.INVERT
            dest_new_acc = dest_row['account']
            dest_old_acc = dest_categoty_match.group(2)
            dest_amount = float(str(dest_row['amount']).replace(" ","").replace("\xa0",""))

            # 1.2.2.2 same -amount value in "amount" field
            # 1.2.2.3 same account from 'category' column (%s in From or To) 
            if ((dest_dir.value == dir.value) or (old_acc != dest_old_acc) or (new_acc != dest_new_acc) or (amount != -dest_amount)): 
                if len(dest_row)-1 == dest_i :
                    print(f"Error, transaction not founded: {dest_dir} is {dir} or {old_acc} is not {dest_old_acc} or {new_acc} is not {dest_new_acc} or {amount} is not {-dest_amount}")
                continue
            
            print(f"Row {i} + {dest_i} is transaction: dir:{str(dir)}, {dest_old_acc} --> {dest_new_acc}, amount: {amount} {currency}")
            
            # 1.2.3 mark dest handled
            dest_row['handled'] = True
            
            # 1.2.4 save transfer
            mm.date = row["date"]
            mm.amount = abs(amount)
            mm.account_rx = old_acc if dir == dir_t.DIRECT else new_acc
            mm.account_tx = old_acc if dir == dir_t.INVERT else new_acc
            mm.categoty = str("transfer")
            mm.comment = old_acc + str(" ") + new_acc
            mm.amount_company = str("") if currency == def_currency else abs(amount_conv)
            mm_table.append(mm)

            # 1.2.5 break, continue
            break
        
    # 1.3 Regular
    else:
        # 1.3.1 save transfer
        dir = dir_t.INVERT if amount > 0 else dir_t.DIRECT
        mm.date = row["date"]
        mm.amount = abs(amount)
        mm.amount_company = str("") if currency == def_currency else abs(amount_conv)
        mm.account_tx = old_acc if dir == dir_t.DIRECT else str("")
        mm.account_rx = old_acc if dir == dir_t.INVERT else str("")
        mm.categoty = category
        dsc = row["description"]
        mm.comment = str("") if ((dsc is None) or (str("nan") == str(dsc))) else row["description"]
        mm_table.append(mm)

        # 1.3.2 save stats transaction
        if dir == dir_t.INVERT:
            stats.sell = stats.sell+1
        else:
            stats.buy = stats.buy+1

    # 1.4 mark handled
    row['handled'] = True
print('>>>>\r\n')

# 3. Check unhandled; Print inspection stat
ungandled = moneyfee_table.loc[(moneyfee_table['handled'] == False)]
print(ungandled)
# ungandled.to_excel("unhandled.xlsl")
print(f'Solds: {stats.sell}, stats.buys: {stats.buy}, stats.transaction: {stats.transaction}, Err: {len(ungandled)}')
print(f'>> Total: {stats.sell+stats.buy+stats.transaction*2+len(ungandled)}/{len(moneyfee_table)}')

# 4. Save new list in specific format
for mm in mm_table:
    print(mm.to_str())
