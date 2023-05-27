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
    conversion = 0
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
        self.tags = str("")                # Теги
        self.comment = str("")             # Коментар		

    def to_str(self) -> str:
        return str("dat: ") + str(self.date) + str("\t\tamo: ") + str(self.amount) + str("\t\tamo_com: ") + str(self.amount_company) + str("\t\tatx: ") + str(self.account_tx) + str("\t\tarx: ") + str(self.account_rx) + str("\t\tcat: ") + str(self.categoty) + str("\t\tmsg: ") + str(self.comment) + str("\t\t")

    def row(self):
        return [self.date, 
                self.date_income, 
                self.date_begin, 
                self.date_end, 
                self.amount, 
                self.amount_company, 
                self.account_tx, 
                self.account_rx, 
                self.categoty, 
                self.categoty_sub, 
                self.agent, 
                self.project, 
                self.tags, 
                self.comment ]

# PARSE ARGS
if len(sys.argv) < 2: 
    print("One arg required")
    sys.exit(-1)


# 0. Init required vars
moneyfee_table = pd.read_excel(sys.argv[1], engine='openpyxl')
moneyfee_table.columns = ['date', 'account', 'category', 'amount', 'currency.1', "converted amount", "converted currency", "description"]

def_currency = str("UAH")
mm_table = pd.DataFrame(columns=['date',
                                 'date_income',
                                 'date_begin',
                                 'date_end',
                                 'amount',
                                 'amount_company',
                                 'account_tx',
                                 'account_rx',
                                 'categoty',
                                 'categoty_sub',
                                 'agent',
                                 'project',
                                 'tags',
                                 'comment'])

print(moneyfee_table)
print("Shape of import", moneyfee_table.shape)

# 1. Add additional column "handled"
moneyfee_table['handled'] = False

# 2. Iterate over rows of dataframe
for i,row in moneyfee_table.iterrows():
    print(f'{i+1}\{len(moneyfee_table)}           ',end='\r')

    if moneyfee_table.at[i, "handled"] == True: continue
    mm = money_move()

    # 1.1 Check category
    category = row["category"]
    old_acc = row['account']
    amount = float(str(row['converted amount']).replace(" ","").replace("\xa0",""))
    amount_orig = float(str(row['amount']).replace(" ","").replace("\xa0",""))
    currency = row['currency.1']
    dest_amount = 0.0

    # 1.2 transaction
    categoty_match = re.match(r"(To|From) '(.*?)'", category)
    if categoty_match: 
        # 1.2.1 meet first "To \'%s\'" or "From\'%s\'"
        dir = dir_t.DIRECT if categoty_match.group(1) == 'From' else dir_t.INVERT
        new_acc = categoty_match.group(2)
    
        # 1.2.2 iterate over unhandled
        for dest_i, dest_row in moneyfee_table.loc[i:].iterrows():
            if moneyfee_table.at[dest_i, 'handled'] == True: continue
            
            # 1.2.2.1 find "To \'%s\'" or "From\'%s\'" with s
            dest_category = dest_row["category"]
            dest_categoty_match = re.match(r"(To|From) '(.*?)'", dest_category)
            if not dest_categoty_match : continue

            dest_dir = dir_t.DIRECT if dest_categoty_match.group(1) == 'From' else dir_t.INVERT
            dest_new_acc = dest_row['account']
            dest_old_acc = dest_categoty_match.group(2)
            dest_amount = float(str(dest_row['converted amount']).replace(" ","").replace("\xa0",""))
            dest_amount_orig = float(str(dest_row['amount']).replace(" ","").replace("\xa0",""))
            dest_currency = dest_row['currency.1']
            if dir == dir_t.INVERT: currency = dest_currency

            # 1.2.2.2 same -amount value in "amount" field
            # 1.2.2.3 same account from 'category' column (%s in From or To) 
            if ((dest_dir.value == dir.value) or (old_acc != dest_old_acc) or (new_acc != dest_new_acc) or (amount != -dest_amount)): 
                if len(dest_row)-1 == dest_i :
                    print(f"Error, transaction not founded: {dest_dir} is {dir} or {old_acc} is not {dest_old_acc} or {new_acc} is not {dest_new_acc} or {amount} is not {-dest_amount}")
                continue
            
            print(f"Row {i} + {dest_i} is transaction: dir:{str(dir)}, {dest_old_acc} --> {dest_new_acc}, amount: {abs(amount) if currency == def_currency else abs(dest_amount_orig)} {currency}")

            # 1.2.3 mark dest handled
            moneyfee_table.at[i, 'handled'] = True
            moneyfee_table.at[dest_i, 'handled'] = True
            if currency == def_currency: stats.transaction = stats.transaction + 1
            else: stats.conversion = stats.conversion + 1
            
            # 1.2.4 save transfer
            mm.date = row["date"]
            mm.amount = abs(amount) if currency == def_currency else abs(dest_amount_orig)
            mm.amount_company = str("") if currency == def_currency else abs(amount)            
            mm.account_rx = old_acc if dir == dir_t.DIRECT else new_acc
            mm.account_tx = old_acc if dir == dir_t.INVERT else new_acc
            mm.categoty = str("transfer") if currency == def_currency else str("exchange")
            mm.comment = old_acc + str(" ") + new_acc
            mm_table.loc[len(mm_table)] = mm.row()

            # 1.2.5 break, continue
            break
        if(amount != -dest_amount):
            print(f"Amount: {amount}, company amount: {dest_amount}")
                
        
    # 1.3 Regular
    else:
        # 1.3.1 save transfer
        dir = dir_t.INVERT if amount > 0 else dir_t.DIRECT
        mm.date = row["date"]
        mm.amount = abs(amount) if currency == def_currency else abs(amount_orig)
        mm.amount_company = str("") if currency == def_currency else abs(amount) 
        mm.account_tx = old_acc if dir == dir_t.DIRECT else str("")
        mm.account_rx = old_acc if dir == dir_t.INVERT else str("")
        mm.categoty = category
        dsc = row["description"]
        mm.comment = str("") if ((dsc is None) or (str("nan") == str(dsc))) else row["description"]
        mm_table.loc[len(mm_table)] = mm.row()

        # 1.3.2 save stats transaction
        if dir == dir_t.INVERT:
            stats.sell = stats.sell+1
        else:
            stats.buy = stats.buy+1

        # 1.3.3 mark handled
        moneyfee_table.at[i, 'handled'] = True
print('>>>>\r\n')

# 3. Check unhandled; Print inspection stat
ungandled = moneyfee_table.loc[(moneyfee_table['handled'] == False)]
print(ungandled)
# if len(ungandled): ungandled.to_excel("unhandled.xlsl")
print(f'Solds: {stats.sell}, stats.buys: {stats.buy}, stats.transaction: {stats.transaction}, conv: {stats.conversion}, Err: {len(ungandled)}')
print(f'>> Total: {stats.sell+stats.buy+stats.transaction*2+stats.conversion*2+len(ungandled)}/{len(moneyfee_table)}')

# 4. Save new list in specific format
mm_table.to_excel("unhandled.xlsx")
# for mm in mm_table:
#     print(mm.to_str())
