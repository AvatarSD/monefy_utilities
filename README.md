# MoneFy to FinMap Convertation Script

This Python script is designed to process and analyze MoneFy financial transaction data in an Excel file. It reads in a dataset of transactions, processes each transaction, checks if it's a regular transaction or a conversion, and handles each case accordingly. It also keeps track of summary statistics and handles exceptions.

## Installation

Before running the script, you'll need to ensure you have the necessary Python packages installed. These include:

- sys
- pandas
- re
- enum
- datetime

You can install these with pip:

```bash
pip install pandas openpyxl
```

**Note:** This script uses the `openpyxl` engine to read Excel files, which you'll need to install separately with pip.

## Usage

Run the script from the command line and provide the path to the Excel file as a command-line argument:

```bash
python moneyfee_inspect_xlsx.py transactions.xlsx
```

## Data Format

The script expects the Excel file to contain the following columns:

- Date
- Account
- Category
- Amount
- Currency
- Converted Amount
- Converted Currency
- Description

The script will rename these columns internally for its processing purposes.

## Output

The script prints out the processed transactions and unhandled transactions, along with summary statistics including the number of buys, sells, conversions, and transactions. It also saves the processed transactions into a new Excel file, `unhandled.xlsx`.

## Customization

Certain parts of the script, such as specific filters applied to account names and rules for updating specific accounts, are commented out. These may need to be customized based on the specific requirements of the dataset you're working with.

## Script Overview

The script consists of several stages:

1. **Definition of classes and types:** The script first defines a few classes and enums to represent a financial transaction and the direction of a transaction.

2. **Argument checking:** The script checks for a command-line argument that specifies the path to the Excel file. If no argument is provided, the script terminates with an error message.

3. **Dataframe creation:** The script reads the Excel file into a Pandas DataFrame, renames the DataFrame columns, and initializes a new DataFrame to store processed transactions.

4. **Data processing:** The script processes each transaction in the DataFrame, handling regular transactions and conversions differently. It keeps track of statistics and handles exceptions.

5. **Output and summary statistics:** The script outputs unhandled transactions and summary statistics, and saves the processed transactions into an Excel file.

#### Note

Please note, there are certain parts in the script that are commented out, like specific filters applied to account names and rules for updating specific accounts. These parts of the script may need to be customized based on the specific requirements of the dataset you are working with.