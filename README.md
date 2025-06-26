# QuickBooks Export Scripts

This repository contains Python scripts to export various types of data from QuickBooks Desktop to CSV format. These scripts use the QuickBooks SDK and QBXML to interact with QuickBooks data.

## Prerequisites

- QuickBooks Desktop installed on your Windows machine
- Python 3.x
- Required Python packages (install using `pip install -r requirements.txt`):
  - pywin32
  - pythoncom

## Installation

1. Clone this repository or download the script files
2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

## Available Scripts

### 1. qb_inv.py - Export Customer and Ship-To Addresses

**Purpose**: Exports all customer information along with their ship-to addresses to a CSV file.

**Usage**:
```
python qb_inv.py
```

**Output**:
- Creates a file named `shipto_addresses.csv` with the following columns:
  - Customer: Customer name
  - ShipToAddress: Formatted ship-to address

### 2. qb_so.py - Export Sales Orders

**Purpose**: Exports sales order data from QuickBooks to a CSV file. You can export by specific SO numbers or for an entire year.

**Usage**:
```
python qb_so.py
```

**Options**:
- When prompted, choose to:
  - (n) Fetch by specific SO numbers (comma-separated)
  - (y) Fetch all SOs for a specific year

**Output**:
- Creates a file named `sales_orders_export.csv` with detailed SO information including:
  - SO Number, Customer Name, Transaction Date, Due Date
  - Bill To and Ship To addresses
  - Line item details (Description, Quantity, Rate, Amount)
  - Totals (Subtotal, Sales Tax, Total Amount)

### 3. qb_po.py - Export Purchase Orders

**Purpose**: Exports purchase order data from QuickBooks to a CSV file. You can export by specific PO numbers or for an entire year.

**Usage**:
```
python qb_po.py
```

**Options**:
- When prompted, choose to:
  - (n) Fetch by specific PO numbers (comma-separated)
  - (y) Fetch all POs for a specific year

**Output**:
- Creates a file named `purchase_orders_export.csv` with detailed PO information including:
  - PO Number, Vendor Name, Transaction Date, Due Date
  - Vendor and Ship To addresses
  - Line item details (Description, Quantity, Rate, Amount)
  - Total Amount

## Important Notes

1. **QuickBooks Connection**:
   - QuickBooks must be open on your computer when running these scripts
   - The first time you run a script, QuickBooks will ask for permission to allow access

2. **Error Handling**:
   - If a script fails, check that QuickBooks is running and logged in
   - Ensure you have the necessary permissions in QuickBooks
   - Some scripts may time out if there's a large amount of data

3. **Output Files**:
   - CSV files are created in the same directory as the scripts
   - Existing files with the same name will be overwritten

## Troubleshooting

- **Permission Errors**: Make sure QuickBooks is running and you're logged in with appropriate permissions
- **Module Not Found**: Ensure all required Python packages are installed
- **Connection Issues**: Check that QuickBooks is not in multi-user mode if you encounter connection problems

## License

This project is provided as-is. Please ensure you have the right to access and export this data according to your QuickBooks license and company policies.
