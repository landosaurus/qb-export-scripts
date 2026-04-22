# Estimate Query Tool Reference

## Location
`/Users/lang/Documents/coding_temp/qb-export-scripts/qb_mcp_tools/tools/estimate_query.py`

## Function Signature

```python
def estimate_query(
    txn_ids: Optional[List[str]] = None,
    ref_numbers: Optional[List[str]] = None,
    customer_name: Optional[str] = None,
    from_date: Optional[str] = None,
    to_date: Optional[str] = None,
    include_line_items: bool = True,
    max_returned: int = 100
) -> List[Dict[str, Any]]
```

## Parameters

- **txn_ids**: List of specific estimate TxnIDs to retrieve
- **ref_numbers**: List of specific estimate reference numbers to retrieve
- **customer_name**: Filter by customer name
- **from_date**: Start of transaction date range (YYYY-MM-DD format)
- **to_date**: End of transaction date range (YYYY-MM-DD format)
- **include_line_items**: Whether to include line item details (default: True)
- **max_returned**: Maximum number of records to return per batch (default: 100)

## Return Structure

Returns a list of estimate dictionaries. Each dictionary contains:

### Header Fields
```python
{
    'TxnID': str,              # Transaction ID
    'RefNumber': str,          # Estimate/quote number
    'TxnDate': str,            # Transaction date (YYYY-MM-DD)
    'CustomerRef': {           # Customer reference
        'ListID': str,
        'FullName': str
    },
    'BillAddress': {           # Billing address
        'Addr1': str,
        'Addr2': str,
        'Addr3': str,
        'Addr4': str,
        'Addr5': str,
        'City': str,
        'State': str,
        'PostalCode': str,
        'Country': str,
        'Note': str
    },
    'ShipAddress': {           # Shipping address (same structure as BillAddress)
        # ... same fields as BillAddress
    },
    'PONumber': str,           # Purchase order number
    'TermsRef': {              # Payment terms reference
        'ListID': str,
        'FullName': str
    },
    'DueDate': str,            # Due date (YYYY-MM-DD)
    'ExchangeRate': str,       # Exchange rate (for multi-currency)
    'Subtotal': str,           # Subtotal amount
    'SalesTaxTotal': str,      # Sales tax total
    'TotalAmount': str,        # Total amount
    'IsActive': str,           # Whether estimate is active (true/false)
    'LineItems': [...]         # List of line items (see below)
}
```

### Line Item Structure (when include_line_items=True)

Each line item in the `LineItems` list can be either a regular line or a group:

#### Regular Line Item
```python
{
    'TxnLineID': str,              # Transaction line ID
    'ItemRef': {                   # Item reference
        'ListID': str,
        'FullName': str
    },
    'Desc': str,                   # Line description
    'Quantity': str,               # Quantity
    'Rate': str,                   # Rate per unit
    'RatePercent': str,            # Rate as percentage
    'Amount': str,                 # Line total amount
    'ClassRef': {                  # Class reference (if applicable)
        'ListID': str,
        'FullName': str
    },
    'SalesTaxCodeRef': {           # Sales tax code reference
        'ListID': str,
        'FullName': str
    },
    'Markup': str,                 # Markup amount
    'MarkupRate': str,             # Markup rate
    'PriceLevelRef': {             # Price level reference
        'ListID': str,
        'FullName': str
    },
    'IsPrintItemsInGroup': str,    # Print items in group flag
    'Type': 'Line'                 # Type indicator
}
```

#### Group Line Item
```python
{
    'TxnLineID': str,              # Transaction line ID
    'ItemGroupRef': {              # Item group reference
        'ListID': str,
        'FullName': str
    },
    'Desc': str,                   # Group description
    'Quantity': str,               # Quantity
    'UnitOfMeasure': str,          # Unit of measure
    'TotalAmount': str,            # Total amount for group
    'IsPrintItemsInGroup': str,    # Print items flag
    'Type': 'Group',               # Type indicator
    'GroupLines': [...]            # List of nested line items (same structure as regular lines)
}
```

## Usage Examples

### Example 1: Query estimates by reference numbers
```python
from qb_mcp_tools.tools.estimate_query import estimate_query

estimates = estimate_query(
    ref_numbers=['EST-001', 'EST-002'],
    include_line_items=True
)

for est in estimates:
    print(f"Estimate {est['RefNumber']}: ${est['TotalAmount']}")
    for line in est['LineItems']:
        print(f"  - {line['Desc']}: {line['Quantity']} x ${line['Rate']}")
```

### Example 2: Query estimates by date range for specific customer
```python
estimates = estimate_query(
    customer_name='ABC Company',
    from_date='2024-01-01',
    to_date='2024-12-31',
    include_line_items=True
)
```

### Example 3: Query specific estimate by TxnID (no line items)
```python
estimates = estimate_query(
    txn_ids=['12345-67890'],
    include_line_items=False
)
```

### Example 4: Query all estimates from 2024
```python
estimates = estimate_query(
    from_date='2024-01-01',
    to_date='2024-12-31',
    include_line_items=True,
    max_returned=100  # Process in batches of 100
)
```

## QBXML Details

The tool generates QBXML version 16.0 requests with proper:
- EstimateQueryRq structure
- Iterator support for large result sets
- Flexible filter combinations (TxnID, RefNumber, date ranges, customer filter)
- Optional line item inclusion
- All relevant estimate return elements

## Dependencies

- `QBConnection`: From `..utils.qb_connection`
- `wrap_qbxml_request`: From `..utils.qb_connection`
- `parse_address`, `format_address`, `parse_ref_element`, `xml_escape`, `build_txn_filter`: From `..utils.qb_helpers`

## Type Hints

The module uses proper type hints:
```python
from typing import List, Dict, Optional, Any
```

All parameters and return types are properly annotated for IDE support and type checking.
