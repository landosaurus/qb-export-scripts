# Customer Query Tool

Complete implementation of QuickBooks Desktop customer querying via QBXML.

## Location
`<repo-root>/qb_mcp_tools/tools/customer_query.py`

## Features

### Main Functions

1. **`customer_query()`** - Primary query function with full filtering support
   - Query by ListIDs, FullNames, or name filter
   - Support for active/inactive status filtering
   - Iterator-based retrieval for large result sets
   - Returns list of customer dictionaries (pandas-compatible)

2. **`customer_search()`** - Search with fuzzy matching
   - Three match types: "contains", "starts_with", "fuzzy"
   - Fuzzy search includes scoring and sorting by relevance
   - Searches both FullName and CompanyName fields

3. **`customer_query_to_dataframe()`** - Direct DataFrame export
   - Convenience wrapper that returns pandas DataFrame
   - All query filters supported

### Data Fields Returned

#### Basic Information
- ListID, Name, FullName
- TimeCreated, TimeModified, EditSequence
- IsActive, Sublevel
- CompanyName, FirstName, LastName, MiddleName
- Salutation, JobTitle

#### Addresses (as dictionaries)
- BillAddress (dict with Addr1-5, City, State, PostalCode, Country, Note)
- ShipAddress (dict with same structure)
- ShipToAddress (list of dicts, each with Name and DefaultShipTo fields)

#### Contact Information
- Phone, AltPhone, Fax
- Email, Cc
- Contact, AltContact

#### References (as dictionaries with ListID and FullName)
- ParentRef
- CustomerTypeRef
- TermsRef
- SalesRepRef
- SalesTaxCodeRef
- ItemSalesTaxRef
- PriceLevelRef
- PreferredPaymentMethodRef
- CurrencyRef

#### Financial Fields
- Balance (float)
- TotalBalance (float)
- CreditLimit (float)
- ResaleNumber
- TaxRegistrationNumber

#### Additional Fields
- AccountNumber
- Notes
- JobStatus, JobStartDate, JobProjectedEndDate, JobEndDate, JobDesc
- PreferredDeliveryMethod
- IsStatementWithParent

## Usage Examples

### Example 1: Query all active customers
```python
from qb_mcp_tools.tools.customer_query import customer_query
import pandas as pd

customers = customer_query(active_status="ActiveOnly")
df = pd.DataFrame(customers)
print(df[['FullName', 'Balance', 'Email']])
```

### Example 2: Query specific customers by ListID
```python
list_ids = ["80000001-1234567890", "80000002-1234567890"]
customers = customer_query(list_ids=list_ids)
```

### Example 3: Query by name
```python
full_names = ["Acme Corporation", "Smith:John Project"]
customers = customer_query(full_names=full_names)
```

### Example 4: Search by name filter
```python
from qb_mcp_tools.tools.customer_query import customer_search

# Contains search
results = customer_search("acme", match_type="contains")

# Fuzzy search with scoring
results = customer_search("acme corp", match_type="fuzzy")
for r in results[:5]:
    print(f"{r['FullName']} - Score: {r['match_score']:.2f}")
```

### Example 5: Export to DataFrame
```python
from qb_mcp_tools.tools.customer_query import customer_query_to_dataframe

df = customer_query_to_dataframe(active_status="ActiveOnly")
df.to_csv("customers.csv", index=False)
```

### Example 6: Access nested data
```python
customers = customer_query(max_returned=10)

for customer in customers:
    print(f"Customer: {customer['FullName']}")
    print(f"  City: {customer['BillAddress']['City']}")
    print(f"  Price Level: {customer['PriceLevelRef']['FullName']}")
    print(f"  Ship-To Addresses: {len(customer['ShipToAddress'])}")
```

### Example 7: Flatten nested data for CSV export
```python
import pandas as pd
from qb_mcp_tools.tools.customer_query import customer_query

customers = customer_query(active_status="ActiveOnly")
df = pd.DataFrame(customers)

# Extract address components
df['BillAddress_City'] = df['BillAddress'].apply(lambda x: x.get('City', ''))
df['BillAddress_State'] = df['BillAddress'].apply(lambda x: x.get('State', ''))

# Export selected columns
df[['FullName', 'Balance', 'BillAddress_City', 'BillAddress_State']].to_csv('customers.csv')
```

## Implementation Details

### QBXML Generation
- Uses `CustomerQueryRq` with iterator support for large datasets
- Properly escapes XML special characters
- Supports multiple filter types (ListID, FullName, NameFilter, ActiveStatus)
- MaxReturned configurable (default: 100 per batch)

### Parsing
- Uses `xml.etree.ElementTree` for XML parsing
- Comprehensive field extraction from `CustomerRet` elements
- Handles optional fields with proper defaults
- Converts nested address and reference data to dictionaries

### Iterator Support
- Automatically handles large result sets via QB's iterator mechanism
- Transparently batches requests and combines results
- No manual pagination required

### Error Handling
- Uses `QBConnection` context manager for safe COM resource cleanup
- Proper exception propagation

## Type Hints
Full type hints provided using:
- `pandas` (pd.DataFrame)
- `typing` (List, Dict, Optional, Any)

## Dependencies
- `xml.etree.ElementTree` (stdlib)
- `pandas`
- `typing` (stdlib)
- `..utils.qb_connection` (QBConnection, wrap_qbxml_request)
- `..utils.qb_helpers` (parse_address, parse_ref_element, parse_decimal, xml_escape, fuzzy_match_score)

## Testing
See `example_customer_query.py` for comprehensive usage examples.
