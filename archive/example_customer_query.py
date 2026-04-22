"""
Example script demonstrating how to use the customer_query tool.

This script shows various ways to query QuickBooks Desktop customers
and export the data to pandas DataFrame or CSV.
"""

import sys
import pandas as pd
from qb_mcp_tools.tools.customer_query import (
    customer_query,
    customer_search,
    customer_query_to_dataframe
)


def example_1_query_all_active():
    """Query all active customers."""
    print("Example 1: Query all active customers")
    print("=" * 60)

    customers = customer_query(active_status="ActiveOnly", max_returned=100)
    print(f"Found {len(customers)} active customers")

    # Convert to DataFrame
    df = pd.DataFrame(customers)
    print(f"\nDataFrame shape: {df.shape}")
    print(f"Columns: {df.columns.tolist()}")

    # Show sample data
    if not df.empty:
        print("\nSample customer data:")
        print(df[['FullName', 'Balance', 'Email', 'Phone']].head())

    return customers


def example_2_query_by_list_ids():
    """Query specific customers by ListID."""
    print("\n\nExample 2: Query customers by ListID")
    print("=" * 60)

    # Replace these with actual ListIDs from your QuickBooks
    list_ids = ["80000001-1234567890", "80000002-1234567890"]

    customers = customer_query(list_ids=list_ids)
    print(f"Found {len(customers)} customers")

    for customer in customers:
        print(f"  - {customer['FullName']} (ID: {customer['ListID']})")

    return customers


def example_3_query_by_names():
    """Query customers by full name."""
    print("\n\nExample 3: Query customers by FullName")
    print("=" * 60)

    # Replace with actual customer names from your QuickBooks
    full_names = ["Acme Corporation", "Smith:John Project"]

    customers = customer_query(full_names=full_names)
    print(f"Found {len(customers)} customers")

    for customer in customers:
        print(f"  - {customer['FullName']}")
        print(f"    Balance: ${customer['Balance']}")
        print(f"    Email: {customer['Email']}")

    return customers


def example_4_search_by_name():
    """Search customers by name (contains)."""
    print("\n\nExample 4: Search customers by name (contains 'corp')")
    print("=" * 60)

    results = customer_search("corp", match_type="contains")
    print(f"Found {len(results)} customers matching 'corp'")

    for customer in results[:5]:  # Show first 5
        print(f"  - {customer['FullName']}")

    return results


def example_5_fuzzy_search():
    """Fuzzy search for customers."""
    print("\n\nExample 5: Fuzzy search for 'acme corp'")
    print("=" * 60)

    results = customer_search("acme corp", match_type="fuzzy")
    print(f"Found {len(results)} customers")

    for customer in results[:5]:  # Show first 5
        score = customer.get('match_score', 0)
        print(f"  - {customer['FullName']} (score: {score:.2f})")

    return results


def example_6_export_to_csv():
    """Export customer data to CSV."""
    print("\n\nExample 6: Export customers to CSV")
    print("=" * 60)

    # Query active customers
    df = customer_query_to_dataframe(active_status="ActiveOnly")

    # Flatten some nested fields for CSV export
    # BillAddress is a dict, so we'll extract key fields
    if not df.empty and 'BillAddress' in df.columns:
        # Extract address components
        df['BillAddress_City'] = df['BillAddress'].apply(
            lambda x: x.get('City', '') if isinstance(x, dict) else ''
        )
        df['BillAddress_State'] = df['BillAddress'].apply(
            lambda x: x.get('State', '') if isinstance(x, dict) else ''
        )
        df['BillAddress_PostalCode'] = df['BillAddress'].apply(
            lambda x: x.get('PostalCode', '') if isinstance(x, dict) else ''
        )

    # Select columns for export
    export_columns = [
        'ListID', 'FullName', 'CompanyName', 'FirstName', 'LastName',
        'Email', 'Phone', 'Balance', 'TotalBalance',
        'BillAddress_City', 'BillAddress_State', 'BillAddress_PostalCode',
        'IsActive', 'TimeCreated'
    ]

    # Filter to only existing columns
    export_columns = [col for col in export_columns if col in df.columns]

    # Export to CSV
    output_file = "customers_export.csv"
    df[export_columns].to_csv(output_file, index=False)
    print(f"Exported {len(df)} customers to {output_file}")
    print(f"Columns: {export_columns}")


def example_7_customer_with_addresses():
    """Show customer with all address details."""
    print("\n\nExample 7: Customer with address details")
    print("=" * 60)

    # Query first customer
    customers = customer_query(max_returned=1)

    if customers:
        customer = customers[0]
        print(f"Customer: {customer['FullName']}")
        print(f"\nBill Address:")
        for key, value in customer['BillAddress'].items():
            if value:
                print(f"  {key}: {value}")

        print(f"\nShip Address:")
        for key, value in customer['ShipAddress'].items():
            if value:
                print(f"  {key}: {value}")

        print(f"\nShip-To Addresses ({len(customer['ShipToAddress'])}):")
        for i, ship_to in enumerate(customer['ShipToAddress'], 1):
            print(f"  {i}. {ship_to.get('Name', 'Unnamed')}")
            if ship_to.get('City'):
                print(f"     {ship_to['City']}, {ship_to.get('State', '')}")


def example_8_customer_references():
    """Show customer reference fields."""
    print("\n\nExample 8: Customer reference fields")
    print("=" * 60)

    customers = customer_query(max_returned=5)

    for customer in customers[:3]:  # Show first 3
        print(f"\nCustomer: {customer['FullName']}")
        print(f"  Customer Type: {customer['CustomerTypeRef']['FullName']}")
        print(f"  Terms: {customer['TermsRef']['FullName']}")
        print(f"  Sales Rep: {customer['SalesRepRef']['FullName']}")
        print(f"  Price Level: {customer['PriceLevelRef']['FullName']}")


def main():
    """Run all examples."""
    print("QuickBooks Customer Query Examples")
    print("=" * 60)
    print("Note: These examples require QuickBooks Desktop to be running")
    print("      and a company file to be open.\n")

    try:
        # Run examples
        example_1_query_all_active()
        # example_2_query_by_list_ids()  # Commented - needs real ListIDs
        # example_3_query_by_names()     # Commented - needs real names
        # example_4_search_by_name()
        # example_5_fuzzy_search()
        # example_6_export_to_csv()
        # example_7_customer_with_addresses()
        # example_8_customer_references()

        print("\n" + "=" * 60)
        print("Examples completed successfully!")

    except Exception as e:
        print(f"\nError: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
