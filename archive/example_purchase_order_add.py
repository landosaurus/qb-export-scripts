"""
Example usage of the purchase_order_add tool

This script demonstrates how to create purchase orders in QuickBooks Desktop
using the purchase_order_add function.
"""

from qb_mcp_tools.tools.purchase_order_add import purchase_order_add


def example_simple_purchase_order():
    """Create a simple purchase order with basic information."""
    print("Example 1: Simple Purchase Order")
    print("-" * 50)

    result = purchase_order_add(
        vendor_name="ABC Supplies",
        line_items=[
            {
                "item_name": "Widget A",
                "desc": "Blue widgets",
                "quantity": 10,
                "rate": 25.00
            },
            {
                "item_name": "Widget B",
                "quantity": 5,
                "rate": 50.00
            }
        ],
        txn_date="2024-01-15",
        ref_number="PO-2024-001",
        memo="Urgent order for Q1 inventory",
        expected_date="2024-01-30"
    )

    if result["success"]:
        print(f"✓ Successfully created purchase order")
        print(f"  PO Number: {result['RefNumber']}")
        print(f"  Transaction ID: {result['TxnID']}")
        print(f"  Total Amount: ${result['TotalAmount']:.2f}")
        print(f"  Created: {result['TimeCreated']}")
    else:
        print(f"✗ Error: {result['error_message']}")
        print(f"  Status Code: {result['status_code']}")

    print()


def example_dropship_purchase_order():
    """Create a drop ship purchase order with customer and shipping info."""
    print("Example 2: Drop Ship Purchase Order")
    print("-" * 50)

    result = purchase_order_add(
        vendor_name="Supplier Inc",
        line_items=[
            {
                "item_name": "Product X",
                "desc": "Special order for customer",
                "quantity": 20,
                "rate": 15.00,
                "customer_name": "Customer ABC"  # Drop ship to this customer
            }
        ],
        ship_address={
            "Addr1": "123 Customer St",
            "City": "Los Angeles",
            "State": "CA",
            "PostalCode": "90001"
        },
        ship_method="FedEx Ground",
        expected_date="2024-02-15",
        memo="Drop ship order for Customer ABC"
    )

    if result["success"]:
        print(f"✓ Successfully created drop ship purchase order")
        print(f"  PO Number: {result['RefNumber']}")
        print(f"  Transaction ID: {result['TxnID']}")
        print(f"  Total Amount: ${result['TotalAmount']:.2f}")
        print(f"  Vendor: {result['VendorRef']['FullName']}")
    else:
        print(f"✗ Error: {result['error_message']}")

    print()


def example_purchase_order_with_classes():
    """Create a purchase order with class tracking."""
    print("Example 3: Purchase Order with Class Tracking")
    print("-" * 50)

    result = purchase_order_add(
        vendor_name="Office Depot",
        line_items=[
            {
                "item_name": "Office Chair",
                "quantity": 3,
                "rate": 150.00,
                "class_name": "Marketing"
            },
            {
                "item_name": "Desk Lamp",
                "quantity": 5,
                "rate": 45.00,
                "class_name": "Sales"
            }
        ],
        txn_date="2024-01-20",
        terms="Net 30",
        due_date="2024-02-20",
        memo="Office furniture for new employees",
        is_to_be_printed=True
    )

    if result["success"]:
        print(f"✓ Successfully created purchase order with class tracking")
        print(f"  PO Number: {result['RefNumber']}")
        print(f"  Total Amount: ${result['TotalAmount']:.2f}")
        print(f"  Due Date: {result.get('DueDate', 'N/A')}")
    else:
        print(f"✗ Error: {result['error_message']}")

    print()


def example_purchase_order_with_addresses():
    """Create a purchase order with vendor and shipping addresses."""
    print("Example 4: Purchase Order with Full Address Details")
    print("-" * 50)

    result = purchase_order_add(
        vendor_name="Hardware Wholesale",
        vendor_address={
            "Addr1": "456 Vendor Lane",
            "City": "Chicago",
            "State": "IL",
            "PostalCode": "60601"
        },
        ship_address={
            "Addr1": "789 Warehouse Blvd",
            "Addr2": "Building B, Dock 3",
            "City": "Houston",
            "State": "TX",
            "PostalCode": "77001"
        },
        line_items=[
            {
                "item_name": "Hammer",
                "quantity": 50,
                "rate": 12.50,
                "desc": "16oz claw hammer"
            },
            {
                "item_name": "Screwdriver Set",
                "quantity": 25,
                "rate": 18.75,
                "desc": "10-piece set"
            }
        ],
        ship_method="UPS Ground",
        fob="FOB Destination",
        expected_date="2024-02-01",
        memo="Monthly tool order"
    )

    if result["success"]:
        print(f"✓ Successfully created purchase order with addresses")
        print(f"  PO Number: {result['RefNumber']}")
        print(f"  Total Amount: ${result['TotalAmount']:.2f}")
    else:
        print(f"✗ Error: {result['error_message']}")

    print()


if __name__ == "__main__":
    print("=" * 50)
    print("QuickBooks Purchase Order Add Examples")
    print("=" * 50)
    print()

    # Run examples
    example_simple_purchase_order()
    example_dropship_purchase_order()
    example_purchase_order_with_classes()
    example_purchase_order_with_addresses()

    print("=" * 50)
    print("Examples complete!")
    print("=" * 50)
