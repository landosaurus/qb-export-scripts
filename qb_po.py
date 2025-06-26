import csv
import xml.etree.ElementTree as ET
import win32com.client
import pythoncom
from datetime import date

def build_qbxml_po_request(po_number):
    """PurchaseOrderQuery filtered by a single PO number."""
    return (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<?qbxml version="16.0"?>\n'
        '<QBXML>\n'
        '  <QBXMLMsgsRq onError="continueOnError">\n'
        f'    <PurchaseOrderQueryRq requestID="1">\n'
        f'      <RefNumber>{po_number}</RefNumber>\n'
        '      <IncludeLineItems>1</IncludeLineItems>\n'
        '    </PurchaseOrderQueryRq>\n'
        '  </QBXMLMsgsRq>\n'
        '</QBXML>'
    )

def build_qbxml_year_po_request(year):
    """
    PurchaseOrderQueryRq that returns every PO from Jan 1 of `year` through today,
    with line items included. Filter must appear before IncludeLineItems.
    """
    from_date = f"{year}-01-01"
    to_date   = date.today().isoformat()

    qbxml = (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<?qbxml version="16.0"?>\n'
        '<QBXML>\n'
        '  <QBXMLMsgsRq onError="continueOnError">\n'
        '    <PurchaseOrderQueryRq requestID="1">\n'
        '      <TxnDateRangeFilter>\n'
        f'        <FromTxnDate>{from_date}</FromTxnDate>\n'
        f'        <ToTxnDate>{to_date}</ToTxnDate>\n'
        '      </TxnDateRangeFilter>\n'
        '      <IncludeLineItems>1</IncludeLineItems>\n'
        '    </PurchaseOrderQueryRq>\n'
        '  </QBXMLMsgsRq>\n'
        '</QBXML>'
    )
    print("DEBUG: Generated QBXML:\n", qbxml)  # for validation in qbValidator.exe
    return qbxml

def parse_address(address_element):
    """
    Helper function to parse an address block from QBXML.
    Concatenates common address fields (for example, Addr1, City, State, etc.) into a single string.
    """
    if address_element is None:
        return ""
    parts = []
    for tag in ["Addr1", "Addr2", "Addr3", "Addr4", "Addr5", "City", "State", "PostalCode", "Country"]:
        text = address_element.findtext(tag, "").strip()
        if text:
            parts.append(text)
    return ", ".join(parts)

def process_po_response(response_xml):
    """
    Parses the QBXML response for purchase orders and returns a list of dictionaries for CSV export.

    For each purchase order (PurchaseOrderRet element) extracted, it collects header information:
      - Purchase Order Number (RefNumber)
      - Vendor Name (from VendorRef/FullName)
      - Transaction Date (TxnDate)
      - Due Date (DueDate, if available)
      - Vendor Address (parsed from VendorAddress)
      - Ship To (parsed from ShipAddress)
      - Total Amount (TotalAmount, if available)

    For each line item (PurchaseOrderLineRet), it extracts:
      - Line Description (Desc)
      - Quantity, Rate, Amount
      - Item Ref Full Name (from ItemRef/FullName, if available)

    Also handles grouped line items (PurchaseOrderLineGroupRet) if present.

    If no line items are found, a single row is created with blank details for the line items.
    """
    root = ET.fromstring(response_xml)
    pos = root.findall('.//PurchaseOrderRet')
    data = []
    
    for po in pos:
        # Header fields
        po_number = po.findtext('RefNumber', default="")
        txn_date = po.findtext('TxnDate', default="")
        due_date = po.findtext('DueDate', default="")
        vendor_elem = po.find('VendorRef')
        vendor_name = vendor_elem.findtext('FullName', default="") if vendor_elem is not None else ""
        vendor_address = parse_address(po.find('VendorAddress'))
        ship_address = parse_address(po.find('ShipAddress'))
        total_amount = po.findtext('TotalAmount', default="")

        # Find simple line items and grouped line items
        line_items = po.findall('PurchaseOrderLineRet')
        group_items = po.findall('PurchaseOrderLineGroupRet')

        # If no line items were returned, output a header-only record.
        if not line_items and not group_items:
            data.append({
                "PO Number": po_number,
                "Vendor Name": vendor_name,
                "Transaction Date": txn_date,
                "Due Date": due_date,
                "Vendor Address": vendor_address,
                "Ship To": ship_address,
                "Total Amount": total_amount,
                "Line Description": "",
                "Quantity": "",
                "Rate": "",
                "Amount": "",
                "Item Ref Full Name": ""
            })
        else:
            # Process simple PurchaseOrderLineRet items.
            for li in line_items:
                line_desc = li.findtext('Desc', default="")
                quantity = li.findtext('Quantity', default="")
                rate = li.findtext('Rate', default="")
                amount = li.findtext('Amount', default="")
                item_ref_elem = li.find('ItemRef')
                item_full_name = item_ref_elem.findtext('FullName', default="") if item_ref_elem is not None else ""
                
                data.append({
                    "PO Number": po_number,
                    "Vendor Name": vendor_name,
                    "Transaction Date": txn_date,
                    "Due Date": due_date,
                    "Vendor Address": vendor_address,
                    "Ship To": ship_address,
                    "Total Amount": total_amount,
                    "Line Description": line_desc,
                    "Quantity": quantity,
                    "Rate": rate,
                    "Amount": amount,
                    "Item Ref Full Name": item_full_name
                })
            # Process grouped line items (PurchaseOrderLineGroupRet)
            for group in group_items:
                # The group typically includes a description, quantity, and a total amount for the group.
                line_desc = group.findtext('Desc', default="")
                quantity = group.findtext('Quantity', default="")
                amount = group.findtext('TotalAmount', default="")
                # The item reference for a grouped item is under ItemGroupRef.
                item_group_elem = group.find('ItemGroupRef')
                item_full_name = item_group_elem.findtext('FullName', default="") if item_group_elem is not None else ""
                
                data.append({
                    "PO Number": po_number,
                    "Vendor Name": vendor_name,
                    "Transaction Date": txn_date,
                    "Due Date": due_date,
                    "Vendor Address": vendor_address,
                    "Ship To": ship_address,
                    "Total Amount": total_amount,
                    "Line Description": line_desc,
                    "Quantity": quantity,
                    "Rate": "",  # Group items generally do not include a per-item rate
                    "Amount": amount,
                    "Item Ref Full Name": item_full_name
                })
    return data

def export_to_csv(data, filename):
    """
    Exports the provided list of dictionaries (data rows) to a CSV file.
    """
    fieldnames = [
        "PO Number",
        "Vendor Name",
        "Transaction Date",
        "Due Date",
        "Vendor Address",
        "Ship To",
        "Total Amount",
        "Line Description",
        "Quantity",
        "Rate",
        "Amount",
        "Item Ref Full Name"
    ]
    with open(filename, mode="w", newline="", encoding="utf-8") as csvfile:
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)
    print(f"Export complete! Data saved to {filename}")

def main():
    choice = input("Fetch by PO numbers (n) or by year (y)? ").strip().lower()

    rp = None
    session = None
    try:
        pythoncom.CoInitialize()
        rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
        rp.OpenConnection("", "PythonPOQBXMLApp")
        session = rp.BeginSession("", 2)

        if choice == 'y':
            year = input("Enter year (e.g. 2023): ").strip()
            qbxml_request = build_qbxml_year_po_request(year)
            print(f"\nSending QBXML Request for purchase orders from {year}-01-01 through today...")
            response = rp.ProcessRequest(session, qbxml_request)
            data = process_po_response(response)
            export_to_csv(data, filename=f"purchase_orders_from_{year}.csv")

        elif choice == 'n':
            po_input = input("Enter comma-separated PO numbers: ").strip()
            po_numbers = [po.strip() for po in po_input.split(",") if po.strip()]
            if not po_numbers:
                print("No PO numbers provided; exiting.")
            else:
                for po in po_numbers:
                    qbxml_request = build_qbxml_po_request(po)
                    print(f"\nSending QBXML Request for Purchase Order {po}...")
                    response = rp.ProcessRequest(session, qbxml_request)
                    data = process_po_response(response)
                    export_to_csv(data, filename=f"purchase_order_{po}.csv")

        else:
            print("Invalid choice; please run again and enter 'n' or 'y'.")

    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
    except Exception as e:
        print("Error communicating with QuickBooks:", e)
    finally:
        # clean up COM session—even if errors or Ctrl+C
        if rp and session:
            try: rp.EndSession(session)
            except: pass
        if rp:
            try: rp.CloseConnection()
            except: pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
