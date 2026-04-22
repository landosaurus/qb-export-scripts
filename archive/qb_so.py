import csv
import xml.etree.ElementTree as ET
import win32com.client
import pythoncom
from datetime import date

def build_qbxml_so_request(so_number):
    """SalesOrderQuery filtered by a single SO number."""
    return (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<?qbxml version="16.0"?>\n'
        '<QBXML>\n'
        '  <QBXMLMsgsRq onError="continueOnError">\n'
        f'    <SalesOrderQueryRq requestID="1">\n'
        f'      <RefNumber>{so_number}</RefNumber>\n'
        '      <IncludeLineItems>1</IncludeLineItems>\n'
        '    </SalesOrderQueryRq>\n'
        '  </QBXMLMsgsRq>\n'
        '</QBXML>'
    )

def build_qbxml_year_so_request(year):
    """
    SalesOrderQueryRq that returns every SO from Jan 1 of `year` through today,
    with line items included. Filter must appear before IncludeLineItems.
    """
    from_date = f"{year}-01-01"
    to_date   = date.today().isoformat()

    qbxml = (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<?qbxml version="16.0"?>\n'
        '<QBXML>\n'
        '  <QBXMLMsgsRq onError="continueOnError">\n'
        '    <SalesOrderQueryRq requestID="1">\n'
        '      <TxnDateRangeFilter>\n'
        f'        <FromTxnDate>{from_date}</FromTxnDate>\n'
        f'        <ToTxnDate>{to_date}</ToTxnDate>\n'
        '      </TxnDateRangeFilter>\n'
        '      <IncludeLineItems>1</IncludeLineItems>\n'
        '    </SalesOrderQueryRq>\n'
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

def process_so_response(response_xml):
    """
    Parses the QBXML response for sales orders and returns a list of dictionaries for CSV export.

    For each sales order (SalesOrderRet element) extracted, it collects header information:
      - Sales Order Number (RefNumber)
      - Customer Name (from CustomerRef/FullName)
      - Transaction Date (TxnDate)
      - Due Date (DueDate, if available)
      - Bill To (parsed from BillAddress)
      - Ship To (parsed from ShipAddress)
      - Total Amount (TotalAmount, if available)
      - Sales Tax Total (SalesTaxTotal, if available)
      - Subtotal (Subtotal, if available)

    For each line item (SalesOrderLineRet), it extracts:
      - Line Description (Desc)
      - Quantity, Rate, Amount
      - Item Ref Full Name (from ItemRef/FullName, if available)

    Also handles grouped line items (SalesOrderLineGroupRet) if present.

    If no line items are found, a single row is created with blank details for the line items.
    """
    root = ET.fromstring(response_xml)
    sos = root.findall('.//SalesOrderRet')
    data = []
    
    for so in sos:
        # Header fields
        so_number = so.findtext('RefNumber', default="")
        txn_date = so.findtext('TxnDate', default="")
        due_date = so.findtext('DueDate', default="")
        customer_elem = so.find('CustomerRef')
        customer_name = customer_elem.findtext('FullName', default="") if customer_elem is not None else ""
        bill_address = parse_address(so.find('BillAddress'))
        ship_address = parse_address(so.find('ShipAddress'))
        total_amount = so.findtext('TotalAmount', default="")
        sales_tax_total = so.findtext('SalesTaxTotal', default="")
        subtotal = so.findtext('Subtotal', default="")

        # Find simple line items and grouped line items
        line_items = so.findall('SalesOrderLineRet')
        group_items = so.findall('SalesOrderLineGroupRet')

        # If no line items were returned, output a header-only record.
        if not line_items and not group_items:
            data.append({
                "SO Number": so_number,
                "Customer Name": customer_name,
                "Transaction Date": txn_date,
                "Due Date": due_date,
                "Bill To": bill_address,
                "Ship To": ship_address,
                "Subtotal": subtotal,
                "Sales Tax Total": sales_tax_total,
                "Total Amount": total_amount,
                "Line Description": "",
                "Quantity": "",
                "Rate": "",
                "Amount": "",
                "Item Ref Full Name": ""
            })
        else:
            # Process simple SalesOrderLineRet items.
            for li in line_items:
                line_desc = li.findtext('Desc', default="")
                quantity = li.findtext('Quantity', default="")
                rate = li.findtext('Rate', default="")
                amount = li.findtext('Amount', default="")
                item_ref_elem = li.find('ItemRef')
                item_full_name = item_ref_elem.findtext('FullName', default="") if item_ref_elem is not None else ""
                
                data.append({
                    "SO Number": so_number,
                    "Customer Name": customer_name,
                    "Transaction Date": txn_date,
                    "Due Date": due_date,
                    "Bill To": bill_address,
                    "Ship To": ship_address,
                    "Subtotal": subtotal,
                    "Sales Tax Total": sales_tax_total,
                    "Total Amount": total_amount,
                    "Line Description": line_desc,
                    "Quantity": quantity,
                    "Rate": rate,
                    "Amount": amount,
                    "Item Ref Full Name": item_full_name
                })
            # Process grouped line items (SalesOrderLineGroupRet)
            for group in group_items:
                # The group typically includes a description, quantity, and a total amount for the group.
                line_desc = group.findtext('Desc', default="")
                quantity = group.findtext('Quantity', default="")
                amount = group.findtext('TotalAmount', default="")
                # The item reference for a grouped item is under ItemGroupRef.
                item_group_elem = group.find('ItemGroupRef')
                item_full_name = item_group_elem.findtext('FullName', default="") if item_group_elem is not None else ""
                
                data.append({
                    "SO Number": so_number,
                    "Customer Name": customer_name,
                    "Transaction Date": txn_date,
                    "Due Date": due_date,
                    "Bill To": bill_address,
                    "Ship To": ship_address,
                    "Subtotal": subtotal,
                    "Sales Tax Total": sales_tax_total,
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
        "SO Number",
        "Customer Name",
        "Transaction Date",
        "Due Date",
        "Bill To",
        "Ship To",
        "Subtotal",
        "Sales Tax Total",
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
    choice = input("Fetch by SO numbers (n) or by year (y)? ").strip().lower()

    rp = None
    session = None
    try:
        pythoncom.CoInitialize()
        rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
        rp.OpenConnection("", "PythonSOQBXMLApp")
        session = rp.BeginSession("", 2)

        if choice == 'y':
            year = input("Enter year (e.g. 2023): ").strip()
            qbxml_request = build_qbxml_year_so_request(year)
            print(f"\nSending QBXML Request for sales orders from {year}-01-01 through today...")
            response = rp.ProcessRequest(session, qbxml_request)
            data = process_so_response(response)
            export_to_csv(data, filename=f"sales_orders_from_{year}.csv")

        elif choice == 'n':
            so_input = input("Enter comma-separated SO numbers: ").strip()
            so_numbers = [so.strip() for so in so_input.split(",") if so.strip()]
            if not so_numbers:
                print("No SO numbers provided; exiting.")
            else:
                for so in so_numbers:
                    qbxml_request = build_qbxml_so_request(so)
                    print(f"\nSending QBXML Request for Sales Order {so}...")
                    response = rp.ProcessRequest(session, qbxml_request)
                    data = process_so_response(response)
                    export_to_csv(data, filename=f"sales_order_{so}.csv")

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
