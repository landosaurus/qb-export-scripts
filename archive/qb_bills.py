import csv
import xml.etree.ElementTree as ET
import win32com.client
import pythoncom
from datetime import date

def build_qbxml_bill_request(ref_number):
    """BillQuery filtered by a single reference number."""
    return (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<?qbxml version="16.0"?>\n'
        '<QBXML>\n'
        '  <QBXMLMsgsRq onError="continueOnError">\n'
        f'    <BillQueryRq requestID="1">\n'
        f'      <RefNumber>{ref_number}</RefNumber>\n'
        '      <IncludeLineItems>1</IncludeLineItems>\n'
        '    </BillQueryRq>\n'
        '  </QBXMLMsgsRq>\n'
        '</QBXML>'
    )

def build_qbxml_year_bills_request(year):
    """
    BillQueryRq that returns every bill from Jan 1 of `year` through today,
    with line items included. Filter must appear before IncludeLineItems.
    """
    from_date = f"{year}-01-01"
    to_date   = date.today().isoformat()

    qbxml = (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<?qbxml version="16.0"?>\n'
        '<QBXML>\n'
        '  <QBXMLMsgsRq onError="continueOnError">\n'
        '    <BillQueryRq requestID="1">\n'
        '      <TxnDateRangeFilter>\n'
        f'        <FromTxnDate>{from_date}</FromTxnDate>\n'
        f'        <ToTxnDate>{to_date}</ToTxnDate>\n'
        '      </TxnDateRangeFilter>\n'
        '      <IncludeLineItems>1</IncludeLineItems>\n'
        '    </BillQueryRq>\n'
        '  </QBXMLMsgsRq>\n'
        '</QBXML>'
    )
    print("DEBUG: Generated QBXML:\n", qbxml)
    return qbxml

def parse_address(address_element):
    if address_element is None:
        return ""
    parts = []
    for tag in ["Addr1","Addr2","Addr3","Addr4","Addr5","City","State","PostalCode","Country"]:
        text = address_element.findtext(tag, "").strip()
        if text:
            parts.append(text)
    return ", ".join(parts)

def process_bill_response(response_xml):
    root = ET.fromstring(response_xml)
    data = []

    for bill in root.findall('.//BillRet'):
        # Header fields
        ref_number   = bill.findtext('RefNumber', "")
        txn_id       = bill.findtext('TxnID', "")
        bill_date    = bill.findtext('TxnDate', "")
        due_date     = bill.findtext('DueDate', "")
        amount_due   = bill.findtext('AmountDue', "")
        memo         = bill.findtext('Memo', "")

        vendor_elem = bill.find('VendorRef')
        if vendor_elem is not None:
            vendor_name = vendor_elem.findtext('FullName', "")
        else:
            vendor_name = ""

        vendor_address = parse_address(bill.find('VendorAddress'))

        # Item line items
        item_lines = bill.findall('ItemLineRet')
        # Expense line items
        expense_lines = bill.findall('ExpenseLineRet')

        all_lines = []

        for li in item_lines:
            line_desc     = li.findtext('Desc', "")
            quantity      = li.findtext('Quantity', "")
            cost          = li.findtext('Cost', "")
            amount        = li.findtext('Amount', "")
            item_ref_elem = li.find('ItemRef')
            if item_ref_elem is not None:
                item_full_name = item_ref_elem.findtext('FullName', "")
            else:
                item_full_name = ""

            account_ref_elem = li.find('AccountRef')
            if account_ref_elem is not None:
                account_name = account_ref_elem.findtext('FullName', "")
            else:
                account_name = ""

            all_lines.append({
                "Line Type": "Item",
                "Line Description": line_desc,
                "Quantity": quantity,
                "Cost": cost,
                "Amount": amount,
                "Item Ref Full Name": item_full_name,
                "Account": account_name
            })

        for el in expense_lines:
            line_desc = el.findtext('Memo', "")
            amount    = el.findtext('Amount', "")

            account_ref_elem = el.find('AccountRef')
            if account_ref_elem is not None:
                account_name = account_ref_elem.findtext('FullName', "")
            else:
                account_name = ""

            all_lines.append({
                "Line Type": "Expense",
                "Line Description": line_desc,
                "Quantity": "",
                "Cost": "",
                "Amount": amount,
                "Item Ref Full Name": "",
                "Account": account_name
            })

        if not all_lines:
            data.append({
                "Ref Number": ref_number,
                "TxnID": txn_id,
                "Vendor Name": vendor_name,
                "Bill Date": bill_date,
                "Due Date": due_date,
                "Amount Due": amount_due,
                "Memo": memo,
                "Vendor Address": vendor_address,
                "Line Type": "",
                "Line Description": "",
                "Quantity": "",
                "Cost": "",
                "Amount": "",
                "Item Ref Full Name": "",
                "Account": ""
            })
        else:
            for line in all_lines:
                data.append({
                    "Ref Number": ref_number,
                    "TxnID": txn_id,
                    "Vendor Name": vendor_name,
                    "Bill Date": bill_date,
                    "Due Date": due_date,
                    "Amount Due": amount_due,
                    "Memo": memo,
                    "Vendor Address": vendor_address,
                    **line
                })

    return data

def export_to_csv(data, filename):
    fieldnames = [
        "Ref Number", "TxnID", "Vendor Name", "Bill Date", "Due Date",
        "Amount Due", "Memo", "Vendor Address", "Line Type", "Line Description",
        "Quantity", "Cost", "Amount", "Item Ref Full Name", "Account"
    ]
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)
    print(f"Export complete! Data saved to {filename}")

def main():
    choice = input("Fetch by bill reference numbers (n) or by year (y)? ").strip().lower()

    rp = None
    session = None
    try:
        pythoncom.CoInitialize()
        rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
        rp.OpenConnection("", "PythonBillQBXMLApp")
        session = rp.BeginSession("", 2)

        if choice == 'y':
            year = input("Enter year (e.g. 2023): ").strip()
            qbxml_request = build_qbxml_year_bills_request(year)
            print(f"\nSending QBXML Request for bills from {year}-01-01 through today...")
            response = rp.ProcessRequest(session, qbxml_request)
            data = process_bill_response(response)
            export_to_csv(data, filename=f"bills_from_{year}.csv")

        elif choice == 'n':
            ref_input = input("Enter comma-separated bill reference numbers: ").strip()
            ref_numbers = [r.strip() for r in ref_input.split(",") if r.strip()]
            if not ref_numbers:
                print("No reference numbers provided; exiting.")
            else:
                for ref in ref_numbers:
                    qbxml_request = build_qbxml_bill_request(ref)
                    print(f"\nSending QBXML Request for Bill {ref}...")
                    response = rp.ProcessRequest(session, qbxml_request)
                    data = process_bill_response(response)
                    export_to_csv(data, filename=f"bill_{ref}.csv")

        else:
            print("Invalid choice; please run again and enter 'n' or 'y'.")

    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
    except Exception as e:
        print("Error communicating with QuickBooks:", e)
    finally:
        if rp and session:
            try: rp.EndSession(session)
            except: pass
        if rp:
            try: rp.CloseConnection()
            except: pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()
