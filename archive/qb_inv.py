import csv
import xml.etree.ElementTree as ET
import win32com.client
import pythoncom
from datetime import date

def build_qbxml_invoice_request(invoice_number):
    """InvoiceQuery filtered by a single invoice number."""
    return (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<?qbxml version="16.0"?>\n'
        '<QBXML>\n'
        '  <QBXMLMsgsRq onError="continueOnError">\n'
        f'    <InvoiceQueryRq requestID="1">\n'
        f'      <RefNumber>{invoice_number}</RefNumber>\n'
        '      <IncludeLineItems>1</IncludeLineItems>\n'
        '    </InvoiceQueryRq>\n'
        '  </QBXMLMsgsRq>\n'
        '</QBXML>'
    )

def build_qbxml_year_invoices_request(year):
    """
    InvoiceQueryRq that returns every invoice from Jan 1 of `year` through today,
    with line items included. Filter must appear before IncludeLineItems.
    """
    from_date = f"{year}-01-01"
    to_date   = date.today().isoformat()

    qbxml = (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<?qbxml version="16.0"?>\n'
        '<QBXML>\n'
        '  <QBXMLMsgsRq onError="continueOnError">\n'
        '    <InvoiceQueryRq requestID="1">\n'
        '      <TxnDateRangeFilter>\n'
        f'        <FromTxnDate>{from_date}</FromTxnDate>\n'
        f'        <ToTxnDate>{to_date}</ToTxnDate>\n'
        '      </TxnDateRangeFilter>\n'
        '      <IncludeLineItems>1</IncludeLineItems>\n'
        '    </InvoiceQueryRq>\n'
        '  </QBXMLMsgsRq>\n'
        '</QBXML>'
    )
    print("DEBUG: Generated QBXML:\n", qbxml)  # for validation in qbValidator.exe
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

def process_invoice_response(response_xml):
    root = ET.fromstring(response_xml)
    data = []

    for inv in root.findall('.//InvoiceRet'):
        # Header fields
        invoice_number = inv.findtext('RefNumber', "")
        invoice_date   = inv.findtext('TxnDate',   "")
        po_number      = inv.findtext('PONumber',  "")
        customer_elem  = inv.find('CustomerRef')
        if customer_elem is not None:
            customer_name = customer_elem.findtext('FullName', "")
        else:
            customer_name = ""

        ship_address = parse_address(inv.find('ShipAddress'))

        # Line items
        line_items = inv.findall('InvoiceLineRet')
        if not line_items:
            data.append({
                "Invoice Number": invoice_number,
                "Customer Name":  customer_name,
                "Invoice Date":   invoice_date,
                "PO Number":      po_number,
                "Ship To":        ship_address,
                "Line Description": "",
                "Quantity":         "",
                "Rate":             "",
                "Amount":           "",
                "Item Ref Full Name": ""
            })
        else:
            for li in line_items:
                line_desc     = li.findtext('Desc',     "")
                quantity      = li.findtext('Quantity',"")
                rate          = li.findtext('Rate',     "")
                amount        = li.findtext('Amount',   "")
                item_ref_elem = li.find('ItemRef')
                if item_ref_elem is not None:
                    item_full_name = item_ref_elem.findtext('FullName', "")
                else:
                    item_full_name = ""

                data.append({
                    "Invoice Number": invoice_number,
                    "Customer Name":  customer_name,
                    "Invoice Date":   invoice_date,
                    "PO Number":      po_number,
                    "Ship To":        ship_address,
                    "Line Description": line_desc,
                    "Quantity":         quantity,
                    "Rate":             rate,
                    "Amount":           amount,
                    "Item Ref Full Name": item_full_name
                })

    return data

def export_to_csv(data, filename):
    fieldnames = [
        "Invoice Number","Customer Name","Invoice Date","PO Number","Ship To",
        "Line Description","Quantity","Rate","Amount","Item Ref Full Name"
    ]
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)
    print(f"Export complete! Data saved to {filename}")

def main():
    choice = input("Fetch by invoice numbers (n) or by year (y)? ").strip().lower()

    rp = None
    session = None
    try:
        pythoncom.CoInitialize()
        rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
        rp.OpenConnection("", "PythonInvoiceQBXMLApp")
        session = rp.BeginSession("", 2)

        if choice == 'y':
            year = input("Enter year (e.g. 2023): ").strip()
            qbxml_request = build_qbxml_year_invoices_request(year)
            print(f"\nSending QBXML Request for invoices from {year}-01-01 through today...")
            response = rp.ProcessRequest(session, qbxml_request)
            data = process_invoice_response(response)
            export_to_csv(data, filename=f"invoices_from_{year}.csv")

        elif choice == 'n':
            inv_input = input("Enter comma-separated invoice numbers: ").strip()
            invoice_numbers = [i.strip() for i in inv_input.split(",") if i.strip()]
            if not invoice_numbers:
                print("No invoice numbers provided; exiting.")
            else:
                for inv in invoice_numbers:
                    qbxml_request = build_qbxml_invoice_request(inv)
                    print(f"\nSending QBXML Request for Invoice {inv}...")
                    response = rp.ProcessRequest(session, qbxml_request)
                    data = process_invoice_response(response)
                    export_to_csv(data, filename=f"invoice_{inv}.csv")

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
