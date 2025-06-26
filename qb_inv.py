import csv
import xml.etree.ElementTree as ET
import win32com.client
import pythoncom


def build_qbxml_customers_request(iterator_mode, iterator_id=None):
    # Build the CustomerQueryRq with FullName + ShipToAddressList
    iterator_attr = f' iterator="{iterator_mode}"'
    if iterator_id:
        iterator_attr += f' iteratorID="{iterator_id}"'

    qbxml = (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<?qbxml version="16.0"?>\n'
        '<QBXML>\n'
        '  <QBXMLMsgsRq onError="continueOnError">\n'
        f'    <CustomerQueryRq requestID="1"{iterator_attr}>\n'
        '      <IncludeRetElement>FullName</IncludeRetElement>\n'
        '      <IncludeRetElement>ShipToAddressList</IncludeRetElement>\n'
        '      <MaxReturned>100</MaxReturned>\n'
        '    </CustomerQueryRq>\n'
        '  </QBXMLMsgsRq>\n'
        '</QBXML>'
    )
    return qbxml


def parse_shipto(response_xml):
    # Extract customer names and their ShipTo addresses
    root = ET.fromstring(response_xml)
    records = []

    for cust in root.findall('.//CustomerRet'):
        name = cust.findtext('FullName', '').strip()
        st_list = cust.find('ShipToAddressList')
        if st_list is None:
            continue

        for st in st_list.findall('ShipToAddress'):
            label = st.findtext('Name', '').strip()
            parts = [st.findtext(tag, '').strip() for tag in (
                'Addr1', 'Addr2', 'City', 'State', 'PostalCode', 'Country'
            )]
            address = ", ".join(p for p in parts if p)
            entry = f"{label}: {address}" if label else address
            records.append((name, entry))

    return records


def export_to_csv(records, filename="shipto_addresses.csv"):
    # Write out to CSV
    with open(filename, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)
        writer.writerow(['Customer', 'ShipToAddress'])
        writer.writerows(records)
    print(f"✅ Exported {len(records)} records to {filename}")


def main():
    # Initialize COM and QuickBooks session
    pythoncom.CoInitialize()
    rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
    rp.OpenConnection("", "QB ShipTo Exporter")
    session = rp.BeginSession("", 2)

    all_records = []
    iterator_id = None

    try:
        while True:
            mode = 'Start' if iterator_id is None else 'Continue'
            qbxml = build_qbxml_customers_request(mode, iterator_id)
            response = rp.ProcessRequest(session, qbxml)
            records = parse_shipto(response)
            all_records.extend(records)

            # Check iteratorRemainingCount
            root = ET.fromstring(response)
            rs = root.find('.//CustomerQueryRs')
            if rs is None or rs.get('iteratorRemainingCount') in (None, '0'):
                break
            iterator_id = rs.get('iteratorID')

    finally:
        try:
            rp.EndSession(session)
        except:
            pass
        try:
            rp.CloseConnection()
        except:
            pass
        pythoncom.CoUninitialize()

    export_to_csv(all_records)


if __name__ == '__main__':
    main()
