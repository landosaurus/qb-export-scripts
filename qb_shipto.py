import csv
import xml.etree.ElementTree as ET
import win32com.client
import pythoncom
from datetime import datetime

def build_qbxml_customers_request(iterator_mode, iterator_id=None):
    # Build CustomerQueryRq WITHOUT IncludeRetElement to get all fields including ShipToAddress
    iterator_attr = f' iterator="{iterator_mode}"'
    if iterator_id:
        iterator_attr += f' iteratorID="{iterator_id}"'
    
    # Don't use IncludeRetElement - let QB return all fields including ShipToAddress
    qbxml = (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<?qbxml version="16.0"?>\n'
        '<QBXML>\n'
        '  <QBXMLMsgsRq onError="continueOnError">\n'
        f'    <CustomerQueryRq requestID="1"{iterator_attr}>\n'
        '      <MaxReturned>100</MaxReturned>\n'
        '    </CustomerQueryRq>\n'
        '  </QBXMLMsgsRq>\n'
        '</QBXML>'
    )
    return qbxml

def parse_shipto(response_xml):
    # Extract customer names and their ShipToAddress entries
    root = ET.fromstring(response_xml)
    records = []
    
    for cust in root.findall('.//CustomerRet'):
        name = cust.findtext('FullName', '').strip()
        
        # Look for ShipToAddress elements directly under CustomerRet
        ship_addresses = cust.findall('ShipToAddress')
        
        for st in ship_addresses:
            # Extract each field separately
            record = {
                'Customer': name,
                'ShipToName': st.findtext('Name', '').strip(),
                'Addr1': st.findtext('Addr1', '').strip(),
                'Addr2': st.findtext('Addr2', '').strip(),
                'Addr3': st.findtext('Addr3', '').strip(),
                'Addr4': st.findtext('Addr4', '').strip(),
                'Addr5': st.findtext('Addr5', '').strip(),
                'City': st.findtext('City', '').strip(),
                'State': st.findtext('State', '').strip(),
                'PostalCode': st.findtext('PostalCode', '').strip(),
                'Country': st.findtext('Country', '').strip(),
                'Note': st.findtext('Note', '').strip(),
                'DefaultShipTo': st.findtext('DefaultShipTo', '').strip()
            }
            
            # Only add if there's at least some address data
            if any([record['ShipToName'], record['Addr1'], record['City']]):
                records.append(record)
    
    return records

def export_to_csv(records, filename="shipto_addresses.csv", exclude_empty_columns=True):
    # Write out to CSV with timestamp
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename_with_time = f"shipto_addresses_{timestamp}.csv"
    
    if records:
        # Define all possible columns
        all_columns = ['Customer', 'ShipToName', 'Addr1', 'Addr2', 'Addr3', 'Addr4', 'Addr5', 
                      'City', 'State', 'PostalCode', 'Country', 'Note', 'DefaultShipTo']
        
        if exclude_empty_columns:
            # Find columns that have at least one non-empty value
            columns_with_data = []
            for col in all_columns:
                if any(record.get(col, '') for record in records):
                    columns_with_data.append(col)
            columns = columns_with_data
            print(f"Including columns: {', '.join(columns)}")
        else:
            columns = all_columns
        
        with open(filename_with_time, 'w', newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=columns, extrasaction='ignore')
            writer.writeheader()
            writer.writerows(records)
        
        print(f"✅ Exported {len(records)} records to {filename_with_time}")
        
        # Show summary of address usage
        print("\nAddress field usage summary:")
        for col in ['Addr1', 'Addr2', 'Addr3', 'Addr4', 'Addr5']:
            count = sum(1 for r in records if r.get(col, ''))
            if count > 0:
                print(f"  {col}: {count} addresses use this line")
    else:
        print("No records to export")

def main():
    print("=== QuickBooks ShipTo Address Exporter ===")
    print(f"Started at: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
    
    # Initialize COM and QuickBooks session
    pythoncom.CoInitialize()
    rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
    rp.OpenConnection("", "QB ShipTo Exporter")
    session = rp.BeginSession("", 2)
    
    all_records = []
    iterator_id = None
    batch_count = 0
    total_customers = 0
    
    try:
        while True:
            batch_count += 1
            mode = 'Start' if iterator_id is None else 'Continue'
            qbxml = build_qbxml_customers_request(mode, iterator_id)
            
            print(f"Processing batch {batch_count}...")
            response = rp.ProcessRequest(session, qbxml)
            
            # Debug: save first batch response
            if batch_count == 1:
                with open('debug_first_batch.xml', 'w', encoding='utf-8') as f:
                    f.write(response)
                print("Saved first batch to debug_first_batch.xml for inspection")
            
            # Parse the response
            root = ET.fromstring(response)
            customers = root.findall('.//CustomerRet')
            total_customers += len(customers)
            
            # Extract ShipTo addresses
            records = parse_shipto(response)
            all_records.extend(records)
            
            print(f"  Found {len(customers)} customers, {len(records)} ShipTo addresses")
            
            # Check iteratorRemainingCount
            rs = root.find('.//CustomerQueryRs')
            if rs is None or rs.get('iteratorRemainingCount') in (None, '0'):
                break
            iterator_id = rs.get('iteratorID')
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        
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
    
    print(f"\n{'='*50}")
    print(f"Processed {total_customers} total customers")
    print(f"Found {len(all_records)} ShipTo addresses")
    
    if all_records:
        # Show sample of data
        print(f"\nSample data (first 3 records):")
        for i, record in enumerate(all_records[:3]):
            print(f"  {i+1}. {record['Customer']} -> {record['ShipToName']} ({record['City']}, {record['State']})")
        
        # Export with option to exclude empty columns (default: True)
        export_to_csv(all_records, exclude_empty_columns=True)
    else:
        print("\nNo ShipTo addresses found!")
        print("This could mean:")
        print("  1. Your customers don't have ShipTo addresses defined")
        print("  2. Only the default ShipAddress is used (not ShipToAddress)")
        print("\nCheck debug_first_batch.xml to see the actual customer data structure")

if __name__ == '__main__':
    main()
