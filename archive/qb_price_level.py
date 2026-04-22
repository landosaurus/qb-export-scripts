import csv
import xml.etree.ElementTree as ET
import win32com.client
import pythoncom
from datetime import datetime

def build_qbxml_price_level_request(name=None):
    """Build PriceLevelQuery request. If name is provided, query specific price level, otherwise query all."""
    if name:
        # Query specific price level by name
        return (
            '<?xml version="1.0" encoding="utf-8"?>\n'
            '<?qbxml version="16.0"?>\n'
            '<QBXML>\n'
            '  <QBXMLMsgsRq onError="continueOnError">\n'
            '    <PriceLevelQueryRq requestID="1">\n'
            f'      <FullName>{name}</FullName>\n'
            '    </PriceLevelQueryRq>\n'
            '  </QBXMLMsgsRq>\n'
            '</QBXML>'
        )
    else:
        # Query all price levels
        return (
            '<?xml version="1.0" encoding="utf-8"?>\n'
            '<?qbxml version="16.0"?>\n'
            '<QBXML>\n'
            '  <QBXMLMsgsRq onError="continueOnError">\n'
            '    <PriceLevelQueryRq requestID="1">\n'
            '      <ActiveStatus>All</ActiveStatus>\n'
            '    </PriceLevelQueryRq>\n'
            '  </QBXMLMsgsRq>\n'
            '</QBXML>'
        )

def process_price_level_response(response_xml):
    """Parse the PriceLevelQuery response and extract data."""
    root = ET.fromstring(response_xml)
    data = []
    
    for pl in root.findall('.//PriceLevelRet'):
        # Common fields
        list_id = pl.findtext('ListID', "")
        name = pl.findtext('Name', "")
        is_active = pl.findtext('IsActive', "")
        price_level_type = pl.findtext('PriceLevelType', "")
        time_created = pl.findtext('TimeCreated', "")
        time_modified = pl.findtext('TimeModified', "")
        
        # Currency reference (if present)
        currency_ref = pl.find('CurrencyRef')
        currency_name = ""
        if currency_ref is not None:
            currency_name = currency_ref.findtext('FullName', "")
        
        # Check if it's a fixed percentage price level
        fixed_percentage = pl.findtext('PriceLevelFixedPercentage', "")
        
        if fixed_percentage:
            # Fixed percentage applies to all items
            data.append({
                "Price Level Name": name,
                "List ID": list_id,
                "Active": is_active,
                "Type": price_level_type,
                "Fixed Percentage": fixed_percentage,
                "Item Name": "All Items",
                "Custom Price": "",
                "Custom Price Percent": fixed_percentage,
                "Currency": currency_name,
                "Created": time_created,
                "Modified": time_modified
            })
        else:
            # Per-item price level
            per_item_list = pl.findall('PriceLevelPerItemRet')
            if not per_item_list:
                # No items in this price level
                data.append({
                    "Price Level Name": name,
                    "List ID": list_id,
                    "Active": is_active,
                    "Type": price_level_type,
                    "Fixed Percentage": "",
                    "Item Name": "",
                    "Custom Price": "",
                    "Custom Price Percent": "",
                    "Currency": currency_name,
                    "Created": time_created,
                    "Modified": time_modified
                })
            else:
                # Process each item in the price level
                for item in per_item_list:
                    item_ref = item.find('ItemRef')
                    item_name = ""
                    if item_ref is not None:
                        item_name = item_ref.findtext('FullName', "")
                    
                    custom_price = item.findtext('CustomPrice', "")
                    custom_price_percent = item.findtext('CustomPricePercent', "")
                    
                    data.append({
                        "Price Level Name": name,
                        "List ID": list_id,
                        "Active": is_active,
                        "Type": price_level_type,
                        "Fixed Percentage": "",
                        "Item Name": item_name,
                        "Custom Price": custom_price,
                        "Custom Price Percent": custom_price_percent,
                        "Currency": currency_name,
                        "Created": time_created,
                        "Modified": time_modified
                    })
    
    return data

def export_to_csv(data, filename):
    """Export the price level data to CSV."""
    fieldnames = [
        "Price Level Name", "List ID", "Active", "Type", "Fixed Percentage",
        "Item Name", "Custom Price", "Custom Price Percent", "Currency",
        "Created", "Modified"
    ]
    
    with open(filename, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(data)
    
    print(f"Export complete! Data saved to {filename}")

def main():
    choice = input("Fetch specific price levels by name (n) or all price levels (a)? ").strip().lower()
    
    rp = None
    session = None
    all_data = []
    
    try:
        pythoncom.CoInitialize()
        rp = win32com.client.Dispatch("QBXMLRP2.RequestProcessor")
        rp.OpenConnection("", "PythonPriceLevelQBXMLApp")
        session = rp.BeginSession("", 2)
        
        if choice == 'a':
            # Fetch all price levels
            qbxml_request = build_qbxml_price_level_request()
            print("\nSending QBXML Request for all price levels...")
            response = rp.ProcessRequest(session, qbxml_request)
            data = process_price_level_response(response)
            all_data.extend(data)
            
            # Generate filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"all_price_levels_{timestamp}.csv"
            export_to_csv(all_data, filename)
            
        elif choice == 'n':
            # Fetch specific price levels by name
            names_input = input('Enter comma-separated price level names (each in quotes, e.g. "Retail", "Wholesale"): ').strip()
            
            # Parse the names - handle quoted names
            import re
            # Match quoted strings
            names = re.findall(r'"([^"]+)"', names_input)
            
            if not names:
                # If no quoted names found, try splitting by comma as fallback
                names = [n.strip() for n in names_input.split(',') if n.strip()]
            
            if not names:
                print("No price level names provided; exiting.")
            else:
                for name in names:
                    qbxml_request = build_qbxml_price_level_request(name)
                    print(f'\nSending QBXML Request for Price Level "{name}"...')
                    try:
                        response = rp.ProcessRequest(session, qbxml_request)
                        data = process_price_level_response(response)
                        all_data.extend(data)
                    except Exception as e:
                        print(f'Error fetching price level "{name}": {e}')
                
                if all_data:
                    # Generate filename with timestamp
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"price_levels_{timestamp}.csv"
                    export_to_csv(all_data, filename)
                else:
                    print("No data retrieved.")
        
        else:
            print("Invalid choice; please run again and enter 'n' or 'a'.")
    
    except KeyboardInterrupt:
        print("\nOperation cancelled by user.")
    except Exception as e:
        print("Error communicating with QuickBooks:", e)
    finally:
        # Clean up COM session—even if errors or Ctrl+C
        if rp and session:
            try:
                rp.EndSession(session)
            except:
                pass
        if rp:
            try:
                rp.CloseConnection()
            except:
                pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    main()