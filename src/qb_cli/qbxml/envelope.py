def wrap_request(inner_body: str) -> str:
    return (
        '<?xml version="1.0" encoding="utf-8"?>\n'
        '<?qbxml version="16.0"?>\n'
        '<QBXML>\n'
        '  <QBXMLMsgsRq onError="continueOnError">\n'
        f'{inner_body}\n'
        '  </QBXMLMsgsRq>\n'
        '</QBXML>\n'
    )
