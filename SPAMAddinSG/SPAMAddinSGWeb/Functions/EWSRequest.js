// The following string is a valid SOAP envelope and request for getting the properties of a mail item
function getItemDataSoap() {
    return '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
        '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
        '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
        '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '  <soap:Header>' +
        '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
        '  </soap:Header>' +
        '  <soap:Body>' +
        '    <GetItem' +
        '                xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '                xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '      <ItemShape>' +
        '        <t:BaseShape>IdOnly</t:BaseShape>' +
        '      </ItemShape>' +
        '      <ItemIds>' +
        '        <t:ItemId Id="' + Office.context.mailbox.item.itemId + '"/>' +
        '      </ItemIds>' +
        '    </GetItem>' +
        '  </soap:Body>' +
        '</soap:Envelope>';
}

// The following string is a valid SOAP envelope and request for Forward a mail item
function getForwardItemSoap(addressesSoap, bodyEmail, changeKey) {
    return '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
        '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
        '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
        '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
        '  <soap:Header>' +
        '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
        '  </soap:Header>' +
        '  <soap:Body>' +
        '    <m:CreateItem MessageDisposition="SendAndSaveCopy">' +
        '      <m:Items>' +
        '        <t:ForwardItem>' +
        '          <t:ToRecipients>' + addressesSoap + '</t:ToRecipients>' +
        '          <t:ReferenceItemId Id="' + Office.context.mailbox.item.itemId + '" ChangeKey="' + changeKey + '" />' +
        '          <t:NewBodyContent BodyType="Text">' + bodyEmail + '</t:NewBodyContent>' +
        '        </t:ForwardItem>' +
        '      </m:Items>' +
        '    </m:CreateItem>' +
        '  </soap:Body>' +
        '</soap:Envelope>';
}

function getSelectedEmailHeaders() {
    // Wrap an Exchange Web Services request in a SOAP envelope.
    return '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages" xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types" xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
        '  <soap:Header>' +
        '    <t:RequestServerVersion Version="Exchange2010" />' +
        '  </soap:Header>' +
        '  <soap:Body>' +
        '    <m:GetItem>' +
        '      <m:ItemShape>' +
        '        <t:BaseShape>IdOnly</t:BaseShape>' +
        '        <t:AdditionalProperties>' +
        '          <t:FieldURI FieldURI="item:Subject" />' +
        '          <t:FieldURI FieldURI="item:MimeContent" />' +
        '        </t:AdditionalProperties>' +
        '      </m:ItemShape>' +
        '      <m:ItemIds>' +
        '         <t:ItemId Id="' + Office.context.mailbox.item.itemId + '" />' +
        '      </m:ItemIds>' +
        '    </m:GetItem>' +
        '  </soap:Body>' +
        '</soap:Envelope>';
}

// The following string is a valid SOAP envelope and request for deleteing a mail item
function getDeleteItemSoap(changeKey) {
    return '<?xml version="1.0" encoding="utf-8"?>' +
        '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
        '               xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
        '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"' +
        '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
        '    <soap:Header>' +
        '        <t:RequestServerVersion Version="Exchange2013" />' +
        '    </soap:Header>' +
        '    <soap:Body>' +
        '        <m:MarkAsJunk IsJunk="true" MoveItem="true">' +
        '            <m:ItemIds>' +
        '                <t:ItemId Id="' + Office.context.mailbox.item.itemId + '" ChangeKey="' + changeKey + '" />' +
        '            </m:ItemIds>' +
        '        </m:MarkAsJunk>' +
        '   </soap:Body>' +
        '</soap:Envelope>';
}