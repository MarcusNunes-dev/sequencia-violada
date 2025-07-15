Function Base64Encode(text As String) As String
    Dim arr() As Byte
    Dim objXML As Object
    Dim objNode As Object
 
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
 
    arr = StrConv(text, vbFromUnicode)
    objNode.DataType = "bin.base64"
    objNode.nodeTypedValue = arr
    Base64Encode = Replace(objNode.text, vbLf, "")
 
    Set objNode = Nothing
    Set objXML = Nothing
End Function
