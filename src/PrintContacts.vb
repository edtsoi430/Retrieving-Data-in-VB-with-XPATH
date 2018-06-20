Sub PrintContacts(contacts As IXMLDOMSelection)
  'A general function to call PrintContact in order to process all the contacts one by one.
    Debug.Print "Printing Contacts"
    Dim contact As IXMLDOMElement
    Dim curRow As Integer, contactCount As Integer, incre As Double
    curRow = 1
    For Each contact In contacts
        contactCount = contactCount + 1
    Next contact
    'Estimate increment
    'Avoid division by zero error
    incre = contactCount / 33
    If incre < 1 Then
        incre = 1
    End If
    For Each contact In contacts
        PrintContact contact, curRow
        If curRow Mod incre = 0 Then
            increCount = increCount + 1
            progress CInt(increCount)
        End If
        
        curRow = curRow + 1
    Next contact
    
    SampleReportSheet.PivotTables(1).SourceData = "Contacts!R1C1:R65536C" + LTrim(Str(NumberOfColumns))
    SampleReportSheet.PivotTables(1).RefreshTable
End Sub

Sub PrintContact(contact As IXMLDOMElement, rowIdx As Integer)
' Process details for each contact in the database
    ContactsSheet.Range(cn.outContactsList)(rowIdx, 1) = contact.SelectSingleNode("Id").Text
    Dim colIdx As Integer
    colIdx = 2
    Dim fieldItem As IXMLDOMElement
    For Each fieldItem In ContactFields
        fieldType = fieldItem.SelectSingleNode("Type").Text
        fieldName = fieldItem.SelectSingleNode("FieldName").Text
        
        If Not IsFieldExcluded(fieldName) Then
            Dim valueNode As IXMLDOMElement
            Dim fieldNode As IXMLDOMElement
            
            Set valueNode = Nothing
            Set fieldNodes = contact.SelectNodes("FieldValues/ContactField")
            For Each fieldNode In fieldNodes
                Dim localFieldName As IXMLDOMElement
                Set localFieldName = fieldNode.SelectSingleNode("FieldName")
                If localFieldName.Text = fieldName Then
                    Set valueNode = fieldNode.SelectSingleNode("Value")
                End If
            Next fieldNode
            
            If Not (valueNode Is Nothing) Then
                ContactsSheet.Range(cn.outContactsList)(rowIdx, colIdx).NumberFormat = GetCellFormat(fieldType)
                If fieldType = "DateTime" And Not valueNode.Text = "" Then
                    '2013-06-27T09:46:34
                    d = DateSerial(CInt(Mid(valueNode.Text, 1, 4)), CInt(Mid(valueNode.Text, 6, 2)), CInt(Mid(valueNode.Text, 9, 2)))
                    T = TimeSerial(CInt(Mid(valueNode.Text, 12, 2)), CInt(Mid(valueNode.Text, 15, 2)), CInt(Mid(valueNode.Text, 8, 2)))
                    ContactsSheet.Range(cn.outContactsList)(rowIdx, colIdx) = (d + T)
                ElseIf fieldType = "MultipleChoice" And InStr(valueNode.getAttribute("i:type"), "ArrayOfanyType") > 0 Then
                    Dim valueArrayItemNode As IXMLDOMElement, strValue As String
                    strValue = ""
                    For Each valueArrayItemNode In valueNode.SelectNodes("*")
                        If Len(valueArrayItemNode.Text) > 0 Then
                            strValue = strValue + ", " + valueArrayItemNode.Text
                        End If
                    Next valueArrayItemNode
                    If Len(strValue) > 2 Then
                        strValue = Mid(strValue, 3, Len(strValue) - 2)
                    End If
                    ContactsSheet.Range(cn.outContactsList)(rowIdx, colIdx) = strValue
                Else
                    ContactsSheet.Range(cn.outContactsList)(rowIdx, colIdx) = valueNode.Text
                End If
            End If
            colIdx = colIdx + 1
        End If

    Next fieldItem
    
    Set membershipLevel = contact.SelectSingleNode("MembershipLevel/Name")
    If Not (membershipLevel Is Nothing) Then
        ContactsSheet.Range(cn.outContactsList)(rowIdx, colIdx) = membershipLevel.Text
    End If
End Sub
