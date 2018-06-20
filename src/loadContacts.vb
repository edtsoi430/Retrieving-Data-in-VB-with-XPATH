Sub LoadContacts(ByVal baseUrl As String)
' Illustrate how contacts from Wild Apricot can be retrieved and loaded
    Dim url As String
    Dim filter As String, selectExpression As String
    selectExpression = BuildSelectExpression
    filter = BuildFilterExpression
    url = baseUrl
    If Len(filter) > 0 Then
        url = url + "?" + filter
    End If
    
    If Len(selectExpression) > 0 Then
        If Len(filter) > 0 Then
            url = url + "&" + selectExpression
        Else
            url = url + "?" + selectExpression
        End If
    End If
    
    Dim contactsRequestDoc As DOMDocument
    Set contactsRequestDoc = LoadXml(url)
    ContactsResultUrl = contactsRequestDoc.SelectSingleNode("/ApiResponse/ResultUrl").Text
    CheckContactsResult
End Sub
