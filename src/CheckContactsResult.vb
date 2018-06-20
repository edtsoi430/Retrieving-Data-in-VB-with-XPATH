Sub CheckContactsResult()
'Check if API state is finished and inform user to boost and improve user experience.
    Dim complete As Boolean, requestState As String
    complete = False
    Dim contactsRequestDoc As DOMDocument
    Set contactsRequestDoc = LoadXml(ContactsResultUrl)
    requestState = contactsRequestDoc.SelectSingleNode("/ApiResponse/State").Text
    SettingsSheet.Range(cn.outLoadingState) = requestState
    SettingsSheet.Range(cn.outDataLoadedDate) = DateTime.Now
    If requestState = "Waiting" Or requestState = "Processing" Then
        Application.OnTime Now + TimeValue("00:00:03"), "CheckContactsResult"
    ElseIf requestState = "Complete" Then
        Dim contacts As IXMLDOMSelection
        Set contacts = contactsRequestDoc.SelectNodes("/ApiResponse/Contacts/Contact")
        PrintContacts contacts
        complete = True
        SettingsSheet.Range(cn.outDataLoadedFinishedDate) = DateTime.Now
    ElseIf requestState = "Failed" Then
        Dim errMessage As String
        errMessage = contactsRequestDoc.SelectSingleNode("/ApiResponse/ErrorDetails").Text
        MsgBox ("Failed to load contacts. Error: " + errMessage)
        complete = True
    End If
End Sub
