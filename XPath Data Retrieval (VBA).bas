Attribute VB_Name = "Main"

Public OAuthToken As String
Public OAuthUrl As String
Public ApiKey As String
Public ApiUrl As String
Public ContactsResultUrl As String
Public ContactFields As IXMLDOMSelection
Public ExcludedFields(200) As String
Public cn As New CellNames
Public NumberOfColumns As Integer
Public RefNumcolidx As Integer
Public RefRowidx As Integer
Public increCount As Double, c As Integer


Sub DownloadButton_Click_MembershipData()
'    Exit Sub
    Application.ScreenUpdating = False
' On Error GoTo ErrorHandler
    SaveExcludedFields
    'Reset percentage in Progress Bar
    increCount = 0
    ProgressBar.Show vbModeless
    ProgressBar.Repaint
  ' clear immediate window
  Debug.Print String(65535, vbCr)
    
    ClearResultCells
    
    SettingsSheet.Range(cn.outLoadingState) = "Downloading Membership Data"
    
    ContactsResultUrl = ""
   
    ApiKey = CleanTrim(SettingsSheet.Range(cn.inApiKey))
    Debug.Print "API Key: " & ApiKey
    ApiUrl = CleanTrim(SettingsSheet.Range(cn.inApiUrl))
    Debug.Print "API URL: " & ApiUrl

    OAuthUrl = CleanTrim(SettingsSheet.Range(cn.inOAuthUrl))
    Debug.Print "API OAuthURL: " & OAuthUrl
    
    Debug.Print ("authorization")
    OAuthToken = GetOAUthToken(OAuthUrl)
    
    Debug.Print ("start downloading membership process")
    
    Set xmlDoc = LoadXml(ApiUrl)
    Set apiVersion = xmlDoc.SelectSingleNode("//ApiVersion/Version")
    SettingsSheet.Range(cn.outApiVersion) = apiVersion.Text
     
    Set versionUrl = xmlDoc.SelectSingleNode("//ApiVersion/Url")
    accountUrl = LoadAccountUrl(versionUrl.Text)
    ' fieldsurl = LoadXml(LoadAccountUrl(xmlDoc.SelectSingleNode("//ApiVersion/Url").Text)).SelectSingleNode("//Resources/Resource[Name='Contact fields']/Url").Text
    Set accountInfoDoc = LoadXml(accountUrl)
    SettingsSheet.Range(cn.outAccountName) = accountInfoDoc.SelectSingleNode("//Name").Text
    SettingsSheet.Range(cn.outDomainName) = "http://" + accountInfoDoc.SelectSingleNode("//PrimaryDomainName").Text
    
    fieldsurl = accountInfoDoc.SelectSingleNode("//Resources/Resource[Name='Contact fields']/Url").Text

    LoadFields fieldsurl
    levelsUrl = accountInfoDoc.SelectSingleNode("//Resources/Resource[Name='Membership levels']/Url").Text
    LoadLevels levelsUrl
    
    contactsBaseUrl = accountInfoDoc.SelectSingleNode("//Resources/Resource[Name='Contacts']/Url").Text
    LoadContacts contactsBaseUrl
    ApiKey = vbNullString
    ApiUrl = vbNullString
    OAuthUrl = vbNullString
    OAuthToken = vbNullString
    Unload ProgressBar
    Exit Sub
    Application.ScreenUpdating = True
'ErrorHandler:
'    SettingsSheet.Range(cn.outLoadingState) = "Failed"
'    SettingsSheet.Shapes("LoadDataButton").Visible = msoTrue
'    MsgBox ("Something goes wrong:" + Err.Description)
 
End Sub

Sub DownloadButton_Click_AuditLogItems()
' On Error GoTo ErrorHandler
  ' clear immediate window
    increCount = 0
    ProgressBar.Show vbModeless
    ProgressBar.Repaint
    Debug.Print String(65535, vbCr)

    ClearResultCells

    SettingsSheet.Range(cn.outLoadingState) = "Downloading AuditLog Data"

    ApiKey = CleanTrim(SettingsSheet.Range(cn.inApiKey))
    Debug.Print "API Key: " & ApiKey
    ApiUrl = CleanTrim(SettingsSheet.Range(cn.inApiUrl))
    Debug.Print "API URL: " & ApiUrl
    OAuthUrl = CleanTrim(SettingsSheet.Range(cn.inOAuthUrl))
    Debug.Print "API OAuthURL: " & OAuthUrl

    Debug.Print ("authorization")
    OAuthToken = GetOAUthToken(OAuthUrl)
    
    Debug.Print ("start downloading auditlog process")

    Set xmlDoc = LoadXml(ApiUrl)
    Set apiVersion = xmlDoc.SelectSingleNode("//ApiVersion/Version")
    SettingsSheet.Range(cn.outApiVersion) = apiVersion.Text

    Set versionUrl = xmlDoc.SelectSingleNode("//ApiVersion/Url")
    accountUrl = LoadAccountUrl(versionUrl.Text)

    Set accountInfoDoc = LoadXml(accountUrl)
    SettingsSheet.Range(cn.outAccountName) = accountInfoDoc.SelectSingleNode("//Name").Text
    SettingsSheet.Range(cn.outDomainName) = "http://" + accountInfoDoc.SelectSingleNode("//PrimaryDomainName").Text

    auditsBaseUrl = accountUrl + "/" + accountInfoDoc.SelectSingleNode("//Id").Text + "/AuditLogItems/"
    LoadAuditLogItems auditsBaseUrl
    Application.Wait (Now + TimeValue("0:00:01") / 5)
    LoadProperties auditsBaseUrl
    
    ' ClearCellWarning

    ApiKey = vbNullString
    ApiUrl = vbNullString
    OAuthUrl = vbNullString
    OAuthToken = vbNullString
    Unload ProgressBar
    Exit Sub
'ErrorHandler:
 '   SettingsSheet.Range(cn.outLoadingState) = "Failed"
  '  SettingsSheet.Shapes("LoadDataButton").Visible = msoTrue
   ' MsgBox ("Something goes wrong:" + Err.Description)

End Sub
Sub DownloadButton_Click_RenewalData()
    Application.ScreenUpdating = False
' On Error GoTo ErrorHandler
    SaveExcludedFields
    'Reset percentage in Progress Bar
    increCount = 0
    c = 0
    ProgressBar.Show vbModeless
    ProgressBar.Repaint
  ' clear immediate window
  Debug.Print String(65535, vbCr)
    
    'ClearResultCells
    
    SettingsSheet.Range(cn.outLoadingState) = "Downloading Renewal Contacts Data"
    
    ContactsResultUrl = ""
   
    ApiKey = CleanTrim(SettingsSheet.Range(cn.inApiKey))
    Debug.Print "API Key: " & ApiKey
    ApiUrl = CleanTrim(SettingsSheet.Range(cn.inApiUrl))
    Debug.Print "API URL: " & ApiUrl

    OAuthUrl = CleanTrim(SettingsSheet.Range(cn.inOAuthUrl))
    Debug.Print "API OAuthURL: " & OAuthUrl
    
    Debug.Print ("authorization")
    OAuthToken = GetOAUthToken(OAuthUrl)
    
    Debug.Print ("start downloading renewal contacts process")
    
    Set xmlDoc = LoadXml(ApiUrl)
    Set apiVersion = xmlDoc.SelectSingleNode("//ApiVersion/Version")
    SettingsSheet.Range(cn.outApiVersion) = apiVersion.Text
     '
    Set versionUrl = xmlDoc.SelectSingleNode("//ApiVersion/Url")
    accountUrl = LoadAccountUrl(versionUrl.Text)

    Set accountInfoDoc = LoadXml(accountUrl)
    SettingsSheet.Range(cn.outAccountName) = accountInfoDoc.SelectSingleNode("//Name").Text
    SettingsSheet.Range(cn.outDomainName) = "http://" + accountInfoDoc.SelectSingleNode("//PrimaryDomainName").Text
    
    fieldsurl = accountInfoDoc.SelectSingleNode("//Resources/Resource[Name='Contact fields']/Url").Text
    
    Set fieldsDoc = LoadXml(fieldsurl)
    Set ContactFields = fieldsDoc.SelectNodes("//ContactFieldDescription")
    Set versionUrl = xmlDoc.SelectSingleNode("//ApiVersion/Url")
    
    accountUrl = LoadAccountUrl(versionUrl.Text)

    Set accountInfoDoc = LoadXml(accountUrl)
    SettingsSheet.Range(cn.outAccountName) = accountInfoDoc.SelectSingleNode("//Name").Text
    SettingsSheet.Range(cn.outDomainName) = "http://" + accountInfoDoc.SelectSingleNode("//PrimaryDomainName").Text
    contactsBaseUrl = accountInfoDoc.SelectSingleNode("//Resources/Resource[Name='Contacts']/Url").Text
    ' Load Renewal Contacts
    LoadRenewalContacts contactsBaseUrl
    ApiKey = vbNullString
    ApiUrl = vbNullString
    OAuthUrl = vbNullString
    OAuthToken = vbNullString
    Unload ProgressBar
    Exit Sub
    Application.ScreenUpdating = True
    

End Sub

Sub ClearCellWarning()

    Dim events As Range
    Dim frowIdx As Integer
    frowIdx = 2
    Dim CellName As String

    CellName = Range(Cells(2, RefNumcolidx), Cells(RefRowidx, RefNumcolidx)).Address
  
    For Each events In AuditPropertiesSheet.Range(CellName)

       If events.Errors.Item(xlNumberAsText).Value = True Then
          events.Errors.Item(xlNumberAsText).Ignore = True
       End If
       frowIdx = frowIdx + 1
       
    Next events

'    SettingsSheet.Range(cn.outDataLoadedFinishedDate) = DateTime.Now
'    SettingsSheet.Range(cn.outLoadingState) = "Audit Log Complete"

End Sub

Sub SaveExcludedFields()
    Dim arrayIdx As Integer, rowIdx As Integer
    
    For arrayIdx = 1 To 200
        ExcludedFields(arrayIdx) = ""
    Next
    arrayIdx = 1
    
    For rowIdx = 2 To 200
        fieldName = FieldsSheet.Cells(rowIdx, 1).Text
        If Not (FieldsSheet.Cells(rowIdx, 6).Text = "Yes" Or fieldName = "") Then
            ExcludedFields(arrayIdx) = fieldName
            arrayIdx = arrayIdx + 1
        End If
    Next rowIdx
End Sub

Sub ClearResultCells()
    SettingsSheet.Range(cn.outGeneralInfo).ClearContents
    SettingsSheet.Range(cn.outLoadingStateInfo).ClearContents
    
    LevelsSheet.Range(cn.outLevelsList).ClearContents
    LevelsSheet.Range(cn.outLevelsList).WrapText = True
    
    FieldsSheet.Range(cn.outFieldsList).ClearContents
    FieldsSheet.Range(cn.outFieldsList).WrapText = True
    
    ContactsSheet.Range(cn.outContactHeaders).ClearContents
    ContactsSheet.Range(cn.outContactHeaders).WrapText = True
    ContactsSheet.Range(cn.outContactsList).Delete
    
    AuditPropertiesSheet.Range("out_AuditPropertiesHeaders").ClearContents
    
    AuditPropertiesSheet.Range(cn.outAuditPropertiesList).ClearContents
    AuditPropertiesSheet.Range(cn.outAuditPropertiesList).WrapText = False
    AuditPropertiesSheet.Range(cn.outAuditPropertiesList).Delete
   
End Sub

Sub LoadContacts(ByVal baseUrl As String)
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

Function BuildSelectExpression()

    Dim result As String, contactFieldItem As IXMLDOMElement, fieldName As String, incre As Integer, count As Integer, c As Integer
    For Each contactFieldItem In ContactFields
        fieldName = contactFieldItem.SelectSingleNode("FieldName").Text
        If Not IsFieldExcluded(fieldName) Then
            count = count + 1
        End If
    Next contactFieldItem
    If (count / 33) < 1 Then
        incre = 1
    Else
        incre = count / 33
    End If
    
    For Each contactFieldItem In ContactFields
        fieldName = contactFieldItem.SelectSingleNode("FieldName").Text
        If Not IsFieldExcluded(fieldName) Then
            If Len(result) > 0 Then
                result = result + ",'" + fieldName + "'"
            Else
                result = "'" + fieldName + "'"
            End If
            If c Mod incre = 0 Then
                increCount = increCount + 1
                progress CInt(increCount)
            End If
            c = c + 1
        End If
    Next contactFieldItem
    
    If Len(result) > 0 Then
        BuildSelectExpression = "$select=" + result
    Else
        BuildSelectExpression = ""
    End If
End Function

Sub CheckContactsResult()
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

Sub PrintContacts(contacts As IXMLDOMSelection)
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

Function GetCellFormat(ByVal fieldType As String) As String
    If fieldType = "String" Then
        GetCellFormat = "@"
    Else
        GetCellFormat = "General"
    End If
End Function

Sub LoadAuditLogItems(ByVal baseUrl As String)
    Dim url As String
    Dim auditfilter As String
    
    auditfilter = BuildAuditFilterExpression
    url = baseUrl
    If Len(auditfilter) > 0 Then
        url = url + "?" + auditfilter
    End If
    
    If Len(auditselectExpression) > 0 Then
        If Len(auditfilter) > 0 Then
            url = url + "&" + auditselectExpression
        Else
             url = url + "?" + auditselectExpression
        End If
    End If
    
    Dim auditsRequestDoc As IXMLDOMDocument
    Set auditsRequestDoc = LoadXml(url)
    
    Dim audits As IXMLDOMSelection
    Set audits = auditsRequestDoc.SelectNodes("/ApiResponse/Items/AuditItem")
    
    SettingsSheet.Range(cn.outDataLoadedDate) = DateTime.Now
    PrintAudits audits

End Sub

Sub LoadProperties(ByVal auditsBaseUrl As String)
    Dim cellinfo As Range
    Dim setid As String
    rowIdx = 1
    Dim rowC As Integer, c As Single, incre2 As Integer
    'Get total number of items by audit log id
If Not IsEmpty(AuditPropertiesSheet.Cells(2, 1)) Then 'Check if empty
    For Each cellinfo In AuditPropertiesSheet.Range(cn.inAuditlogIdList)
         rowC = rowC + 1
    Next cellinfo
    
    incre2 = rowC / 50
    'Avoid division by zero error when doing modular
    If incre2 < 1 Then
        incre2 = 1
    End If
    
    For Each cellinfo In AuditPropertiesSheet.Range(cn.inAuditlogIdList)
        setid = cellinfo.Value
        AuditIdurl = auditsBaseUrl + setid
        
        'the divisor slows down the calls to ensure 200 calls per minute are not exceeded (Wild Apricot API rate limit)
        
        Application.Wait [Now()] + TimeValue("00:00:01") / 5
        
        Dim auditpropertyRequestDoc As DOMDocument
        Set auditpropertyRequestDoc = LoadXml(AuditIdurl)
       
        Dim colIdx As Integer
        'colIdx = 13
        colIdx = 14
            Dim PropertiesNodes As IXMLDOMSelection
            Dim PropertiesLayer As IXMLDOMElement
                
            Set PropertiesNodes = auditpropertyRequestDoc.SelectNodes("/AuditItem/Properties/d2p1:KeyValueOfstringstring")
                For Each PropertiesLayer In PropertiesNodes
                   Dim localFieldKey As IXMLDOMElement
                   Dim localFieldValue As IXMLDOMElement
                
                   Set localFieldKey = PropertiesLayer.SelectSingleNode("d2p1:Key")
                   Set localFieldValue = PropertiesLayer.SelectSingleNode("d2p1:Value")
                   
                   RefNumcolidx = colIdx
                   AuditPropertiesSheet.Range(cn.outAuditPropertiesList)(1, colIdx).Offset(-1, 0) = localFieldKey.Text
                   'AuditPropertiesSheet.Columns(colIdx).EntireColumn.AutoFit
                   
                   If Len(localFieldValue.Text) > 15 Then
                      AuditPropertiesSheet.Columns(colIdx).NumberFormat = "@"
                      AuditPropertiesSheet.Range(cn.outAuditPropertiesList)(rowIdx, colIdx) = localFieldValue.Text
                      AuditPropertiesSheet.Columns(colIdx).EntireColumn.AutoFit
                   Else
                      AuditPropertiesSheet.Range(cn.outAuditPropertiesList)(rowIdx, colIdx) = localFieldValue.Text
                      AuditPropertiesSheet.Columns(colIdx).EntireColumn.AutoFit
                   End If
                                
                   colIdx = colIdx + 1
                Next PropertiesLayer
            If rowIdx Mod incre2 = 0 Then
               If increCount < 100 Then
                  increCount = increCount + 1
                  progress CInt(increCount)
               End If
            End If
    rowIdx = rowIdx + 1
Next cellinfo

RefRowidx = rowIdx

Dim lngNumberOfCharacters As Long

fieldinfo = AuditPropertiesSheet.Cells(1, 14).Text
lngNumberOfCharacters = Len(fieldinfo)

If lngNumberOfCharacters = 0 Then

 Dim newcolidx As Integer
 newcolidx = colIdx - 1

 SampleAuditLogSheet.PivotTables(1).SourceData = "AuditProperties!R1C1:R" + LTrim(Str(RefRowidx)) + "C" + LTrim(Str(newcolidx))
 SampleAuditLogSheet.PivotTables(1).RefreshTable
Else
 SampleAuditLogSheet.PivotTables(1).SourceData = "AuditProperties!R1C1:R" + LTrim(Str(RefRowidx)) + "C" + LTrim(Str(RefNumcolidx))
 SampleAuditLogSheet.PivotTables(1).RefreshTable
End If

End If 'End Emtpy Check if statement
    
    SettingsSheet.Range(cn.outDataLoadedFinishedDate) = DateTime.Now
    SettingsSheet.Range(cn.outLoadingState) = "Audit Log Complete"
    ' Finish up progress
    progress (100)
End Sub

Function BuildFilterExpression() As String
    Dim result As String
    
    If HasCriteria(cn.inLastUpdatedCriteria) Then
        compareOperator = GetOperator(cn.inLastUpdatedCriteria)
        valueToCompare = GetDate(cn.inLastUpdatedCriteria)
        expression = ComposeFilterExpression("Profile last updated", compareOperator, Format(valueToCompare, "yyyy-MM-ddTHH:mm:ss"))
        result = AppendToFilter(result, expression)
    End If

    If HasCriteria(cn.inRenewalDateCriteria) Then
        compareOperator = GetOperator(cn.inRenewalDateCriteria)
        valueToCompare = GetDate(cn.inRenewalDateCriteria)
        expression = ComposeFilterExpression("Renewal due", compareOperator, Format(valueToCompare, "yyyy-MM-ddTHH:mm:ss"))
        result = AppendToFilter(result, expression)
    End If
    
    If HasCriteria(cn.inMembershipStatusCriteria) Then
        compareOperator = GetOperator(cn.inMembershipStatusCriteria)
        valueToCompare = SettingsSheet.Range(cn.inMembershipStatusCriteria)(1, 2)
        expression = ComposeFilterExpression("Membership Status", compareOperator, valueToCompare)
        result = AppendToFilter(result, expression)
    End If
    
    If HasCriteria(cn.inMembershipEnabledCriteria) Then
        compareOperator = GetOperator(cn.inMembershipEnabledCriteria)
        valueToCompare = GetBool(cn.inMembershipEnabledCriteria)
        expression = ComposeFilterExpression("Member", compareOperator, valueToCompare)
        result = AppendToFilter(result, expression)
    End If
    
    If HasCriteria(cn.inIsArchivedCriteria) Then
        compareOperator = GetOperator(cn.inIsArchivedCriteria)
        valueToCompare = GetBool(cn.inIsArchivedCriteria)
        expression = ComposeFilterExpression("Archived", compareOperator, valueToCompare)
        result = AppendToFilter(result, expression)
    End If
    
    Dim IsArchivedExpressiohn As String
        
    If Len(result) > 0 Then
        BuildFilterExpression = "$filter=" + result
    Else
        BuildFilterExpression = ""
    End If
    
End Function

Function BuildAuditFilterExpression() As String
' Uses the AuditLog filtering criteria from Settings sheet to build dynamic date url

    Dim auditresult As String
    Dim auditexpression As String
    
    If HasCriteria(cn.inAuditlogStartDateCriteria) Then
        compareOperator = GetOperator(cn.inAuditlogStartDateCriteria)
        valueToCompare = GetDate(cn.inAuditlogStartDateCriteria)
        expression = ComposeFilterExpression("StartDate", compareOperator, Format(valueToCompare, "yyyy-mm-dd"))
        auditexpressionstart = "StartDate=" & Format(valueToCompare, "yyyy-mm-dd")
        auditresult = AppendToFilter(auditresult, expression)
     End If
  
    If HasCriteria(cn.inAuditlogEndDateCriteria) Then
       compareOperator = GetOperator(cn.inAuditlogEndDateCriteria)
       valueToCompare = GetDate(cn.inAuditlogEndDateCriteria)
       expression = ComposeFilterExpression("EndDate", compareOperator, Format(valueToCompare, "yyyy-mm-dd"))
       auditexpressionend = "EndDate=" & DateAdd("d", 1, Format(valueToCompare, "yyyy-mm-dd"))
      auditresult = AppendToFilter(auditresult, expression)
    End If
    
    auditexpression = auditexpressionstart + "&" + auditexpressionend
        
    If Len(auditresult) > 0 Then
        BuildAuditFilterExpression = auditexpression
    Else
        BuildAuditFilterExpression = ""
    End If

End Function

Function HasCriteria(ByVal rangeName As String) As Boolean
' Check if specified row contains valid filtering criteria
     criteriaRange = SettingsSheet.Range(rangeName)
     HasCriteria = Not (criteriaRange(1, 1) = "" Or criteriaRange(1, 2) = "")
End Function

Function ComposeFilterExpression(ByVal fieldName As String, ByVal comparisonOperator As String, ByVal valueToCompare As String) As String
    ComposeFilterExpression = "'" + fieldName + "' " + comparisonOperator + " '" + valueToCompare + "'"
End Function

Function GetOperator(ByVal rangeName As String) As String
    criteriaRange = SettingsSheet.Range(rangeName)
    Dim operatorText As String
    operatorText = criteriaRange(1, 1)
    If operatorText = "greater or equals" Then
        GetOperator = "ge"
    ElseIf operatorText = "less or equals" Then
        GetOperator = "le"
    ElseIf operatorText = "equals" Then
        GetOperator = "eq"
    ElseIf operatorText = "not equals" Then
        GetOperator = "ne"
    Else
        Err.Raise -1, "API sample for VBA", "unexpected comparison operator"
    End If
End Function

Function GetDate(ByVal rangeName As String) As Date
    criteriaRange = SettingsSheet.Range(rangeName)
    GetDate = DateValue(criteriaRange(1, 2))
End Function

Function GetBool(ByVal rangeName As String) As Boolean
    criteriaRange = SettingsSheet.Range(rangeName)
    strText = criteriaRange(1, 2)
    If strText = "Yes" Then
        GetBool = True
    ElseIf strText = "No" Then
        GetBool = False
    Else
        Err.Raise -1, "API sample for VBA", "Can't convert to bool"
    End If
End Function

Function AppendToFilter(ByVal curFilter As String, ByVal newExpression As String) As String
    If Len(newExpression) = 0 Then
        AppendToFilter = curFilter
        Exit Function
    End If
    
    If Len(curFilter) > 0 Then
        AppendToFilter = curFilter + " and " + newExpression
    Else
        AppendToFilter = newExpression
    End If
    
End Function

Sub PrintAudits(audits As IXMLDOMSelection)
    Dim auditproperty As IXMLDOMElement
    Dim curRow As Integer, numRow As Integer
    Dim percentComp As Double, incre As Double
    curRow = 1
    For Each auditproperty In audits
        numRow = numRow + 1
    Next auditproperty
    incre = numRow / 50
    'Avoid division by zero error
    If incre < 1 Then
        incre = 1
    End If
    
    For Each auditproperty In audits
        PrintAudit auditproperty, curRow
        If curRow Mod incre = 0 Then
            increCount = increCount + 1
            progress CInt(increCount)
        End If
        curRow = curRow + 1
    Next auditproperty

End Sub

Sub PrintAudit(auditproperty As IXMLDOMElement, rowIdx As Integer)
    AuditPropertiesSheet.Range(cn.outAuditPropertiesList)(rowIdx, 1) = auditproperty.SelectSingleNode("Id").Text
    AuditPropertiesSheet.Range(cn.outAuditPropertiesList)(rowIdx, 2) = auditproperty.SelectSingleNode("Url").Text
    AuditPropertiesSheet.Range(cn.outAuditPropertiesList)(rowIdx, 3) = auditproperty.SelectSingleNode("Timestamp").Text
    AuditPropertiesSheet.Range(cn.outAuditPropertiesList)(rowIdx, 4) = auditproperty.SelectSingleNode("Contact/Id").Text
    AuditPropertiesSheet.Range(cn.outAuditPropertiesList)(rowIdx, 5) = auditproperty.SelectSingleNode("Contact/Url").Text
    AuditPropertiesSheet.Range(cn.outAuditPropertiesList)(rowIdx, 6) = auditproperty.SelectSingleNode("FirstName").Text
    AuditPropertiesSheet.Range(cn.outAuditPropertiesList)(rowIdx, 7) = auditproperty.SelectSingleNode("LastName").Text
    AuditPropertiesSheet.Range(cn.outAuditPropertiesList)(rowIdx, 8) = auditproperty.SelectSingleNode("Email").Text
    AuditPropertiesSheet.Range(cn.outAuditPropertiesList)(rowIdx, 9) = auditproperty.SelectSingleNode("Organization").Text
    AuditPropertiesSheet.Range(cn.outAuditPropertiesList)(rowIdx, 10) = auditproperty.SelectSingleNode("Message").Text
    AuditPropertiesSheet.Range(cn.outAuditPropertiesList)(rowIdx, 11) = auditproperty.SelectSingleNode("Severity").Text
    AuditPropertiesSheet.Range(cn.outAuditPropertiesList)(rowIdx, 12) = auditproperty.SelectSingleNode("OrderType").Text
   ' AuditPropertiesSheet.Range(cn.outAuditPropertiesList)(rowIdx, 13) = auditproperty.SelectSingleNode("OrderType").Text
    Dim colIdx As Integer
    colIdx = 2

End Sub

Sub LoadLevels(ByVal url As String)
    Set levelsDoc = LoadXml(url)
    Set Levels = levelsDoc.SelectNodes("//MembershipLevel")
    Dim levelElement As IXMLDOMElement
    curRowIdx = 2
    Dim levelCount As Double, incre As Double
    For Each levelElement In Levels
        levelCount = levelCount + 1
    Next levelElement
    'estimation of increment
    incre = levelCount / 33
    If incre < 1 Then
        incre = 1
    End If
    For Each levelElement In Levels
        LevelsSheet.Cells(curRowIdx, 1) = levelElement.SelectSingleNode("Id").Text
        LevelsSheet.Cells(curRowIdx, 2) = levelElement.SelectSingleNode("Name").Text
        LevelsSheet.Cells(curRowIdx, 3) = levelElement.SelectSingleNode("Description").Text
        LevelsSheet.Cells(curRowIdx, 4) = levelElement.SelectSingleNode("PublicCanApply").Text
        LevelsSheet.Cells(curRowIdx, 5) = levelElement.SelectSingleNode("Type").Text
        LevelsSheet.Cells(curRowIdx, 6) = levelElement.SelectSingleNode("MembershipFee").Text
        
        If (curRowIdx - 1) Mod incre = 0 Then
            increCount = increCount + 1
            progress CInt(increCount)
        End If
        
        curRowIdx = curRowIdx + 1
    Next levelElement
End Sub

Sub LoadFields(ByVal url As String)
    Set fieldsDoc = LoadXml(url)
    Set ContactFields = fieldsDoc.SelectNodes("//ContactFieldDescription")
    Dim membershipLevelRequested As Boolean
    Dim fieldElement As IXMLDOMElement, contactColumnIdx As Integer, contactCount As Integer, incre As Integer
    contactColumnIdx = 2
    curRowIdx = 2
    membershipLevelRequested = True
    
    For Each fieldElement In ContactFields
        Dim filedName1 As String
        fieldName1 = fieldElement.SelectSingleNode("FieldName").Text
        If Not (fieldName1 = "Registred for specific event") Then
            Set Labels = fieldElement.SelectNodes("AllowedValues/ListItem/Label")
            For Each Label In Labels
                contactCount = contactCount + 1
            Next Label
        End If
    Next fieldElement
    ' estimate increment
    If contactCount < 33 Then
        incre = 1 'Avoid division by zero error
    Else
        incre = contactCount / 33
    End If
    
    For Each fieldElement In ContactFields
        Dim filedName As String
        fieldName = fieldElement.SelectSingleNode("FieldName").Text
        
        If Not (fieldName = "Registred for specific event") Then
            FieldsSheet.Cells(curRowIdx, 1) = fieldName
            FieldsSheet.Cells(curRowIdx, 2) = fieldElement.SelectSingleNode("Type").Text
            FieldsSheet.Cells(curRowIdx, 3) = fieldElement.SelectSingleNode("Description").Text
            FieldsSheet.Cells(curRowIdx, 4) = fieldElement.SelectSingleNode("FieldInstructions ").Text
            FieldsSheet.Cells(curRowIdx, 4) = fieldElement.SelectSingleNode("FieldInstructions ").Text
            
            Set Labels = fieldElement.SelectNodes("AllowedValues/ListItem/Label")
            
            Dim allowedVals As String
            allowedVals = ""
            For Each Label In Labels
                allowedVals = allowedVals + ", " + Label.Text
            Next Label
            If Len(allowedVals) > 2 Then
                allowedVals = Mid(allowedVals, 3, Len(allowedVals) - 2)
            End If
            FieldsSheet.Cells(curRowIdx, 5) = allowedVals
            
            
            If Not IsFieldExcluded(fieldName) Then
                ContactsSheet.Cells(1, contactColumnIdx) = FieldsSheet.Cells(curRowIdx, 1)
                fieldType = fieldElement.SelectSingleNode("Type").Text
                If fieldType = "Boolean" Or fieldType = "Number" Or fieldType = "DateTime" Then
                    ContactsSheet.Columns(contactColumnIdx).ColumnWidth = 15
                Else
                    ContactsSheet.Columns(contactColumnIdx).ColumnWidth = 35
                End If
                ContactsSheet.Columns(contactColumnIdx).WrapText = True
                contactColumnIdx = contactColumnIdx + 1
                
                FieldsSheet.Cells(curRowIdx, 6) = "Yes"
                
                If LCase(fieldName) = LCase("Membership level ID") Then
                    membershipLevelRequested = True
                End If
                
            Else
                FieldsSheet.Cells(curRowIdx, 6) = "No"
            End If
            'Progress Bar percentage increases
            If (curRowIdx - 1) Mod incre = 0 Then
                increCount = increCount + 1
                progress CInt(increCount)
            End If
            curRowIdx = curRowIdx + 1
        End If
        
    Next fieldElement
    
    If membershipLevelRequested Then
        ContactsSheet.Cells(1, contactColumnIdx) = "Membership Level"
        ContactsSheet.Columns(contactColumnIdx).ColumnWidth = 35
    End If
    
    NumberOfColumns = contactColumnIdx
    
End Sub

Function IsFieldExcluded(ByVal fieldName As String) As Boolean

    If fieldName = "Registred for specific event" Then
        IsFieldExcluded = True
        Exit Function
    End If
    
    Dim itemIdx As Integer
    For itemIdx = 1 To 100
        If ExcludedFields(itemIdx) = fieldName Then
            IsFieldExcluded = True
            Exit Function
        End If
    Next itemIdx
    IsFieldExcluded = False
End Function

Function EncodeBase64(Text As String) As String
  Dim arrData() As Byte
  arrData = StrConv(Text, vbFromUnicode)

  Dim objXML As MSXML2.DOMDocument
  Dim objNode As MSXML2.IXMLDOMElement

  Set objXML = New MSXML2.DOMDocument
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  EncodeBase64 = objNode.Text

  Set objNode = Nothing
  Set objXML = Nothing
End Function

Sub SetOAuthCredentials(httpClient As IXMLHTTPRequest)
    httpClient.setRequestHeader "User-Agent", "VBA sample app" ' This header is optional, it tells what application is working with API
    httpClient.setRequestHeader "Authorization", "Basic " + EncodeBase64("APIKEY:" + ApiKey) ' This header is required, it provides API key for authentication
    httpClient.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
End Sub

Function GetOAUthToken(ByVal url As String) As String
    Debug.Print ("Loading data from " + url)
    Dim httpClient As IXMLHTTPRequest
    Set httpClient = CreateObject("Msxml2.XMLHTTP.3.0")
    httpClient.Open "POST", url, False
    SetOAuthCredentials httpClient
    
    httpClient.send ("grant_type=client_credentials&scope=auto")
    
    If Not httpClient.Status = 200 Then
        msg = "Call to " + url + " returned error:" + httpClient.statusText
        Err.Raise -1, "GetOAUthToken", msg
    End If
    
    Dim resp As String
    resp = httpClient.responseText
    resp = Mid(resp, Len("{""access_token"":""") + 1, InStr(resp, """,""token_type""") - Len("{""access_token"":""") - 1)

    GetOAUthToken = resp
End Function

Sub SetCredentials(httpClient As IXMLHTTPRequest)
    httpClient.setRequestHeader "User-Agent", "VBA sample app" ' This header is optional, it tells what application is working with API
    httpClient.setRequestHeader "Authorization", "Bearer " + OAuthToken  ' This header is required, it provides API key for authentication
    httpClient.setRequestHeader "Accept", "application/xml" ' This header is required, it tells to return data in XML format
End Sub

Function LoadXml(ByVal url As String) As DOMDocument
    Debug.Print ("Loading data from " + url + " " + Format(Now, "Hh:Nn:ss") + "." + _
    Right(Format(Timer, "#0.00"), 2)) + " " + Format(Now, "AM/PM dd-mmm-yyyy")
    
    Dim httpClient As IXMLHTTPRequest
    Set httpClient = CreateObject("Msxml2.XMLHTTP.3.0")
    httpClient.Open "GET", url, False
    Application.Wait (Now + TimeValue("0:00:01") / 5)
    SetCredentials httpClient
    Application.Wait (Now + TimeValue("0:00:01") / 2)
    httpClient.send
    Application.Wait (Now + TimeValue("0:00:01"))
    If Not httpClient.Status = 200 Then
        msg = "Call to " + url + vbNewLine + "returned error: " + httpClient.statusText + vbNewLine + _
        Format(Now, "Hh:Nn:ss") + "." + Right(Format(Timer, "#0.00"), 2) + " " + _
        Format(Now, "AM/PM dd-mmm-yyyy")
        
        Debug.Print msg
        Err.Raise -1, "LoadXML", msg
    End If

    Set xmlDoc = httpClient.responseXML
    Set LoadXml = xmlDoc
End Function

Function LoadAccountUrl(versionUrl As String) As String
    Set versionResourcesXml = LoadXml(versionUrl)
    Set accountsUrlNode = versionResourcesXml.SelectSingleNode("//Resource[Name='Accounts']/Url")
    LoadAccountUrl = accountsUrlNode.Text
End Function
Sub progress(percentCompleted As Single)
If percentCompleted >= 100 Then
    ProgressBar.Text.Caption = "99% Completed"
    ProgressBar.Bar.Width = percentCompleted * 2
Else
    ProgressBar.Text.Caption = percentCompleted & "% Completed"
    ProgressBar.Bar.Width = percentCompleted * 2
End If
DoEvents
End Sub


Sub CheckRenewalContactsResult()
    Dim complete As Boolean, requestState As String
    complete = False
    Dim contactsRequestDoc As DOMDocument
    Set contactsRequestDoc = LoadXml(ContactsResultUrl)
    requestState = contactsRequestDoc.SelectSingleNode("/ApiResponse/State").Text
    SettingsSheet.Range(cn.outLoadingState) = requestState
    SettingsSheet.Range(cn.outDataLoadedDate) = DateTime.Now
    If requestState = "Waiting" Or requestState = "Processing" Then
        Application.OnTime Now + TimeValue("00:00:01"), "CheckContactsResult"
    ElseIf requestState = "Complete" Then
        Dim contacts As IXMLDOMSelection
        Set contacts = contactsRequestDoc.SelectNodes("/ApiResponse/Contacts/Contact")
        'PrintContacts contacts
        PrintRenewalContacts contacts, 1
        complete = True
        SettingsSheet.Range(cn.outDataLoadedFinishedDate) = DateTime.Now
    ElseIf requestState = "Failed" Then
        Dim errMessage As String
        errMessage = contactsRequestDoc.SelectSingleNode("/ApiResponse/ErrorDetails").Text
        MsgBox ("Failed to load contacts. Error: " + errMessage)
        complete = True
    End If
End Sub

Sub LoadRenewalContacts(ByVal baseUrl As String)
    Dim url As String, filter As String, selectExpression As String
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
    CheckRenewalContactsResult
End Sub

Sub CheckRenewal(contact As IXMLDOMElement, rowIdx As Integer, DayRange As Integer)
    Dim EndRange As Date
    'temoporary set as within 20 day
    EndRange = DateAdd("d", 100, Format(Date, "yyyy-mm-dd"))
    Dim colIdx As Integer
    colIdx = 2
    Dim fieldItem As IXMLDOMElement
    For Each fieldItem In ContactFields
        fieldType = fieldItem.SelectSingleNode("Type").Text
        fieldName = fieldItem.SelectSingleNode("FieldName").Text
    If fieldName = "Renewal due" Then
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
                If fieldType = "DateTime" And Not valueNode.Text = "" Then
                        d = DateSerial(CInt(Mid(valueNode.Text, 1, 4)), CInt(Mid(valueNode.Text, 6, 2)), CInt(Mid(valueNode.Text, 9, 2)))
                        If d <= EndRange Then
                            'temporarily increment c for testing, switching to printing contact names in a spread sheet later
                            c = c + 1
                        End If
                End If
            End If
            colIdx = colIdx + 1
        End If
    End If
    Next fieldItem
    'Debugging Use
    If d <= EndRange Then
        Debug.Print "Increment Success!"
        Debug.Print "Printing the number c: "
        Debug.Print c
        Debug.Print "Printing Renewal Date: "
        Debug.Print d
    End If
End Sub

Sub PrintRenewalContacts(contacts As IXMLDOMSelection, DayRange As Integer)
    Debug.Print "Printing Renewal Contacts"
    Dim contact As IXMLDOMElement
    Dim curRow As Integer, contactCount As Integer, incre As Double
    curRow = 1
    For Each contact In contacts
        contactCount = contactCount + 1
    Next contact
    'Estimate increment
    incre = contactCount / 100
    'Avoid Divide by zero error when doing modular arithmetic
    If incre < 1 Then
        incre = 1
    End If
    
    For Each contact In contacts
        CheckRenewal contact, curRow, DayRange
        If curRow Mod incre = 0 Then
            increCount = increCount + 1
            progress CInt(increCount)
        End If
        curRow = curRow + 1
    Next contact
    
End Sub
