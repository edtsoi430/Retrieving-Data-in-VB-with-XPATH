Sub LoadAuditLogItems(ByVal baseUrl As String)
'Function to load all audit log items from the database
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
