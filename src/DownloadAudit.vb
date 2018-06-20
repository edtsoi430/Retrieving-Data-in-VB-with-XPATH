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
