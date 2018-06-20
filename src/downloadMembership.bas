Sub DownloadButton_Click_MembershipData()
' To speed up processing time
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
ErrorHandler:
    SettingsSheet.Range(cn.outLoadingState) = "Failed"
    SettingsSheet.Shapes("LoadDataButton").Visible = msoTrue
    MsgBox ("Something goes wrong:" + Err.Description)
 
End Sub
