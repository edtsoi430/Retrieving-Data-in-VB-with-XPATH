Function BuildSelectExpression()
  ' Build Selected expressions to process necessary data
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
