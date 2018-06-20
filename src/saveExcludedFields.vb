Sub SaveExcludedFields()
' Use array to store excluded fields prompted by user, then use the information to retrieve necessary data
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
