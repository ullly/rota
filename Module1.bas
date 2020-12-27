Attribute VB_Name = "Module1"
Sub Magic()
    
    Dim docs As Range
    Set docs = Range("F2:F6")
    
    Dim totalDocs As Integer
    totalDocs = docs.Rows.Count
    
    Dim days As Range
    Set days = Range("B4:B192")
    
    Dim totalDays As Integer
    totalDays = days.Rows.Count
    
    Dim sL As Range
    Set sL = Range("D4:D192")
    
    Dim aL As Range
    Set aL = Range("E4:E192")
    
    ' Assign the first doctor to the first day
    days.Rows(1).Value() = docs.Rows(1).Value()
    
    ' Assign a random doctor to the second day onwards
    Dim i As Integer
    For i = 2 To totalDays
    
        Dim assigned As Boolean
        assigned = False
        
        Do While assigned = False
            
            Dim doc As String
            
            ' Select a random doctor
            doc = docs.Rows(Int(1 + Rnd * docs.Rows.Count)).Value
            
            ' Assign the doctor if the following conditions are met:
            '   Doctor is not on sick leave on this day
            '   Doctor is not on annual leave on this day
            '   Doctor is not on call the previous day
            If assigned = False And _
                Check(sL.Rows(i).Value, doc) And _
                    Check(aL.Rows(i).Value, doc) And _
                        days.Rows(i - 1).Value <> doc Then
                
                days.Rows(i).Value = doc
                
                ' Successfully assigned:
                '   Set flag to true so we can move to the next day
                assigned = True
                
            End If
            
        Loop
        
    Next i
    
End Sub

Function Check(ByVal content As String, ByVal doc As String) As Boolean

    ' Returns False if doc is a substring of content
    If InStr(content, doc) > 0 Then
        Check = False
    Else
        Check = True
    End If

End Function
