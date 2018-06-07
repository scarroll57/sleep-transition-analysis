Attribute VB_Name = "sleep_transistion_analysis"
Sub sleep_transistion_analysis()

    'Select the spreadsheet where the data is located
    Worksheets("Sheet1").Activate

    'Find the last row of the data
    Dim FinalRow As Long
    FinalRow = ActiveSheet.Cells(Rows.Count, 2).End(xlUp).Row

    'The data should be kept in Column 2 or B
    'For each transition we will make use of a count function
    'Before using the count function we will convert the values to numerical formats for easy use

    'Create numerical conversion column
    Range("C1").Value = "Numerical Staging"
    For x = 2 To FinalRow
        Cells(x, 3).Value = conversion(Cells(x, 2).Value)
    Next x

    'Distingish variables
    Dim n2n1, n3n2, Rn1, Rn2, Rn3, n1w, Rw As Long
    Dim tracker1, tracker 2 As Long

    'Analyze each epoch set for transitions
    For x = 3 To FinalRow
        tracker1 = Cells(x - 1, 3).Value
        tracker2 = Cells(x, 3).Value



    Next x

End Sub

Function conversion(raw)

    If raw = "U" Then
        conversion = -1
    ElseIf raw = "W" Then
        conversion = 0
    ElseIf raw = "N1" Then
        conversion = 1
    ElseIf raw = "N2" Then
        conversion = 2
    ElseIf raw = "N3" Then
        conversion = 3
    ElseIf raw = "R" Then
        conversion = 5
    End If

End Function


Function count(stagefrom, stageto)

    If stageto > stagefrom Then

End Function