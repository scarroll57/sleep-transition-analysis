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
    Dim n2n1, n3n2, n3n1, Rn1, Rn2, Rn3, n1w, n2w, n3w, Rw As Long
    Dim tracker1, tracker2 As Long

    n2n1 = 0
    n3n2 = 0
    n3n1 = 0
    Rn1 = 0
    Rn2 = 0
    Rn3 = 0
    n1w = 0
    n2w = 0
    n3w = 0
    Rw = 0


    'Analyze each epoch set for transitions
    For x = 3 To FinalRow
        tracker1 = Cells(x - 1, 3).Value
        tracker2 = Cells(x, 3).Value

        If tracker2 < tracker1 Then
            'test to see which transition set to add the transition to
            'this should protect against adding unstaged transitions
            'test to see if transitioning from N2 to N1
            If tracker2 = 1 And tracker1 = 2 Then
                n2n1 = n2n1 + 1
            End If
            'test to see if transitioning from N3 to N2
            If tracker2 = 2 And tracker1 = 3 Then
                n3n2 = n3n2 + 1
            End If
            'test to see if transitioning from Rem to N3
            If tracker2 = 3 And tracker1 = 5 Then
                Rn3 = Rn3 + 1
            End If
            'test to see if transitioning from Rem to N2
            If tracker2 = 2 And tracker1 = 5 Then
                Rn2 = Rn2 + 1
            End If
            'test to see if transitioning from Rem to N1
            If tracker2 = 1 And tracker1 = 5 Then
                Rn1 = Rn1 + 1
            End If
            'test to see if transitioning from Rem to Wake
            If tracker2 = 0 And tracker1 = 5 Then
                Rw = Rw + 1
            End If
            'test to see if transitioning from N1 to Wake
            If tracker2 = 0 And tracker1 = 1 Then
                n1w = n1w + 1
            End If
            'test to see if transitioning from N2 to Wake
            If tracker2 = 0 And tracker1 = 2 Then
                n2w = n2w + 1
            End If
            'test to see if transitioning from N3 to wake
            If tracker2 = 0 And tracker1 = 3 Then
                n3w = n3w + 1
            End If

        End If

    Next x

    'Convert Total Sleep Time from minutes to hours
    TST_hour = Cells(2, 5).Value / 60

    'Print Hour Conversion on The Spreadsheet
    Range("E3").Value = "Total Sleep Time in Hours"
    Range("E4").Value = TST_hour

    'Print Values on the Spreadsheet
    Range("G1").Value = "N2 to N1"
    Range("H1").Value = "N3 to N2"
    Range("I1").Value = "N3 to N1"
    Range("J1").Value = "REM to N1"
    Range("K1").Value = "REM to N2"
    Range("L1").Value = "REM to N3"
    Range("M1").Value = "REM to Wake"
    Range("N1").Value = "N1 to Wake"
    Range("O1").Value = "N2 to Wake"
    Range("P1").Value = "N3 to Wake"

    Range("G2").Value = n2n1
    Range("H2").Value = n3n2
    Range("I2").Value = n3n1
    Range("J2").Value = Rn1
    Range("K2").Value = Rn2
    Range("L2").Value = Rn3
    Range("M2").Value = Rw
    Range("N2").Value = n1w
    Range("O2").Value = n2w
    Range("P2").Value = n3w

    'Math to Calculate Sleep State Transitions
    'Also print to spreadsheet
    Range("G4").Value = "Lightening of Sleep transitions"
    Range("G5").Value = n2n1 + n3n2 + Rn1 + Rn2 + Rn3 + Rw + n1w + n2w + n3w
    Range("H5").Value = Range("G5").Value / TST_hour

    Range("G6").Value = "REM to NREM transitions"
    Range("G7").Value = Rn1 + Rn2 + Rn3
    Range("H7").Value = Range("G7").Value / TST_hour

    Range("G8").Value = "NREM to lesser NREM transitions"
    Range("G9").Value = n2n1 + n3n2 + n3n1
    Range("H9").Value = Range("G9").Value / TST_hour

    Range("I4").Value = "Sleep to Wake transitions"
    Range("I5").Value = n1w + n2w + n3w + Rw
    Range("I5").Value = Range("I5").Value / TST_hour

    Range("I6").Value = "REM to Wake transitions"
    Range("I7").Value = Rw
    Range("I7").Value = Range("I7").Value / TST_hour

    Range("I8").Value = "NREM to Wake transitions"
    Range("I9").Value = n1w + n2w + n3w
    Range("I9").Value = Range("I9").Value / TST_hour

    Range("L4").Value = "The table beside this contains the total number of events and the index listed in Event, Index format"

End Sub

Function conversion(raw)
    'This function serves as a simplification of the code to convert the epoch staging to a numerical format
    'The If statement belos uses conditional logic to check each input value
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