Sub ThresholdHighlighter2000()
Dim samples As Range, thresholds As Range

Set thresholds = Application.InputBox("Select the thresholds", "Threshold highlighter 2000", Type:=8)
Set samples = Application.InputBox("Select the samples", "Threshold highlighter 2000", Type:=8)

If thresholds.Rows.Count <> samples.Rows.Count Then
    MsgBox ("Select all samples and thresholds to use, please!")
    Exit Sub
End If

For currentRow = 1 To samples.Rows.Count
    For sampleColumn = 1 To samples.Columns.Count
        For thresholdColumn = 1 To thresholds.Columns.Count
            currentSample = samples.Cells(currentRow, sampleColumn)
            currentThreshold = thresholds.Cells(currentRow, thresholdColumn)
            
            If IsNull(currentSample) Or IsNull(currentThreshold) Then
                GoTo NextIteration
            End If
            
            Dim lessThanSignInSample As Boolean
            Dim sampleValue As Double
            lessThanSignInSample = InStr(currentSample, "<") <> 0
            If lessThanSignInSample Then
                tokens = Split(currentSample, "<")
                sampleValue = CDbl(tokens(1) * 0.999999)
            Else
                sampleValue = currentSample
            End If
            
            ' Debug.Print currentSample & " at " & currentRow & "," & sampleColumn & " has value " & sampleValue
            
            If sampleValue > currentThreshold Then
                ' Are we over the threshold because lab could not measure lower or not?
                If lessThanSignInSample Then
                    ' Lab could not measure lower
                    Debug.Print "OVER, BUT DUE TO LAB " & currentThreshold & " because " & currentSample & " at " & currentRow & "," & sampleColumn & " has value " & sampleValue
                    samples.Cells(currentRow, sampleColumn).Font.Color = vbRed
                    samples.Cells(currentRow, sampleColumn).Font.Bold = True
                Else
                    ' Lab was sure, and the sample is over
                    Debug.Print "OVER " & currentThreshold & " because " & currentSample & " at " & currentRow & "," & sampleColumn & " has value " & sampleValue
                    thresholds.Cells(currentRow, thresholdColumn).Copy
                    samples.Cells(currentRow, sampleColumn).PasteSpecial Paste:=xlPasteFormats
                End If
            End If
            
NextIteration:
        Next thresholdColumn
    Next sampleColumn
Next currentRow

End Sub

