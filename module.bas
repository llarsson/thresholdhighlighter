Attribute VB_Name = "ThresholdHighlighter"
Sub ThresholdHighlighter2000()
Dim samples As Range, thresholds As Range

Set thresholds = Application.InputBox("Select the thresholds", "Threshold highlighter 2000", Type:=8)
Set samples = Application.InputBox("Select the samples", "Threshold highlighter 2000", Type:=8)

If thresholds.Rows.Count <> samples.Rows.Count Then
    MsgBox ("Select all samples and thresholds to use, please!")
    Exit Sub
End If

' Where do we want our results to wind up? Below or to the side?
Dim resultRowDelta As Integer, resultColumnDelta As Integer
If MsgBox("Show results to the side (= Yes)? If not, they will be below (= No).", vbYesNo, "Results to the side or not?") = vbYes Then
    resultRowDelta = 0
    resultColumnDelta = samples.Columns.Count ' Place them this many columns to the side
Else
    resultRowDelta = samples.Rows.Count ' Place them this many rows below
    resultColumnDelta = 0
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
                sampleValue = CDbl(tokens(1) * 0.999999999)
            Else
                sampleValue = currentSample
            End If
            
            If sampleValue > currentThreshold Then
                ' Are we over the threshold because lab could not measure lower or not?
                If lessThanSignInSample Then
                    ' Actual level inconclusive, as lab could not measure lower
                    Debug.Print "OVER (INCONCLUSIVE) " & currentThreshold & " because " & currentSample & " at " & currentRow & "," & sampleColumn & " has value " & sampleValue
                    samples.Cells(currentRow, sampleColumn).Font.Color = vbRed
                    samples.Cells(currentRow, sampleColumn).Font.Bold = True
                    
                    samples.Cells(currentRow + resultRowDelta, sampleColumn + resultColumnDelta).Value = "Rapporteringsgräns > RV"
                Else
                    ' Lab was sure, and the sample is over
                    Debug.Print "OVER " & currentThreshold & " because " & currentSample & " at " & currentRow & "," & sampleColumn & " has value " & sampleValue
                    thresholds.Cells(currentRow, thresholdColumn).Copy
                    samples.Cells(currentRow, sampleColumn).PasteSpecial Paste:=xlPasteFormats
                    
                    samples.Cells(currentRow + resultRowDelta, sampleColumn + resultColumnDelta).PasteSpecial Paste:=xlPasteFormats
                    samples.Cells(currentRow + resultRowDelta, sampleColumn + resultColumnDelta).Value = sampleValue / currentThreshold
                    samples.Cells(currentRow + resultRowDelta, sampleColumn + resultColumnDelta).NumberFormat = "0.0"
                End If
            End If
            
NextIteration:
        Next thresholdColumn
    Next sampleColumn
Next currentRow


End Sub
