Attribute VB_Name = "ThresholdHighlighter"
Sub ThresholdHighlighter2000()
Dim samples As Range, thresholds As Range

Set parameterNames = Application.InputBox("Select parameter names", "Threshold highlighter 2000", Type:=8)
Set thresholds = Application.InputBox("Select thresholds", "Threshold highlighter 2000", Type:=8)
Set samples = Application.InputBox("Select samples", "Threshold highlighter 2000", Type:=8)

If thresholds.Rows.Count <> samples.Rows.Count Or parameterNames.Rows.Count <> samples.Rows.Count Then
    MsgBox ("Different number of parameters, thresholds, or samples selected!")
    Exit Sub
End If

' Where do we want our results to wind up? Below or to the side?
Dim resultRowDelta As Integer, resultColumnDelta As Integer

'If MsgBox("Choose Yes for answers on the side, and No for answers below", vbYesNo, "Results on the side or not?") = vbYes Then
'    resultRowDelta = 0
'    resultColumnDelta = samples.Columns.Count ' Place them this many columns to the side'
'Else
'    resultRowDelta = samples.Rows.Count + 1 ' Place them this many rows below (add one row for extra space)
'    resultColumnDelta = 0
'End If

' Nope, always just put the results in a box below
Dim resultOffset As Integer
resultOffset = Application.InputBox("How many rows under the last threshold should the results be shown?", "Threshold highlighter 2000", Type:=1)
Debug.Print "resultOffset = " & resultOffset
resultRowDelta = samples.Rows.Count + resultOffset
resultColumnDelta = 0

For currentRow = 1 To samples.Rows.Count
    For sampleColumn = 1 To samples.Columns.Count
        For thresholdColumn = 1 To thresholds.Columns.Count
            Dim currentSample As String
            currentSample = samples.Cells(currentRow, sampleColumn)
            currentThreshold = thresholds.Cells(currentRow, thresholdColumn)
            
            If IsEmpty(currentThreshold) Or IsEmpty(currentSample) Or IsNull(currentSample) Or IsNull(currentThreshold) Or currentThreshold = 0 Or currentSample = "" Or currentThreshold = "" Then
                Debug.Print "Skipping for threshold at (" & currentRow & "," & thresholdColumn & ") and sample at (" & currentRow & "," & sampleColumn & ") due to malformed data"
                GoTo NextIteration
            Else
                Debug.Print "Working with sample value [" & currentSample & "]"
            End If
            
            Dim lessThanSignInSample As Boolean
            Dim sampleValue As Double
            currentSample = Replace(currentSample, ".", Application.International(xlDecimalSeparator))
            currentSample = Replace(currentSample, ",", Application.International(xlDecimalSeparator))
            lessThanSignInSample = InStr(currentSample, "<") <> 0
            If lessThanSignInSample Then
                tokens = Split(currentSample, "<")
                sampleValue = CDbl(tokens(1) * 0.999999999)
            Else
                sampleValue = CDbl(currentSample)
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

' TODO add borders to all cells in result box

If resultRowDelta > 0 Then
' If we get here, the results are set to show up below, and we should add a gray line and the names of the tests

' TODO gray row first...

' ...and now the parameter names
Dim parameterIndex As Integer
For parameterIndex = 1 To parameterNames.Rows.Count
    parameterNames.Cells(parameterIndex, 1).Copy
    parameterNames.Cells(parameterNames.Rows.Count + resultOffset + parameterIndex, 1).Value = parameterNames.Cells(parameterIndex, 1)
    parameterNames.Cells(parameterNames.Rows.Count + resultOffset + parameterIndex, 1).PasteSpecial Paste:=xlPasteFormats
Next parameterIndex

' ...copy down the thresholds as well, with formatting
Dim thresholdIndex As Integer
For thresholdIndex = 1 To thresholds.Rows.Count
    thresholds.Cells(thresholdIndex, 1).Copy
    thresholds.Cells(thresholds.Rows.Count + resultOffset + thresholdIndex, 1).Value = thresholds.Cells(thresholdIndex, 1)
    thresholds.Cells(thresholds.Rows.Count + resultOffset + thresholdIndex, 1).PasteSpecial Paste:=xlPasteFormats
Next thresholdIndex

End If

End Sub
