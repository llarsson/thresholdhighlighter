Attribute VB_Name = "ThresholdHighlighter"
Sub ThresholdHighlighter2000()
Dim samples As Range, thresholds As Range

Set testNames = Application.InputBox("Välj namnen på ämnena", "Threshold highlighter 2000", Type:=8)
Set thresholds = Application.InputBox("Välj gränsvärdena", "Threshold highlighter 2000", Type:=8)
Set samples = Application.InputBox("Välj proverna", "Threshold highlighter 2000", Type:=8)

If thresholds.Rows.Count <> samples.Rows.Count Then
    MsgBox ("Olika antal prover och gränsvärden har valts!")
    Exit Sub
End If

' Where do we want our results to wind up? Below or to the side?
Dim resultRowDelta As Integer, resultColumnDelta As Integer
'If MsgBox("Välj Ja för att se svaren på sidan och Nej för att se svaren nedanför.", vbYesNo, "Resultaten på sidan eller ej?") = vbYes Then
'    resultRowDelta = 0
'    resultColumnDelta = samples.Columns.Count ' Place them this many columns to the side'
'Else
'    resultRowDelta = samples.Rows.Count + 1 ' Place them this many rows below (add one row for extra space)
'    resultColumnDelta = 0
'End If

' Nope, always just put the results in a box below
resultRowDelta = samples.Rows.Count + 1
resultColumnDelta = 0

For currentRow = 1 To samples.Rows.Count
    For sampleColumn = 1 To samples.Columns.Count
        For thresholdColumn = 1 To thresholds.Columns.Count
            currentSample = samples.Cells(currentRow, sampleColumn)
            currentThreshold = thresholds.Cells(currentRow, thresholdColumn)
            
            If IsEmpty(currentThreshold) Or IsEmpty(currentSample) Or IsNull(currentSample) Or IsNull(currentThreshold) Or currentThreshold = 0 Then
                Debug.Print "Skipping for threshold at (" & currentRow & "," & thresholdColumn & ") and sample at (" & currentRow & "," & sampleColumn & ") due to malformed data"
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
                    
                    samples.Cells(currentRow + resultRowDelta, sampleColumn + resultColumnDelta).Value = "Rapporteringsgr√§ns > RV"
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

' ...and now the test names
Dim nameIndex As Integer
For nameIndex = 1 To testNames.Rows.Count
    testNames.Cells(nameIndex, 1).Copy
    testNames.Cells(testNames.Rows.Count + 1 + nameIndex, 1).Value = testNames.Cells(nameIndex, 1)
    testNames.Cells(testNames.Rows.Count + 1 + nameIndex, 1).PasteSpecial Paste:=xlPasteFormats
Next nameIndex

End If

End Sub
