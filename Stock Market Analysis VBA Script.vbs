

Sub stockinfo()
Dim WS As Worksheet
    
    
For Each WS In ActiveWorkbook.Worksheets
WS.Activate

Dim ticker As String
Dim volume As Double
    volume = 0
Dim resultrow As Double
    resultrow = 2
Dim sp As Double
Dim lp As Double
Dim lastrow As Long
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    
Dim max As Double
Dim min As Double
Dim maxvol As Double
Dim maxticker As String
Dim minticker As String
Dim maxvolticker As String




Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volume"

sp = Cells(2, 3).Value

For i = 2 To lastrow
    If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
        volume = (volume) + Cells(i, 7).Value
    
    Else
        ticker = Cells(i, 1).Value
        Cells(resultrow, 9).Value = ticker
        lp = Cells(i, 6).Value
        Cells(resultrow, 10).Value = lp - sp
        If Cells(resultrow, 10).Value < 0 Then
            Cells(resultrow, 10).Interior.ColorIndex = 3
            Else
            Cells(resultrow, 10).Interior.ColorIndex = 4
            End If
        If sp = 0 Then
            Cells(resultrow, 11).Value = "N/A"
            Else
                Cells(resultrow, 11).Value = ((lp / sp) - 1)
                Cells(resultrow, 11).Style = "Percent"
                Cells(resultrow, 11).NumberFormat = "0.00%"
            End If
        volume = volume + Cells(i, 7).Value
        Cells(resultrow, 12).Value = volume
        volume = 0
        lp = 0
        sp = 0
        sp = Cells(i + 1, 3).Value
        resultrow = resultrow + 1

    End If
Next i

max = Cells(2, 11).Value
maxticker = Cells(2, 9).Value
For j = 2 To lastrow
    If max > Cells(j + 1, 11) Then
        max = max
        Cells(2, 16).Value = maxticker
        Else
        max = Cells(j + 1, 11).Value
        maxticker = Cells(j + 1, 9).Value
        Cells(2, 16).Value = maxticker
        End If
    Next j
Cells(2, 17).Value = max
Cells(2, 17).NumberFormat = "0.00%"

min = Cells(2, 11).Value
minticker = Cells(2, 9).Value
For k = 2 To lastrow
    If min < Cells(k + 1, 11) Then
        min = min
        Cells(3, 16).Value = minticker
        Else
        min = Cells(k + 1, 11).Value
        minticker = Cells(k + 1, 9).Value
        Cells(3, 16).Value = minticker
        End If
    Next k
Cells(3, 17).Value = min
Cells(3, 17).NumberFormat = "0.00%"

maxvol = Cells(2, 12).Value
maxvolticker = Cells(2, 9).Value
For m = 2 To lastrow
    If maxvol > Cells(m + 1, 11) Then
        maxvol = maxvol
        Cells(4, 16).Value = maxvolticker
        Else
        maxvol = Cells(m + 1, 11).Value
        maxvolticker = Cells(m + 1, 9).Value
        Cells(4, 16).Value = maxvolticker
        End If
    Next m
Cells(4, 17).Value = maxvol

Next WS

End Sub
