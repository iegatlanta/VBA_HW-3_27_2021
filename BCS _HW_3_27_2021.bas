Attribute VB_Name = "Module8"
Sub CreateStartButton()

     
    ' CreateStartButton Macro
    ' CreateStartButton
    '

    'Set Column widths
    Columns("H:H").Select
    Selection.ColumnWidth = 17
    
    Columns("J:J").Select
    Selection.ColumnWidth = 10
    
    Columns("K:K").Select
    Selection.ColumnWidth = 10
    
    Columns("L:M").Select
    Selection.ColumnWidth = 17
        
        
    'Install Start Button
    Range("H1").Select
    ActiveSheet.Buttons.Add(480, 15, 70, 100).Select
    Selection.OnAction = "StockMarketSummaryF"
    Selection.Characters.Text = "To create summary click here..."
    With Selection.Characters(Start:=1, Length:=31).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 15
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 10
    End With
    Range("H9").Select

     'CreateDeleteButton Macro
     'Create  Delete Button

    ActiveSheet.Buttons.Add(480, 126, 70, 74.25).Select
    Selection.OnAction = "DeleteSummary2"
    Selection.Characters.Text = "Delete Summary"
    With Selection.Characters(Start:=1, Length:=21).Font
        .Name = "Calibri"
        .FontStyle = "Regular"
        .Size = 15
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 3
    End With
    Range("M1").Select
 
End Sub
Sub StockMarketSummaryF()
      
    'Establish Dims
    Dim firsti As Long, lasti As Long, rowlist As Long, I As Long, k As Long
    
    'Summary headers
    Cells(1, 10).Value = "Ticker"
    Cells(1, 11).Value = "Annual price change"
    Cells(1, 12).Value = "Precentage change"
    Cells(1, 13).Value = "Annual trading volume"
    
   Dim Total As Double
   Dim change As Double
   Dim Percentchange As Double
   Dim Start As Long
   Dim j As Long
   Dim x As Long

   Start = 2
   change = 0
   Total = 0
   x = 2
   
    'Identify Last row
    Lastrow = Cells(Rows.Count, 1).End(xlUp).Row
    'MsgBox (Lastrow)
    For j = 2 To Lastrow
        
            If (Cells(j + 1, 1).Value <> Cells(j, 1).Value) Then
                Total = Total + Cells(j, 7).Value
                If Cells(Start, 3) = 0 Then
                    For nonzero = Start To j
                        If Cells(nonzero, 3).Value <> zero Then
                            Start = nonzero
                            Exit For
                        End If
                    Next nonzero
                End If
                change = Cells(j, 6) - Cells(Start, 3)
                Percentchange = Round((change / Cells(Start, 3)) * 100, 2)
                
                
                'Start = j + 1
                'lasti = j - 1
                    
                    
                   'Summary values
                   Range("j" & x).Value = Cells(j, 1).Value
                   Range("k" & x).Value = change
                   Range("L" & x).Value = Percentchange
                   Range("M" & x).Value = Total
                   
                   
                  ' Fill "Annual Price Change",with Green and Red colors
                If Range("k" & x).Value > 0 Then
                    'Fill column with GREEN color - good
                    Range("k" & x).Interior.ColorIndex = 4
                ElseIf Range("k" & x).Value <= 0 Then
                    'Fill column with RED color - bad
                    Range("k" & x).Interior.ColorIndex = 3
                End If
                   
                   
                   'Value reset
                   Total = 0
                   change = 0
                   x = x + 1
                
    Else
    
    Total = Total + Cells(j, 7).Value
  
    End If
    Next j




            
  'Summary complete message
  MsgBox ("Summary complete")
    
End Sub
Sub DeleteSummary2()
    
    ' Delete Summary Button
    ' Macro2 Macro
    Range("J1:M500").Select
    Selection.ClearContents
    
    ' Clearcolor Macro
    ' Clear Color
    Columns("K:K").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

    'Reset borders
    Columns("K:K").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

        'Cursor reset
         Range("I1").Select
MsgBox ("Deletion complete")
    
    
End Sub

