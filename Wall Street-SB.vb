Sub datarun()


'dimension variables etc
Dim lastRow As Double
Dim Lastcol As Double
Dim ticker As String
Dim datenum As Double
Dim openfig As Double
Dim highfig As Double
Dim lowfig As Double
Dim closefig As Double
Dim volfig As Double
Dim totalstock As Double
Dim oldstock As Double
Dim opentotalfig As Double
Dim closetotalfig As Double
Dim tickerflag As Boolean
Dim writerow As Integer
Dim lastIRow As Long ' last row in column I
Dim firstbrow As Long ' last blank row in column i
Dim lastIcolrow As Integer


'Set Up Columns

    Range("I1").Select
    ActiveCell.FormulaR1C1 = "Ticker"
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Yearly Change"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Percent Change"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Total Stock Volume"
    Range("R1").Select
    ActiveCell.FormulaR1C1 = "Open Value"
    Range("S1").Select
    ActiveCell.FormulaR1C1 = "Last Value"
    Range("T1").Select
    ActiveCell.FormulaR1C1 = "Open Vol"
    Range("U1").Select
    ActiveCell.FormulaR1C1 = "Close Vol"
    Range("V1").Select
    ActiveCell.FormulaR1C1 = "Total Stock"
    Range("V2").Select



'obtain number of rows and columns and assign to variables
lastRow = ActiveSheet.UsedRange.Rows.Count
Lastcol = ActiveSheet.UsedRange.Columns.Count

'work through the sheet data to create new columns of additional information
'based on ticker code

For i = 2 To lastRow




        
    'get the row and assign to variables
    ticker = Cells(i, 1).Value
    datenum = Cells(i, 2).Value
    openfig = Cells(i, 3).Value
    highfig = Cells(i, 4).Value
    lowfig = Cells(i, 5).Value
    closefig = Cells(i, 6).Value
    volfig = Cells(i, 7).Value
      
    'search the target ticker summary table
    'check if a code is present
    'if not cycle down to the next blank and add in
    'we have the code, opening price i.e. the first time the code is encountered
    'and also the opening volume. we also have the closing volume and code.
    
    
    
    'cycle down the ticker column to find the code if it exists
    
     
    lastIRow = Cells(Rows.Count, "I").End(xlUp).Row
    
    
    
    
    firstbrow = lastIRow + 1
    
    
    'cycle down and check if the ticker exists, if not add it at the end
        
        'check if the ticker is in the table
            For r = 2 To lastIRow
            
                     
                     If Cells(r, 9) = ticker Then
                     tickerflag = True
                     writerow = r
                     Exit For
                     
                     Else
                     tickerflag = False
                                        
                     End If
                     
            Next
            
            
            Select Case tickerflag
            
            Case Is = True
                        'if it is add the info in
                        Cells(writerow, 19) = closefig
                        Cells(writerow, 21) = volfig
                        totalstock = Cells(writerow, 22) + volfig
                        Cells(writerow, 22) = totalstock
                        
                         
                        
                        
            Case Is = False
                        'if it isnt add the info in a row at the end
                        Cells(firstbrow, 9) = ticker
                        Cells(firstbrow, 18) = openfig
                        Cells(firstbrow, 20) = volfig
                        Cells(firstbrow, 19) = closefig
                        Cells(firstbrow, 21) = volfig
                        Cells(firstbrow, 22) = volfig
                        
                        
                    
            End Select
              
    
Next i

' Calculate values and input into summary table

lastIcolrow = Cells(Rows.Count, 9).End(xlUp).Row

For i = 2 To lastIcolrow

        Cells(i, 10) = Cells(i, 19) - Cells(i, 18)
        Cells(i, 11) = Cells(i, 10) / Cells(i, 18)
        Cells(i, 12) = Cells(i, 22)
        
        
        
        
Next


Columns("K:K").Select
    Selection.NumberFormat = "0.00%"


Columns("J:J").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16752384
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13561798
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, _
        Formula1:="=0"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Range("J1").Select
    Selection.FormatConditions.Delete


Columns("J:J").Select
    Selection.NumberFormat = "0.00"
    
Columns("K:K").Select
    Selection.NumberFormat = "0.00"

End Sub
