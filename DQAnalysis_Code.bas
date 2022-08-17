Attribute VB_Name = "DQAnalysis"
'============================================================
'Define a new subroutine
'============================================================

Sub DQAnalysis()

'____________________________________________________________
'WRITE HEADER
'____________________________________________________________

'Output everything to the DQ Analysis worksheet
Worksheets("DQ Analysis").Activate
    
'Create a subtitle
Range("A1").Value = "DAQ0 (Ticker: DQ)"
            
'Create a header row with Range:
Range("A3").Value = "Year"
Range("B3").Value = "Total Daily Volume"
Range("C3").Value = "Return"

'____________________________________________________________
'DEFINE VARIABLES
'____________________________________________________________
        
     rowStart = 2
       rowEnd = 3013
     RowCount = Cells(Rows.Count, "A").End(xlUp).Row
  totalVolume = 0

Dim startingPrice As Double
Dim endingPrice As Double
      
'Work only in sheet "2018"
 Worksheets("2018").Activate

'____________________________________________________________
'FOR LOOP
'____________________________________________________________

'Begin to iterate through the whole column
For i = rowStart To rowEnd

'____________________________________________________________
'IF CONDITIONAL 1 (WITHIN FOR LOOP)
'____________________________________________________________
            
    'Conditional to be applied only if ticker is "DQ"
    If Cells(i, 1).Value = "DQ" Then
        'Increase totalVolume
        totalVolume = totalVolume + Cells(i, 8).Value
    End If
    
'____________________________________________________________
'IF CONDITIONAL 2 (WITHIN FOR LOOP)
'____________________________________________________________
    If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
        startingPrice = Cells(i, 6).Value
    End If
    
    If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
        endingPrice = Cells(i, 6).Value
    End If
    
'Increase i by one
Next i

'____________________________________________________________
'WRITE OUTPUT
'____________________________________________________________

'Display sum of Volumes
'MsgBox (totalVolume)
            
'Output results to DQ Analysis work sheet
Worksheets("DQ Analysis").Activate
    
'Write down output values
Cells(4, 1).Value = 2018
Cells(4, 2).Value = totalVolume
Cells(4, 3).Value = (endingPrice / startingPrice) - 1

'============================================================
End Sub
'============================================================
