Attribute VB_Name = "DQAnalysis"
'============================================================
' Define a new subroutine
Sub DQAnalysis()
    
    ' Output everything to the DQ Analysis worksheet
    Worksheets("DQ Analysis").Activate
    
        ' Create a subtitle
        Range("A1").Value = "DAQ0 (Ticker: DQ)"
        
        ' Create a header row
        Cells(3, 1).Value = "Year"
        Cells(3, 2).Value = "Total Daily Volume"
        Cells(3, 3).Value = "Return"
       
        ' We can also create a header row with Range():
        'Range("A3").Value = "Year"
        'Range("B3").Value = "Total Daily Volume"
        'Range("C3").Value = "Return"
'____________________________________________________________
        
       rowStart = 2
         rowEnd = 3013
    totalVolume = 0
        
    ' Work only in sheet "2018"
    Worksheets("2018").Activate

        ' Begin to iterate through the whole column
        For i = rowStart To rowEnd
            
            ' Conditional to be applied only if ticker is "DQ"
            If Cells(i, 1).Value = "DQ" Then
                ' Increase totalVolume
                totalVolume = totalVolume + Cells(i, 8).Value
           End If
           
        ' Increase i by one
        Next i

            ' Display sum of Volumes
            MsgBox (totalVolume)
            
    ' Output results to DQ Analysis work sheet
    Worksheets("DQ Analysis").Activate
    
        ' Write down output values
        Cells(4, 1).Value = 2018
        Cells(4, 2).Value = totalVolume
'____________________________________________________________
End Sub
'============================================================
