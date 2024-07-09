Attribute VB_Name = "Module1"
Sub analyze_stocks_data()

    ' --------------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' --------------------------------------------
    For Each ws In Worksheets
    
        ' --------------------------------------------
        ' INSERT THE YEAR
        ' --------------------------------------------
    
        ' Create a Variable to Hold File Name, Last Row, and Last Column
        Dim WorksheetName As String
        Dim lastColumn As Long
        
        ' Determine the Last Column
        lastColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column
        
        ' Determine the Last Row
        lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        ' Grabbed the WorksheetName
        WorksheetName = ws.Name
        'MsgBox WorksheetName
        
        ' Add a Column for the Ticker, Quaterly Change, Percentage Change,Total Stock Volume, Ticker and Value
        ws.Cells(1, lastColumn + 2).Value = "Ticker"
        ws.Cells(1, lastColumn + 3).Value = "Quaterly Change"
        ws.Cells(1, lastColumn + 4).Value = "Percentage Change"
        ws.Cells(1, lastColumn + 5).Value = "Total Stock Volume"
        ws.Cells(1, lastColumn + 9).Value = "Ticker"
        ws.Cells(1, lastColumn + 10).Value = "Value"
       'Variables to process the read data
        Dim tickerValue As String
        Dim openPrice As Variant
        Dim closePrice As Variant
        Dim quaterlyChange As Variant
        Dim volume As Variant
        Dim volumeTotal As Variant
        Dim percentage As Variant
        Dim calRow As Integer
        Dim readOpenPrice As Integer
        'Variables to populate the max and min percentage and max of volume
        Dim maxPercentage As Variant
        Dim minPercentage As Variant
        Dim maxVolume As Variant
        Dim maxPercCell As Range
        Dim minPercCell As Range
        Dim maxVolumeCell As Range
        Dim rowIndex As Integer
        'Row count for the column that contains the calculated % and volume per Ticker
        Dim summaryRowCount As Long
        rowIndex = 2
        calRow = 2
        'Navigate through the table
        For i = 2 To lastRow
        
           If readOpenPrice = 0 Then
            openPrice = ws.Cells(i, 3).Value
            readOpenPrice = 1
           End If
                  
           tickerValue = ws.Cells(i, 1).Value
           volume = ws.Cells(i, 7).Value
           volumeTotal = volumeTotal + volume
           
           'Next Ticker Value is different from current value
           If tickerValue <> ws.Cells(i + 1, 1).Value Then
           'Calculate the values
            closePrice = ws.Cells(i, 6).Value
            quaterlyChange = closePrice - openPrice
            percentage = (quaterlyChange / openPrice)
           'Write the data to the respective cells
            ws.Cells(calRow, lastColumn + 2).Value = tickerValue
            ws.Cells(calRow, lastColumn + 3).Value = quaterlyChange
            'Color change
            If quaterlyChange >= 0 Then
             ws.Cells(calRow, lastColumn + 3).Interior.Color = vbGreen
            Else
             ws.Cells(calRow, lastColumn + 3).Interior.Color = vbRed
            End If
            ws.Cells(calRow, lastColumn + 4).Value = Format(percentage, "0.00%")
            ws.Cells(calRow, lastColumn + 5).Value = volumeTotal
            'Reset all the values
            volumeTotal = 0
            quaterlyChange = 0
            percentage = 0
            closePrice = 0
            openPrice = 0
            tickerValue = ""
            calRow = calRow + 1
            readOpenPrice = 0
           End If
            
        Next i
            
        summaryRowCount = ws.Cells(ws.Rows.Count, lastColumn + 4).End(xlUp).Row
        
        'Populate Greatest % increase
        ws.Cells(rowIndex, lastColumn + 8).Value = "Greatest % increase"
        ' Find the greatest % increase
        maxPercentage = Format(Application.WorksheetFunction.Max(ws.Range("K2:K" & summaryRowCount)), "0.00%")
        ' Find the corresponding Ticker Value
        Set maxPercCell = ws.Range("K2:K" & summaryRowCount).Find(maxPercentage, LookAt:=xlWhole)
        ' Populate the corresponding Ticker and Max Percentage Value
        ws.Cells(rowIndex, lastColumn + 9).Value = maxPercCell.Offset(0, -2).Value
        ws.Cells(rowIndex, lastColumn + 10).Value = maxPercentage
        
        rowIndex = rowIndex + 1
        
        'Populate Greatest % decrease
        ws.Cells(rowIndex, lastColumn + 8).Value = "Greatest % decrease"
        ' Find the greatest % increase
        minPercentage = Format(Application.WorksheetFunction.Min(ws.Range("K2:K" & summaryRowCount)), "0.00%")
        ' Find the corresponding Ticker Value
        Set minPercCell = ws.Range("K2:K" & summaryRowCount).Find(minPercentage, LookAt:=xlWhole)
        ' Populate the corresponding Ticker Value
        ws.Cells(rowIndex, lastColumn + 9).Value = minPercCell.Offset(0, -2).Value
        ' Min Percentage
        ws.Cells(rowIndex, lastColumn + 10).Value = minPercentage
        rowIndex = rowIndex + 1
        
        'Populate Greatest total volume
        ws.Cells(rowIndex, lastColumn + 8).Value = "Greatest total volume"
        ' Find the greatest % increase cell info
        maxVolume = Application.WorksheetFunction.Max(ws.Range("L2:L" & summaryRowCount))
        ' Find the corresponding Ticker Value
        Set maxVolumeCell = ws.Range("L2:L" & summaryRowCount).Find(maxVolume, LookAt:=xlWhole)
        ' Add Ticker and Volume Result
        ws.Cells(rowIndex, lastColumn + 9).Value = maxVolumeCell.Offset(0, -3).Value
        ws.Cells(rowIndex, lastColumn + 10).Value = maxVolumeCell
                
        Next ws

End Sub
