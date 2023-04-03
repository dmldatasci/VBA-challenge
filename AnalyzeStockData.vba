Sub AnalyzeStocks()

    Dim read_row As Long
    Dim write_row As Long
    Dim open_price As Double
    Dim close_price As Double
    Dim total_volume As Double
    Dim ticker_symbol As String
    Dim worksheet_counter As Integer
    Dim min_index As Integer
    Dim max_index As Integer
    Dim tv_max_index As Integer
    
    ' Loop over all worksheets in the workbook
    For worksheet_counter = 1 To ActiveWorkbook.Worksheets.Count
    
        ' Instantiate the write row at the top of the workbook
        write_row = 1
        ' Instantiate the read row just under the header of the workbook
        read_row = 2
        
        ' Instantiate the ticker symbol for the first stock in the list
        ticker_symbol = ActiveWorkbook.Worksheets(worksheet_counter).Range("A" & read_row).Value
        ' Instantiate the open price for the first stock in the list
        open_price = ActiveWorkbook.Worksheets(worksheet_counter).Range("C" & read_row).Value
        ' Instantiate the total volume tracker
        total_volume = 0
        ' Instantiate the percent change minimum value index tracker
        min_index = 2
        ' Instantiate the percent change maximum value index tracker
        max_index = 2
        ' Instantiate the total volume maximum value index tracker
        tv_max_index = 2
        
        ' Loop through all populated rows in the worksheet
        Do While Not ActiveWorkbook.Worksheets(worksheet_counter).Range("A" & read_row).Value = ""
        
            ' Write the header
            If write_row = 1 Then
            
                ActiveWorkbook.Worksheets(worksheet_counter).Range("I" & write_row).Value = "Ticker"
                ActiveWorkbook.Worksheets(worksheet_counter).Range("J" & write_row).Value = "Yearly Change"
                ActiveWorkbook.Worksheets(worksheet_counter).Range("K" & write_row).Value = "Percent Change"
                ActiveWorkbook.Worksheets(worksheet_counter).Range("L" & write_row).Value = "Total Stock Volume"
                
                write_row = write_row + 1
                
            End If
            
            
            ' Test whether the current row is still referencing the ticker symbol we are currently considering
            If Not ActiveWorkbook.Worksheets(worksheet_counter).Range("A" & read_row).Value = ticker_symbol Then
            
                ' Set the close price of the current ticker symbol
                close_price = ActiveWorkbook.Worksheets(worksheet_counter).Range("F" & (read_row - 1)).Value
                
                ' Write the appropriate values to the current write row
                ActiveWorkbook.Worksheets(worksheet_counter).Range("I" & write_row).Value = ticker_symbol
                ActiveWorkbook.Worksheets(worksheet_counter).Range("J" & write_row).Value = close_price - open_price
                
                ' Format the percentage change cell based on increase or decrease
                If close_price - open_price < 0 Then
                    ActiveWorkbook.Worksheets(worksheet_counter).Range("J" & write_row).Interior.Color = RGB(255, 0, 0)
                ElseIf close_price - open_price > 0 Then
                    ActiveWorkbook.Worksheets(worksheet_counter).Range("J" & write_row).Interior.Color = RGB(0, 255, 0)
                End If
                
                ActiveWorkbook.Worksheets(worksheet_counter).Range("K" & write_row).NumberFormat = "0.00%"
                ActiveWorkbook.Worksheets(worksheet_counter).Range("K" & write_row).Value = (close_price - open_price) / open_price
                ActiveWorkbook.Worksheets(worksheet_counter).Range("L" & write_row).Value = total_volume
                
                ' Assess whether this stock represents a min/max percent change or a maximum total volume
                If ((close_price - open_price) / open_price) > ActiveWorkbook.Worksheets(worksheet_counter).Range("K" & max_index).Value Then
                    max_index = write_row
                End If
                If ((close_price - open_price) / open_price) < ActiveWorkbook.Worksheets(worksheet_counter).Range("K" & min_index).Value Then
                    min_index = write_row
                End If
                If total_volume > ActiveWorkbook.Worksheets(worksheet_counter).Range("L" & tv_max_index).Value Then
                    tv_max_index = write_row
                End If
                
                ' Set the new ticker symbol
                ticker_symbol = ActiveWorkbook.Worksheets(worksheet_counter).Range("A" & read_row).Value
                ' Set the opening price for the new stock of interest
                open_price = ActiveWorkbook.Worksheets(worksheet_counter).Range("C" & read_row).Value
                ' Reset the total volume tracker
                total_volume = 0
                
                ' Increase the write row counter
                write_row = write_row + 1
            
            End If
            
            ' Add the current row volume to the volume tracker
            total_volume = total_volume + ActiveWorkbook.Worksheets(worksheet_counter).Range("G" & read_row).Value
            
            ' Increase the read row counter
            read_row = read_row + 1
    
        Loop
        
        ' Write header for extreme values
        ActiveWorkbook.Worksheets(worksheet_counter).Range("P1").Value = "Ticker"
        ActiveWorkbook.Worksheets(worksheet_counter).Range("Q1").Value = "Value"
        
        ' Extract and write important values
        ActiveWorkbook.Worksheets(worksheet_counter).Range("O2").Value = "Greatest % Increase"
        ActiveWorkbook.Worksheets(worksheet_counter).Range("Q2").NumberFormat = "0.00%"
        ActiveWorkbook.Worksheets(worksheet_counter).Range("Q2").Value = ActiveWorkbook.Worksheets(worksheet_counter).Range("K" & max_index).Value
        ActiveWorkbook.Worksheets(worksheet_counter).Range("P2").Value = ActiveWorkbook.Worksheets(worksheet_counter).Range("I" & max_index).Value
        
        ActiveWorkbook.Worksheets(worksheet_counter).Range("O3").Value = "Greatest % Decrease"
        ActiveWorkbook.Worksheets(worksheet_counter).Range("Q3").NumberFormat = "0.00%"
        ActiveWorkbook.Worksheets(worksheet_counter).Range("Q3").Value = ActiveWorkbook.Worksheets(worksheet_counter).Range("K" & min_index).Value
        ActiveWorkbook.Worksheets(worksheet_counter).Range("P3").Value = ActiveWorkbook.Worksheets(worksheet_counter).Range("I" & min_index).Value
        
        ActiveWorkbook.Worksheets(worksheet_counter).Range("O4").Value = "Greatest Total Volume"
        ActiveWorkbook.Worksheets(worksheet_counter).Range("Q4").Value = ActiveWorkbook.Worksheets(worksheet_counter).Range("L" & tv_max_index).Value
        ActiveWorkbook.Worksheets(worksheet_counter).Range("P4").Value = ActiveWorkbook.Worksheets(worksheet_counter).Range("I" & tv_max_index).Value
    
    Next worksheet_counter
    
    
End Sub