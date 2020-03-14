
 Sub Multi_year_stock_data()
 
 ' Set Ws_Current as a worksheet object variable.
    Dim Ws_Current As Worksheet
    Dim Summary_Table As Boolean
    Dim COMMAND_SPREADSHEET As Boolean
    
    Summary_Table = False       'Set Header flag
    COMMAND_SPREADSHEET = True              'Hard part flag
    
    ' Loop through all of the worksheets in the active workbook.
    For Each Ws_Current In Worksheets
    
        ' Set initial variable for holding the ticker name
        Dim Ticker As String
        Ticker = " "
        
        ' Set an initial variable for holding the total per ticker name
        Dim Total_Ticker As Double
        Total_Ticker = 0
        
        ' Set new variables 
        Dim openprice As Double
        openprice = 0
        Dim closeprice As Double
        closeprice = 0
        Dim PriceChange As Double
        PriceChange= 0
        Dim Price_Percent As Double
        Price_Percent = 0
        ' Set new variables for Hard Solution Part
        Dim max_ticker As String
        max_ticker = " "
        Dim min_ticker As String
        min_ticker = " "
        Dim max_percent As Double
        max_percent = 0
        Dim min_percent As Double
        min_percent = 0
        Dim max_total_ticker As String
        max_total_ticker = " "
        Dim max_volume As Double
        max_volume = 0
     
        ' Keep track of the location for each ticker name
        ' in the summary table for the current worksheet
        Dim Summary_Table_Row As Long
        Summary_Table_Row = 2
        
        ' Set initial row count for the current worksheet
        Dim Lastrow As Long
        Dim i As Long
        
        Lastrow = Ws_Current.Cells(Rows.Count, 1).End(xlUp).Row

        
        If Summary_Table Then
            ' Set Titles for the Summary Table for current worksheet
            Ws_Current.Range("I1").Value = "Ticker"
            Ws_Current.Range("J1").Value = "Yearly Change"
            Ws_Current.Range("K1").Value = "Percent Change"
            Ws_Current.Range("L1").Value = "Total Stock Volume"
            ' Set Additional Titles for new Summary Table on the right for current worksheet
            Ws_Current.Range("O2").Value = "Greatest % Increase"
            Ws_Current.Range("O3").Value = "Greatest % Decrease"
            Ws_Current.Range("O4").Value = "Greatest Total Volume"
            Ws_Current.Range("P1").Value = "Ticker"
            Ws_Current.Range("Q1").Value = "Value"
        Else
            'This is the first, resulting worksheet, reset flag for the rest of worksheets
            Summary_Table = True
        End If
        
        ' Set initial value of Open Price for the first Ticker of Ws_Current,
        ' The rest ticker's open price 
        openprice = Ws_Current.Cells(2, 3).Value
        
        ' Loop from the beginning of the current worksheet(Row2) till last row
        For i = 2 To Lastrow
        
      
            ' Check if we are still within the same ticker name,
            ' if not - write results to summary table
            If Ws_Current.Cells(i + 1, 1).Value <> Ws_Current.Cells(i, 1).Value Then
            
                ' Set the ticker name, we are ready to insert this ticker name data
                Ticker = Ws_Current.Cells(i, 1).Value
                
                ' Calculate PriceChanged 
                closeprice = Ws_Current.Cells(i, 6).Value
                PriceChange= closeprice - openprice
                ' Check Division by 0 condition
                If openprice <> 0 Then
                    Price_Percent = (PriceChange/ openprice) * 100
                Else
                    ' Unlikely, but it needs to be checked to avoid program crushing
                    'MsgBox ("For " & Ticker & ", Row " & CStr(i) & ": Open Price =" & 'openprice & ". Fix <open> field manually and save the spreadsheet.")
                End If
                
                ' Add to the Ticker name total volume
                Total_Ticker = Total_Ticker + Ws_Current.Cells(i, 7).Value
              
                
                ' Print the Ticker Name in the Summary Table, Column I
                Ws_Current.Range("I" & Summary_Table_Row).Value = Ticker
                ' Print the Ticker Name in the Summary Table, Column I
                Ws_Current.Range("J" & Summary_Table_Row).Value = Delta_Price
                ' Fill "Yearly Change", i.e. PriceChangewith Green and Red colors
                If (PriceChange> 0) Then
                    'Fill column with GREEN color - lets goo
                    Ws_Current.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
                ElseIf (PriceChange<= 0) Then
                    'Fill column with RED color - danger
                    Ws_Current.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
                End If
                
                 ' Print the Ticker Name in the Summary Table, Column I
                Ws_Current.Range("K" & Summary_Table_Row).Value = (CStr(Price_Percent) & "%")
                ' Print the Ticker Name in the Summary Table, Column J
                Ws_Current.Range("L" & Summary_Table_Row).Value = Total_Ticker
                
                ' Add 1 to the summary table row count
                Summary_Table_Row = Summary_Table_Row + 1
                ' Resetting place holders , as we will be working with new Ticker
                PriceChange= 0
               
                closeprice = 0
                ' Capture next Ticker's openprice
                openprice = Ws_Current.Cells(i + 1, 3).Value
              
                
               
                ' Keep track of all counters and do calculations within the current spreadsheet
                If (Price_Percent > max_percent) Then
                    max_percent = Price_Percent
                    max_ticker = Ticker
                ElseIf (Price_Percent < min_percent) Then
                    min_percent = Price_Percent
                    min_ticker = Ticker
                End If
                       
                If (Total_Ticker > max_volume) Then
                    max_volume = Total_Ticker
                    max_total_ticker = Ticker
                End If
                
                ' resetting counters
                Price_Percent = 0
                Total_Ticker = 0
                
            
            'Else - If the cell immediately following a row is still the same ticker name,
            'just add to Totl Ticker Volume
            Else
                ' Encrease the Total Ticker Volume
                Total_Ticker = Total_Ticker + Ws_Current.Cells(i, 7).Value
            End If
            
        Next i

            
            ' Check if it is not the first spreadsheet
            ' Record all new counts to the new summary table on the right of the current spreadsheet
            If Not COMMAND_SPREADSHEET Then
            
                Ws_Current.Range("Q2").Value = (CStr(max_percent) & "%")
                Ws_Current.Range("Q3").Value = (CStr(min_percent) & "%")
                Ws_Current.Range("P2").Value = max_ticker
                Ws_Current.Range("P3").Value = min_ticker
                Ws_Current.Range("Q4").Value = max_volume
                Ws_Current.Range("P4").Value = max_total_ticker
                
            Else
                COMMAND_SPREADSHEET = False
            End If
        
     Next Ws_Current
End Sub
