Attribute VB_Name = "Module1"
Sub Outputs()
   'Creates i and Sheet_Num variables for loop
    Dim i As Long
    Dim Sheet_Num As Long
    
    'Source for loop code: https://excelchamps.com/vba/loop-sheets/
    'Loops through each sheet and calls each sub
    Sheet_Num = Sheets.Count

    For i = 1 To Sheet_Num
        Worksheets(i).Activate
        Call Module1.Headers
        Call Module1.Ticker
        Call Module1.Yearly_And_Percent_Change
        Call Module1.Total_Vol
        Call Module1.Greatest
        Call Module1.Conditional_Formatting
        
    Next i
End Sub

Sub Headers()
    'Creates Headers for worksheet
    'Source for autofit code: https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit
    Dim Headers As Variant
    Headers = Array("Ticker", "Yearly Change", "Percent Change", "Total Stock Volume")
    Range("I1:L1").Value = Headers
    Range("I1:L1").Columns.AutoFit
    
End Sub

Sub Ticker()
    'Row variable for iterating through rows in Col A
    Dim Row As Long
    'Cout variable to iterate through rows to print out Ticker in a new cell
    Dim Count As Long
    
    'Defining variables
    ' Source: https://www.wallstreetmojo.com/vba-usedrange/
    Total_Row = ActiveSheet.UsedRange.Rows.Count
    Count = 2
    
    'For loop for Ticker in Column A
        For Row = 2 To Total_Row
            'Checks if the rows are different, Prints the value in Col I, Increases count and row variables by one
            If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
                Cells(Count, 9).Value = Cells(Row, 1).Value
                Count = Count + 1
                Row = Row + 1
               
            End If
        Next Row
End Sub
    
Sub Yearly_And_Percent_Change()
    'Row variable for iterating through rows in Col
    Dim Row As Long
    'Cout variable to iterate through rows to print out yearly change in new row
    Dim Count As Long
    'Variable to store first time a ticker was seen for calculations
    Dim First_Sight As Long
    'Variable for yearly change calculation
    Dim Yearly_Change As Double
    'Variable for percent change calculation
    Dim Percent_Change As Double
    
    
    'Defining variables
    ' Source: https://www.wallstreetmojo.com/vba-usedrange/
    Total_Row = ActiveSheet.UsedRange.Rows.Count
    Count = 2
    First_Sight = 2
    
    'For loop for Ticker in Column A
        For Row = 2 To Total_Row
            'Checks if the rows are different, Prints the value in Col I, Increases count and row variables by one
            If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
                'Yearly change calculation
                Yearly_Change = (Cells(Row, 6).Value) - (Cells(First_Sight, 3).Value)
                Cells(Count, 10).Value = Yearly_Change
                
                'Percent change calculation
                Percent_Change = (Yearly_Change) / (Cells(First_Sight, 3).Value)
                Cells(Count, 11).Value = Percent_Change
                
                Count = Count + 1
                First_Sight = Row + 1
                
            End If
        Next Row
    Range("K:K").NumberFormat = "0.00%"

End Sub

Sub Total_Vol()
    'Row variable for iterating through rows in Col
    Dim Row As Long
    'Cout variable to iterate through rows to print out yearly change in new row
    Dim Count As Long
    'Variable for total stock Volume
    Dim Total As LongLong

    'Defining variables
    ' Source: https://www.wallstreetmojo.com/vba-usedrange/
    Total_Row = ActiveSheet.UsedRange.Rows.Count
    Count = 2
    Total = 0
    
    'For loop for Ticker in Column A
        For Row = 2 To Total_Row
            'Checks if the rows are different, Prints the value in Col I, Increases count and row variables by one
            If Cells(Row + 1, 1).Value <> Cells(Row, 1).Value Then
                Total = Total + Cells(Row, 7).Value
                Cells(Count, 12).Value = Total
    
                Count = Count + 1
                Total = 0
            'Else continue to tally total
            Else
                Total = Total + Cells(Row, 7).Value
                 
            End If
        Next Row

End Sub

Sub Greatest()
    'Creates header variables and defines arrays
    Dim Row_Headers As Variant
    Dim Col_Headers As Variant
    Row_Headers = Array("Greatest % Increase", "Greatest % Decrease", "Greatest Total Volume")
    Col_Headers = Array("Ticker", "Value")
    
    'Fills range with array values
    'Source: https://stackoverflow.com/questions/13982511/dumping-array-vertical-direction-in-excel-vba
    Range("N2:N4").Value = Application.Transpose(Row_Headers)
    Range("O1:P1").Value = Col_Headers
    
    'Creates variables for values and loop variable
    Dim Row As Long
    Dim Greatest_Inc As Double
    Dim Greatest_Dec As Double
    Dim Greatest_Vol As LongLong
    
    'Cells(2, 11).Value = Greatest_Inc
    'Cells(2, 11).Value = Greatest_Dec
    'Cells(2, 12).Value = Greatest_Inc
    
    ' Source: https://www.wallstreetmojo.com/vba-row-count/
    Total_Row = Range("I1").End(xlDown).Row
    Cells(2, 16).NumberFormat = "0.00%"
    Cells(3, 16).NumberFormat = "0.00%"
     
    'For loop to loop through rows of data
    For Row = 2 To Total_Row
        'Checks for greatest increase
        If Cells(Row, 11).Value > Cells(Row + 1, 11).Value And Cells(Row, 11).Value > Greatest_Inc Then
            Greatest_Inc = Cells(Row, 11).Value
            Cells(2, 15).Value = Cells(Row, 9).Value
            Cells(2, 16).Value = Greatest_Inc
        End If
        'Checks for greatest decrease
        If Cells(Row, 11).Value < Cells(Row + 1, 11).Value And Cells(Row, 11).Value < Greatest_Dec Then
            Greatest_Dec = Cells(Row, 11).Value
            Cells(3, 15).Value = Cells(Row, 9).Value
            Cells(3, 16).Value = Greatest_Dec
        End If
        'Checks for greatest total Volume
        If Cells(Row, 12).Value > Cells(Row + 1, 12).Value And Cells(Row, 12).Value > Greatest_Vol Then
            Greatest_Vol = Cells(Row, 12).Value
            Cells(4, 15).Value = Cells(Row, 9).Value
            Cells(4, 16).Value = Greatest_Vol
        End If
    Next Row
    
    
    'Autofit
    'Source for autofit code: https://learn.microsoft.com/en-us/office/vba/api/excel.range.autofit
    Range("N1:P4").Columns.AutoFit
    
End Sub

Sub Conditional_Formatting()
Dim Row As Integer

' Source: https://www.wallstreetmojo.com/vba-row-count/
Total_Row = Range("J1").End(xlDown).Row

'Conditional formatting for yearly change
For Row = 2 To Total_Row
    'Positive rows green
    If Cells(Row, 10).Value > 0 Then
        Cells(Row, 10).Interior.ColorIndex = 4
    'Negative rows red
    ElseIf Cells(Row, 10).Value < 0 Then
        Cells(Row, 10).Interior.ColorIndex = 3
    End If
Next Row

'Conditional formatting for percent change
For Row = 2 To Total_Row
    'Positive rows green
    If Cells(Row, 11).Value > 0 Then
        Cells(Row, 11).Interior.ColorIndex = 4
    'Negative rows red
    ElseIf Cells(Row, 11).Value < 0 Then
        Cells(Row, 11).Interior.ColorIndex = 3
    End If
Next Row

End Sub

