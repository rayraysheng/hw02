Attribute VB_Name = "Hw02Hard02"
Sub Hw02Hard02():

Dim ws As Worksheet

For Each ws In Worksheets

    ' first set up the summary area
    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"
    
    ' set up variables for which row in the summary we'll be writing in, starting at row 2
    Dim summary_row As Long
    summary_row = 2
        
    ' keep track of a running total and the yearly open and closing
    Dim running_total As Double
    Dim opening As Double
    Dim closing As Double
    
    ' go down the rows looking at the ticker value
    ' if the ticker is not the same as the next ticker, then that section is over, so record the summary
    
    ' iterate through the data rows
        Dim last_row As Long
        last_row = ws.Cells(Rows.Count, 1).End(xlUp).Row
        
        For i = 2 To last_row
            
            ' first off is the section start
            If (ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value) Then
                    
                ' record the section ticker name in the summary
                ws.Cells(summary_row, 9).Value = ws.Cells(i, 1).Value
                
                ' reset running total
                running_total = ws.Cells(i, 7).Value
                
                ' take the opening value
                opening = ws.Cells(i, 3).Value
                
            End If
                    
            ' this is for the rows that are not the section start
            If (ws.Cells(i, 1).Value = ws.Cells(i - 1, 1).Value) Then
                
                ' add volume to running total
                running_total = running_total + ws.Cells(i, 7).Value
            End If
                
            ' if the current row is also a section end
            If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
                    
                'take the closing value
                closing = ws.Cells(i, 6).Value
                
                ' calculate the yearly change and record it
                ws.Cells(summary_row, 10).Value = closing - opening
                ' format it for color
                If (ws.Cells(summary_row, 10).Value > 0) Then
                    ws.Cells(summary_row, 10).Interior.ColorIndex = 4
                ElseIf (ws.Cells(summary_row, 10).Value < 0) Then
                    ws.Cells(summary_row, 10).Interior.ColorIndex = 3
                End If
                
                ' calculate the percent change and record it
                If (opening <> 0) Then
                    ws.Cells(summary_row, 11).Value = (closing - opening) / opening
                Else
                    ws.Cells(summary_row, 11).Value = 1
                End If
                ' format it to percent
                ws.Cells(summary_row, 11).Style = "Percent"
                
                ' record the total
                ws.Cells(summary_row, 12).Value = running_total
                    
                ' set the new summary row
                summary_row = summary_row + 1
            End If
        
        Next i
        
    ' set up the gpi gpd gtv summary area
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"
    
    ' define variables for greatest % increase, % decrease, and total volume
    Dim greatest_percent_increase As Double
    Dim gpi_ticker As String
    
    Dim greatest_percent_decrease As Double
    Dim gpd_ticker As String
    
    Dim greatest_total_volume As Double
    Dim gtv_ticker As String
    
    ' go down the summary rows
    Dim last_summary_row As Long
    last_summary_row = ws.Cells(Rows.Count, 9).End(xlUp).Row
    
    ' set checked values to zero
    greatest_percent_increase = 0
    greatest_percent_decrease = 0
    greatest_total_volume = 0
        
    For j = 2 To last_summary_row
        
        ' check the row to see its percent increase is higher than the previous gpi
        If (ws.Cells(j, 11).Value > greatest_percent_increase) Then
            
            ' nice!
            greatest_percent_increase = ws.Cells(j, 11).Value
            gpi_ticker = ws.Cells(j, 9).Value
        End If
        
        ' check the row to see its percent decrease is higher than the previous gpd
        If (ws.Cells(j, 11).Value < greatest_percent_decrease) Then
            
            ' wow you really tanked
            greatest_percent_decrease = ws.Cells(j, 11).Value
            gpd_ticker = ws.Cells(j, 9).Value
        End If
        
        ' check the row to see its total volume is higher than the previous gtv
        If (ws.Cells(j, 12).Value > greatest_total_volume) Then
            
            ' you're the man now dawg!
            greatest_total_volume = ws.Cells(j, 12).Value
            gtv_ticker = ws.Cells(j, 9).Value
        End If
    
    Next j
    
    ' now record it all
    ws.Cells(2, 16).Value = gpi_ticker
    ws.Cells(3, 16).Value = gpd_ticker
    ws.Cells(4, 16).Value = gtv_ticker
    
    ws.Cells(2, 17).Value = greatest_percent_increase
    ws.Cells(2, 17).Style = "Percent"
    ws.Cells(3, 17).Value = greatest_percent_decrease
    ws.Cells(3, 17).Style = "Percent"
    ws.Cells(4, 17).Value = greatest_total_volume
        
Next ws

End Sub



