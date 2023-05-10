Attribute VB_Name = "Module1"
Sub Stock_Market()

Dim i, summary_row As Integer
Dim last_row As Long
Dim gtv As Double
Dim t_vol As Double


Dim ticker As String
Dim git, gdt As Double
Dim yr_change, open_p, close_p, pctg_change As Double


For Each ws In Worksheets
    ws.Activate
    worksheetName = ActiveSheet.Name
open_p = 0
yr_change = 0
summary_row = 1
pctg_change = 0
t_vol = 0

last_row = Cells(Rows.Count, 1).End(xlUp).Row
    

ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"

    For i = 2 To last_row
        
        ticker = ws.Cells(i, 1).Value
        
        If open_p = 0 Then
            open_p = ws.Cells(i, 3).Value
        End If
        
        t_vol = t_vol + ws.Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        summary_row = summary_row + 1
        ws.Range("I" & summary_row).Value = ticker
        
        close_p = ws.Cells(i, 6).Value
        
        yr_change = close_p - open_p
        
        ws.Range("J" & summary_row).Value = yr_change
        
            If yr_change > 0 Then
                ws.Range("J" & summary_row).Interior.ColorIndex = 4
            Else
                ws.Range("J" & summary_row).Interior.ColorIndex = 3
            End If
        
        
        If open_p = 0 Then
            pctg_change = 0
        Else
        
            pctg_change = (yr_change / open_p)
        
        End If
        
        ws.Range("K" & summary_row).Value = Format(pctg_change, "Percent")
        
        
        open_p = 0
        
        ws.Range("L" & summary_row).Value = t_vol
        
        t_vol = 0
        
        
        End If
    Next i
     
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
     
    last_row = ws.Cells(Rows.Count, "I").End(xlUp).Row
     
    git = ws.Application.WorksheetFunction.Max(Range("K:K"))
    gdt = ws.Application.WorksheetFunction.Min(Range("K:K"))
    gtv = ws.Application.WorksheetFunction.Max(Range("L:L"))
 
    For i = 2 To last_row
        If ws.Cells(i, 12).Value = gtv Then
            ws.Range("P4").Value = ws.Cells(i, 9).Value
            ws.Range("Q4").Value = ws.Cells(i, 12).Value
        
        ElseIf ws.Cells(i, 11).Value = git Then
            ws.Range("P2").Value = ws.Cells(i, 9).Value
            ws.Range("Q2").Value = ws.Cells(i, 11).Value
        
        
        ElseIf ws.Cells(i, 11).Value = gdt Then
            ws.Range("P3").Value = ws.Cells(i, 9).Value
            ws.Range("Q3").Value = ws.Cells(i, 11).Value
        End If
        

    Next i
    
    Range("Q2").Value = Format(git, "Percent")
    Range("Q3").Value = Format(gdt, "Percent")
    
     
    Worksheets(worksheetName).Cells.EntireColumn.AutoFit
    Worksheets(worksheetName).Cells.EntireRow.AutoFit

Next ws

End Sub





