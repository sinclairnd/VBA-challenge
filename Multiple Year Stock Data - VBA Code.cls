VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub stockMacro()

Dim analysisWs, dataWs As Worksheet

Dim yearInput, ticker, lastTicker As String

Dim lastRow, i, counter As Long

Dim yearOpen, yearClose, totalVolume As Double

'Input Box
yearInput = InputBox("What year would you like to analyze?")

Set analysisWs = Sheets("Stocks Analysis")
analysisWs.Cells.ClearContents


Set dataWs = Sheets(yearInput)

lastRow = dataWs.Cells(dataWs.Rows.Count, "A").End(xlUp).Row

'Formatting Analysis ws
analysisWs.Cells(1, 1).Value = "Stocks Analysis for Year " + yearInput
analysisWs.Cells(3, 1).Value = "Ticker"
analysisWs.Cells(3, 2).Value = "Yearly Change"
analysisWs.Cells(3, 3).Value = "Percent Change"
analysisWs.Cells(3, 4).Value = "Total Stock Volume"

lastTicker = ""
totalVolume = 0
counter = 4


'For loop
For i = 2 To lastRow
    ticker = dataWs.Cells(i, 1).Value
    If ticker <> lastTicker Or i = lastRow Then
        If i <> 2 Then
            If i <> lastRow Then
                yearClose = dataWs.Cells(i - 1, 6)
            Else
                yearClose = dataWs.Cells(i, 6)
            End If
            
            analysisWs.Cells(counter, 1).Value = lastTicker
            analysisWs.Cells(counter, 2).Value = yearClose - yearOpen
            
            If yearClose - yearOpen > 0 Then
                analysisWs.Cells(counter, 2).Interior.ColorIndex = 4
            ElseIf yearClose - yearOpen < 0 Then
                analysisWs.Cells(counter, 2).Interior.ColorIndex = 3
            Else
                analysisWs.Cells(counter, 2).Interior.ColorIndex = 15
            End If
            
            If yearOpen <> 0 Then
                analysisWs.Cells(counter, 3).Value = (yearClose - yearOpen) / yearOpen
            Else
                analysisWs.Cells(counter, 3).Value = "N/A"
            End If
            
            analysisWs.Cells(counter, 4).Value = totalVolume
            counter = counter + 1
            totalVolume = 0
        End If
        
        yearOpen = dataWs.Cells(i, 3).Value
    End If
    
    totalVolume = totalVolume + dataWs.Cells(i, 7).Value
        
    lastTicker = ticker
    
Next i

Range("c:c").NumberFormat = "0.00%"


End Sub
