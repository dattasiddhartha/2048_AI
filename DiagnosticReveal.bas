Attribute VB_Name = "DiagnosticReveal"
Option Explicit

' This module will make user select to show all diagnostic / backend sheets, or hide them

Dim trigger As Integer

Sub RevealAllSheets()
    
    Dim ws As Worksheet
    
    ' trigger acts as boolean; if triggered, it will hide all sheets and show all sheets alternately
    
    If trigger = 0 Then ' setting all sheets as hidden
        For Each ws In Application.ActiveWorkbook.Worksheets
            If ws.Name <> "Advanced" Then ' setting exclusion sheets (sheets not to be hidden)
                If ws.Name <> "Basic" Then ' setting exclusion sheets (sheets not to be hidden)
                    ws.Visible = xlSheetHidden ' hiding all visible sheets
                End If
            End If
        Next
        trigger = 1 ' setting trigger as opposite to current mode
        Exit Sub
    End If
    
    If trigger = 1 Then ' converse of trigger=0
        For Each ws In ActiveWorkbook.Worksheets
            ws.Visible = xlSheetVisible ' unhide all hidden sheets
        Next ws
        trigger = 0
        Sheets("Advanced").Activate
        Exit Sub
    End If


End Sub
