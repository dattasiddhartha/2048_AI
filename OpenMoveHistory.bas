Attribute VB_Name = "OpenMoveHistory"
Option Explicit

Sub OpenMoveHistory()
Attribute OpenMoveHistory.VB_ProcData.VB_Invoke_Func = " \n14"

    ' This subroutine shows users all the moves they have been performing consecutively within a game
    ' It shows complete history not just for current but historical games
    '' consecutive moves, e.g. U L U D means Up Left Up Down
    
    Dim ws As Worksheet
    
    For Each ws In ActiveWorkbook.Worksheets
            ws.Visible = xlSheetVisible ' unhide all hidden sheets
    Next ws
    Sheets("UserMovesList").Activate
    
End Sub
