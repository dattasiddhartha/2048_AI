Attribute VB_Name = "ClearAISimContents"
Option Explicit

Sub ClearAISimSheet()
Attribute ClearAISimSheet.VB_ProcData.VB_Invoke_Func = " \n14"

    ' This is a macro on the AISim worksheet to clear all cells with input

    Rows("2:1048576").Select
    Selection.ClearContents
    
End Sub
