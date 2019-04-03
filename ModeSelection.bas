Attribute VB_Name = "ModeSelection"
Option Explicit

' This module contains the macros that help users activate the simpler, basic game & the more sophisticated, advanced game

Sub SelectBasic()

    Sheets("Basic").Activate
    Application.Wait (Now + TimeValue("0:00:01")) ' delay
    Advanced_Execution.InitializeGame

End Sub

Sub SelectAdvanced()

    Sheets("Advanced").Activate
    Application.Wait (Now + TimeValue("0:00:01")) ' delay
    Advanced_Execution.InitializeGame

End Sub
