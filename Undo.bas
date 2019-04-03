Attribute VB_Name = "Undo"
Option Explicit

' This function will store moves in a document and restore previous moves if user clicks Undo


Function LoadDataFromUserValueList()

    ' load second last row in cell, then restore it
    Dim rowCount As Integer
    Dim iterator As Integer
    Dim emptyrowCounter As Integer
    Dim indexcell As Range
    
    ' finding first completely empty row
    emptyrowCounter = 0 ' counts number of rows until a completely empty one (empty defined as all 9 cell indices are empty)
    rowCount = 0
    Do Until emptyrowCounter = 9 ' only check until index 9, not whole row
        emptyrowCounter = 0
        For iterator = 0 To 8 ' checking through index 1 to 9 for board values
            If IsEmpty(Sheets("UserValuePositionList").Range("A2").Offset(rowCount, iterator)) Then ' finding first empty row from the top row down
                emptyrowCounter = emptyrowCounter + 1
                ' Debug.Print Str(rowCount) & Str(iterator) & Str(emptyrowCounter)
            End If
        Next iterator
        rowCount = rowCount + 1
    Loop
    
    ' restore row: desired row should be 2 before the empty row
    Range("index1") = Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 3, 0)
    Range("index2") = Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 3, 1)
    Range("index3") = Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 3, 2)
    Range("index4") = Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 3, 3)
    Range("index5") = Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 3, 4)
    Range("index6") = Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 3, 5)
    Range("index7") = Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 3, 6)
    Range("index8") = Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 3, 7)
    Range("index9") = Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 3, 8)
    ' set prior score
    Range("score") = Sheets("UserValuePositionList").Range("A1").Offset(rowCount - 3, 9)
    
    
    ' clear out row that should never have happened
    If rowCount > 2 Then
        Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 2, 0) = ""
        Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 2, 1) = ""
        Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 2, 2) = ""
        Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 2, 3) = ""
        Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 2, 4) = ""
        Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 2, 5) = ""
        Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 2, 6) = ""
        Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 2, 7) = ""
        Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 2, 8) = ""
        Sheets("UserValuePositionList").Range("A2").Offset(rowCount - 2, 9) = ""
    End If
    
    ' adjust rowUVL value
    Sheets("variableStorage").Range("B3") = Sheets("variableStorage").Range("B3") - 1
    
End Function

Sub Undo()
    
    ' only do undo if not in first move
    Dim filledCounter As Integer
    
    If Not IsEmpty(Range("index1")) Then filledCounter = filledCounter + 1
    If Not IsEmpty(Range("index2")) Then filledCounter = filledCounter + 1
    If Not IsEmpty(Range("index3")) Then filledCounter = filledCounter + 1
    If Not IsEmpty(Range("index4")) Then filledCounter = filledCounter + 1
    If Not IsEmpty(Range("index5")) Then filledCounter = filledCounter + 1
    If Not IsEmpty(Range("index6")) Then filledCounter = filledCounter + 1
    If Not IsEmpty(Range("index7")) Then filledCounter = filledCounter + 1
    If Not IsEmpty(Range("index8")) Then filledCounter = filledCounter + 1
    If Not IsEmpty(Range("index9")) Then filledCounter = filledCounter + 1
    
    If filledCounter > 2 Then LoadDataFromUserValueList

End Sub
