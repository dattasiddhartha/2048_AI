Attribute VB_Name = "MonteCarlo_Module"
Option Explicit
Dim cell_index1 As Integer, cell_index2 As Integer, cell_index3 As Integer
Dim cell_index4 As Integer, cell_index5 As Integer, cell_index6 As Integer
Dim cell_index7 As Integer, cell_index8 As Integer, cell_index9 As Integer
Dim Score As Long
Dim gameOverScore As Long
Dim randomCell As Long
Dim countCell As Long
Dim twoFour As Long
Dim moveCell As Long


Dim counterDirection As Long
Dim rownumber As Long
Dim oldRownumber As Long

' Monte Carlo script
'' This script wil perform random walk Monte Carlo simulations (i.e. perform random moves till the iteration count is reached)
'' It will execute a branching tree outwards (breadth first search -> depth first search)
'' It will first look through up down left right combinations along each row/iteration
'' If a new Max (2/4/8/16/32/64/...) is located, then that breadth level is skipped and the current branch will be expanded

Function RunSimulation()
    
    Dim startingBoard As Collection ' store initial board values in array
    Set startingBoard = New Collection
    Dim iterator As Long ' iterate through values in board
    
    Dim i As Long ' iterator for DFS-BFS section
    Dim existingMaxValue As Long ' a placeholder for the maximum board value, to be set and beaten, so that depth can be executed instead of breadth
    
    
    '''''''' first iteration (loading initial state)
    If counterDirection = 0 Then ' ensuring that this is the first move
        
        ' Load initial Board
        startingBoard.Add Sheets("Advanced").Range("index1")
        startingBoard.Add Sheets("Advanced").Range("index2")
        startingBoard.Add Sheets("Advanced").Range("index3")
        startingBoard.Add Sheets("Advanced").Range("index4")
        startingBoard.Add Sheets("Advanced").Range("index5")
        startingBoard.Add Sheets("Advanced").Range("index6")
        startingBoard.Add Sheets("Advanced").Range("index7")
        startingBoard.Add Sheets("Advanced").Range("index8")
        startingBoard.Add Sheets("Advanced").Range("index9")
        
        iterator = 0
        Do Until iterator = startingBoard.Count ' storing all values in board (from indices 1 to 9) onto AISim data sheet
            Sheets("AISim").Range("A2").Offset(0, iterator) = startingBoard(iterator + 1)
            iterator = iterator + 1
        Loop
        
        counterDirection = 1 ' signify loading of the first initial state
        
    End If
    ''''''''''' completed first iteration ^
    
    rownumber = 0
    oldRownumber = 0
    
    ' begin tree search (establishing first set of breadth branches)
    PressUp
    rownumber = rownumber + 1
    LoadBoard (rownumber)
    PressLeft
    rownumber = rownumber + 1
    LoadBoard (rownumber)
    PressRight
    rownumber = rownumber + 1
    LoadBoard (rownumber)
    PressDown
    rownumber = rownumber + 1
    LoadBoard (rownumber)
    
    oldRownumber = oldRownumber + 1
    
    '''''''''''''''''''''''''''''''''''''
    ' begin depth search upon each branch -> breadth search upon the branches -> repeat

    For i = 0 To Sheets("variableStorage").Range("B5").Value ' users store the number of runs they would like to execute (in AI runs field)
    ' Note: moves to be executed tend to be double of this number
        
        RestoreBoard (oldRownumber)
        existingMaxValue = maxValue(rownumber) ' setting max value to beat
        
        PressUp
        rownumber = rownumber + 1
        LoadBoard (rownumber)
         'if max value of this row > previous max value
          ' then set new max
          ' restore current row, do this action up again, and continue
          ' set oldRownumber as current rownumber
        '' if existing max value is beat, then we explore this branch instead of going through the other breadth branches
        '' (i.e. we shift from BFS to DFS when we get closer to finding path of highest max values)
        If maxValue(rownumber) > existingMaxValue Then
            existingMaxValue = maxValue(rownumber)
            oldRownumber = rownumber ' set the lagging row counter to the current maxvalue-containing row
            RestoreBoard (oldRownumber)
            PressUp
            rownumber = rownumber + 1
            LoadBoard (rownumber)
        End If
        
        PressLeft
        rownumber = rownumber + 1
        LoadBoard (rownumber)
        If maxValue(rownumber) > existingMaxValue Then
            existingMaxValue = maxValue(rownumber)
            oldRownumber = rownumber
            RestoreBoard (oldRownumber)
            PressLeft
            rownumber = rownumber + 1
            LoadBoard (rownumber)
        End If
        
        PressRight
        rownumber = rownumber + 1
        LoadBoard (rownumber)
        If maxValue(rownumber) > existingMaxValue Then
            existingMaxValue = maxValue(rownumber)
            oldRownumber = rownumber
            RestoreBoard (oldRownumber)
            PressRight
            rownumber = rownumber + 1
            LoadBoard (rownumber)
        End If
        
        PressDown
        rownumber = rownumber + 1
        LoadBoard (rownumber)
        If maxValue(rownumber) > existingMaxValue Then
            existingMaxValue = maxValue(rownumber)
            oldRownumber = rownumber
            RestoreBoard (oldRownumber)
            PressDown
            rownumber = rownumber + 1
            LoadBoard (rownumber)
        End If
    
        oldRownumber = oldRownumber + 1
        
        i = i + 1
        
        ' Reset game if lose; objective is to win the game regardless of how many games/moves it takes
        '' Note: The guided simulation (pruning AI) is more conservative with respect to moves
        If Sheets("AISim").Range("A2").Offset(rownumber, 9) = "Lose" Then Advanced_Execution.InitializeGame
        
        If Sheets("AISim").Range("A2").Offset(rownumber, 9) = "Win" Then
            MsgBox "AI wins with " & Sheets("Advanced").Range("moves_count").Value & " moves"
            Exit For
        End If
    
        ' The workbook tends to crash or stutter with no display when allowed to run without lag
        '' Setting save function at every 100 iterations allows it to (1) save current values, and
        '' (2) slow down the Monte Carlo process and reduce likelihood of crashing
        If i Mod 100 = 0 Then ActiveWorkbook.Save ' in case workbook crashes during simulation
        
        ' eliminate score overflow
        Sheets("Advanced").Range("score").Value = 0
        ' no score computation in below functions due to high risk of overflow error in monte carlo simulations
    
    Next i
    
End Function

Sub Main()
    
    ' This macro integrates the simulation function, and first seeks input from user regarding iteration count / stopping limit
    
    Dim response As Variant
    Dim iterationInput As Long
    
    ' warning message to users: Monte Carlo can be extremely time intensive if users are not careful
    response = MsgBox("If the number of moves you set is too high, then your workbook will take an extremely long time to finish processing. Do you understand?", vbYesNo + vbCritical, "Monte Carlo Warning")
    
    If response = vbYes Then
        iterationInput = Application.InputBox("How many moves would you like?", "Monte Carlo", Type:=1)
        Sheets("Advanced").Range("I12").Value = iterationInput ' setting iteration count
    Else: MsgBox "You clicked No"
    End If
    
    RunSimulation

End Sub

Function maxValue(rownumber) As Long

    ' this function is used to determine the largest max value derived so far in the board layout of index rownumber (current)

    Dim max_to_beat As Long ' this initializes a max to beat, within the default function
    Dim iterator As Long

    max_to_beat = 0
    
    iterator = 0
    Do Until iterator = 9 ' iteratively changing the max_to_beat based on the latest max while going throught the row
        If Sheets("AISim").Range("A2").Offset(rownumber, iterator) > max_to_beat Then
            max_to_beat = Sheets("AISim").Range("A2").Offset(rownumber, iterator)
        End If
        iterator = iterator + 1
    Loop
    
    maxValue = max_to_beat

End Function

' below functions have explanations in other modules

Function LoadBoard(rownumber)

    Dim iterator As Long
    Dim startingBoard As Collection
    Set startingBoard = New Collection

    startingBoard.Add Sheets("Advanced").Range("index1")
    startingBoard.Add Sheets("Advanced").Range("index2")
    startingBoard.Add Sheets("Advanced").Range("index3")
    startingBoard.Add Sheets("Advanced").Range("index4")
    startingBoard.Add Sheets("Advanced").Range("index5")
    startingBoard.Add Sheets("Advanced").Range("index6")
    startingBoard.Add Sheets("Advanced").Range("index7")
    startingBoard.Add Sheets("Advanced").Range("index8")
    startingBoard.Add Sheets("Advanced").Range("index9")
        
    iterator = 0
    Do Until iterator = startingBoard.Count
        Sheets("AISim").Range("A2").Offset(rownumber, iterator) = startingBoard(iterator + 1)
        iterator = iterator + 1
    Loop

End Function

Function RestoreBoard(oldRownumber)

    Sheets("Advanced").Range("index1") = Sheets("AISim").Range("A2").Offset(oldRownumber, 1 - 1)
    Sheets("Advanced").Range("index2") = Sheets("AISim").Range("A2").Offset(oldRownumber, 2 - 1)
    Sheets("Advanced").Range("index3") = Sheets("AISim").Range("A2").Offset(oldRownumber, 3 - 1)
    Sheets("Advanced").Range("index4") = Sheets("AISim").Range("A2").Offset(oldRownumber, 4 - 1)
    Sheets("Advanced").Range("index5") = Sheets("AISim").Range("A2").Offset(oldRownumber, 5 - 1)
    Sheets("Advanced").Range("index6") = Sheets("AISim").Range("A2").Offset(oldRownumber, 6 - 1)
    Sheets("Advanced").Range("index7") = Sheets("AISim").Range("A2").Offset(oldRownumber, 7 - 1)
    Sheets("Advanced").Range("index8") = Sheets("AISim").Range("A2").Offset(oldRownumber, 8 - 1)
    Sheets("Advanced").Range("index9") = Sheets("AISim").Range("A2").Offset(oldRownumber, 9 - 1)

End Function

Function CellShifting(cellA, cellB, cellC, scoreCell)
    ' cellA
    ' cellB
    ' cellC
    If Range(cellB).Value <> "" And Range(cellA).Value = "" Then
        Range(cellA).Value = Range(cellB).Value
        Range(cellB).Value = ""
        moveCell = True
    End If
    
    If Range(cellC).Value <> "" And Range(cellB).Value = "" Then
        If Range(cellA).Value = "" Then
            Range(cellA).Value = Range(cellC).Value
            Range(cellC).Value = ""
            moveCell = True
        Else
            Range(cellB).Value = Range(cellC).Value
            Range(cellC).Value = ""
            moveCell = True
        End If
    End If
    
    If Range(cellB).Value <> "" Then
        If Range(cellB).Value = Range(cellA).Value Then
            Range(cellA).Value = Range(cellA).Value * 2
            Range(cellB).Value = Range(cellC).Value
            Range(cellC).Value = ""
            moveCell = True
        End If
    End If
    
    If Range(cellC).Value <> "" Then
        If Range(cellC).Value = Range(cellB).Value Then
            If Range(cellB).Value = Range(cellA).Value Then
                Range(cellA).Value = Range(cellA).Value * 2
                Range(cellC).Value = ""
                moveCell = True
            Else
                Range(cellB).Value = Range(cellB).Value * 2
                Range(cellC).Value = ""
                moveCell = True
            End If
        End If
    End If
End Function

Function Clearing()

    ' Add random two or four
    If Range("index1").Value = "" Then countCell = countCell + 1: cell_index1 = countCell
    If Range("index2").Value = "" Then countCell = countCell + 1: cell_index2 = countCell
    If Range("index3").Value = "" Then countCell = countCell + 1: cell_index3 = countCell
    If Range("index4").Value = "" Then countCell = countCell + 1: cell_index4 = countCell
    If Range("index5").Value = "" Then countCell = countCell + 1: cell_index5 = countCell
    If Range("index6").Value = "" Then countCell = countCell + 1: cell_index6 = countCell
    If Range("index7").Value = "" Then countCell = countCell + 1: cell_index7 = countCell
    If Range("index8").Value = "" Then countCell = countCell + 1: cell_index8 = countCell
    If Range("index9").Value = "" Then countCell = countCell + 1: cell_index9 = countCell
   
    gameOverScore = 0
   
    If countCell <> 0 And moveCell = True Then
        Randomize
        randomCell = CInt((countCell - 1) * Rnd() + 1) ' 1 to countCell
        Randomize
        twoFour = CInt(1 * Rnd() + 1) * 2
        If randomCell = cell_index1 Then Range("index1").Value = twoFour
        If randomCell = cell_index2 Then Range("index2").Value = twoFour
        If randomCell = cell_index3 Then Range("index3").Value = twoFour
        If randomCell = cell_index4 Then Range("index4").Value = twoFour
        If randomCell = cell_index5 Then Range("index5").Value = twoFour
        If randomCell = cell_index6 Then Range("index6").Value = twoFour
        If randomCell = cell_index7 Then Range("index7").Value = twoFour
        If randomCell = cell_index8 Then Range("index8").Value = twoFour
        If randomCell = cell_index9 Then Range("index9").Value = twoFour
   
    ElseIf countCell = 0 Then
        ' Check Game Over - lose
        If Range("index2").Value <> Range("index1").Value And Range("index2").Value <> Range("index5").Value And Range("index2").Value <> Range("index3").Value Then
            gameOverScore = gameOverScore + 1
        End If
        If Range("index4").Value <> Range("index1").Value And Range("index4").Value <> Range("index5").Value And Range("index4").Value <> Range("index7").Value Then
            gameOverScore = gameOverScore + 1
        End If
        If Range("index6").Value <> Range("index3").Value And Range("index6").Value <> Range("index5").Value And Range("index6").Value <> Range("index9").Value Then
            gameOverScore = gameOverScore + 1
        End If
        If Range("index8").Value <> Range("index7").Value And Range("index8").Value <> Range("index5").Value And Range("index8").Value <> Range("index9").Value Then
            gameOverScore = gameOverScore + 1
        End If
    End If
    
    ' Check Game Over - win
    If Range("index1").Value = [difficulty] Then gameOverScore = 5
    If Range("index2").Value = [difficulty] Then gameOverScore = 5
    If Range("index3").Value = [difficulty] Then gameOverScore = 5
    If Range("index4").Value = [difficulty] Then gameOverScore = 5
    If Range("index5").Value = [difficulty] Then gameOverScore = 5
    If Range("index6").Value = [difficulty] Then gameOverScore = 5
    If Range("index7").Value = [difficulty] Then gameOverScore = 5
    If Range("index8").Value = [difficulty] Then gameOverScore = 5
    If Range("index9").Value = [difficulty] Then gameOverScore = 5
    
    If gameOverScore = 4 Then Sheets("AISim").Range("A2").Offset(rownumber, 9) = "Lose" ' MsgBox "Game Over!"
    If gameOverScore = 5 Then Sheets("AISim").Range("A2").Offset(rownumber, 9) = "Win" ' MsgBox "You win!"
    If gameOverScore < 4 Then gameOverScore = 0

End Function


Function PressUp()

    Dim output As Variant

    moveCell = False
    countCell = 0
    cell_index1 = 0: cell_index2 = 0: cell_index3 = 0
    cell_index4 = 0: cell_index5 = 0: cell_index6 = 0
    cell_index7 = 0: cell_index8 = 0: cell_index9 = 0
    
    output = CellShifting("index1", "index4", "index7", "score")
    output = CellShifting("index2", "index5", "index8", "score")
    output = CellShifting("index3", "index6", "index9", "score")
    
    Clearing
    
End Function

Function PressDown()

    Dim output As Variant

    moveCell = False
    countCell = 0
    cell_index1 = 0: cell_index2 = 0: cell_index3 = 0
    cell_index4 = 0: cell_index5 = 0: cell_index6 = 0
    cell_index7 = 0: cell_index8 = 0: cell_index9 = 0
    
    ' compare A to B, B to C, C to A
    output = CellShifting("index7", "index4", "index1", "score")
    output = CellShifting("index8", "index5", "index2", "score")
    output = CellShifting("index9", "index6", "index3", "score")
    
    Clearing
    
End Function

Function PressLeft()

    Dim output As Variant

    moveCell = False
    countCell = 0
    cell_index1 = 0: cell_index2 = 0: cell_index3 = 0
    cell_index4 = 0: cell_index5 = 0: cell_index6 = 0
    cell_index7 = 0: cell_index8 = 0: cell_index9 = 0
    
    ' compare A to B, B to C, C to A
    output = CellShifting("index1", "index2", "index3", "score")
    output = CellShifting("index4", "index5", "index6", "score")
    output = CellShifting("index7", "index8", "index9", "score")
    
    Clearing
    
End Function

Function PressRight()

    Dim output As Variant

    moveCell = False
    countCell = 0
    cell_index1 = 0: cell_index2 = 0: cell_index3 = 0
    cell_index4 = 0: cell_index5 = 0: cell_index6 = 0
    cell_index7 = 0: cell_index8 = 0: cell_index9 = 0
    
    ' compare A to B, B to C, C to A
    output = CellShifting("index3", "index2", "index1", "score")
    output = CellShifting("index6", "index5", "index4", "score")
    output = CellShifting("index9", "index8", "index7", "score")
    
    Clearing
    
End Function

