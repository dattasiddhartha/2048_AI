Attribute VB_Name = "AI_Module"
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

' Expectimax Alpha-beta pruning script
'' This script is better than the random walk monte carlo in the sense that it is more conservative with moves
'' It works by forecasting (by monte carlo simulations) the total scores that would be obtained from a certain move at time t
'' then when it is done, it determines the weighted score of moving in each direction (in percentage)

Sub Main()

    ' executing the AI simulation function
    RunAI

End Sub

Function RunAI()
    
    Dim dirString As String ' this string will store the cell name to be used in Range to store percentage weights
    Dim totalCount As Long
    Dim testingValue As Long
    Dim initialDirection As Long
    Dim Direction As Long
    Dim ScoreArray(1 To 4) As Long ' array of scores
    Dim DirectionalWeights(1 To 4) As Long ' array of weights assigned to each direction
    Dim ScoreSumofWeights As Long ' total score to be used in weight function
    Dim MaxWeight As Long ' largest weight which determines which direction to follow next
    Dim DefaultColor As Long
    Dim DirectionalIterationCount As Long
    Dim DirectionalIterationScore As Long
    Dim desiredNumber As Integer ' the score to congratulate user for reaching
    Dim refreshRate As Long ' rate to save and prevent crashing
    
    ' setting the difficulty number (e.g. 256) as the trigger, that when it is obtained, the simulation can end
    desiredNumber = Sheets("Advanced").Range("E3").Value
    
    refreshRate = 0
    ScoreSumofWeights = 0
    totalCount = Sheets("Advanced").Range("AIruns")
    
    ' default state
    '' reset weight values
    Range("N11").Value = 0: Range("M12").Value = 0: Range("O12").Value = 0: Range("N13").Value = 0
    
    '' reset colors
    Range("N11").Interior.Color = RGB(217, 223, 242)
    Range("M12").Interior.Color = RGB(217, 223, 242)
    Range("O12").Interior.Color = RGB(217, 223, 242)
    Range("N13").Interior.Color = RGB(217, 223, 242)
    
    '' storing initial values
    Sheets("boardDuplicate").Range("A1") = Sheets("Advanced").Range("index1")
    Sheets("boardDuplicate").Range("B1") = Sheets("Advanced").Range("index2")
    Sheets("boardDuplicate").Range("C1") = Sheets("Advanced").Range("index3")
    Sheets("boardDuplicate").Range("A2") = Sheets("Advanced").Range("index4")
    Sheets("boardDuplicate").Range("B2") = Sheets("Advanced").Range("index5")
    Sheets("boardDuplicate").Range("C2") = Sheets("Advanced").Range("index6")
    Sheets("boardDuplicate").Range("A3") = Sheets("Advanced").Range("index7")
    Sheets("boardDuplicate").Range("B3") = Sheets("Advanced").Range("index8")
    Sheets("boardDuplicate").Range("C3") = Sheets("Advanced").Range("index9")
    
    
    ' for each direction to be traversed from the initial board state
    For initialDirection = 1 To 4
        DirectionalWeights(initialDirection) = 0 ' initial value of weight assigned as 0
        ScoreArray(initialDirection) = 0
        For testingValue = 1 To totalCount
            DirectionalIterationCount = 0
            
            ' loading board layout to test directional movements upon
            Sheets("Advanced").Range("index1") = Sheets("boardDuplicate").Range("A1")
            Sheets("Advanced").Range("index2") = Sheets("boardDuplicate").Range("B1")
            Sheets("Advanced").Range("index3") = Sheets("boardDuplicate").Range("C1")
            Sheets("Advanced").Range("index4") = Sheets("boardDuplicate").Range("A2")
            Sheets("Advanced").Range("index5") = Sheets("boardDuplicate").Range("B2")
            Sheets("Advanced").Range("index6") = Sheets("boardDuplicate").Range("C2")
            Sheets("Advanced").Range("index7") = Sheets("boardDuplicate").Range("A3")
            Sheets("Advanced").Range("index8") = Sheets("boardDuplicate").Range("B3")
            Sheets("Advanced").Range("index9") = Sheets("boardDuplicate").Range("C3")
            
            Direction = initialDirection
            Do ' run a combination of Up Down Left Right
                DirectionalIterationCount = DirectionalIterationCount + 1
                Select Case Direction
                    Case 1: PressUp
                    Case 2: PressLeft
                    Case 3: PressRight
                    Case 4: PressDown
                End Select
                
                DirectionalIterationScore = Range("score") ' used as the objective/utility value to determine performance of a certain direction
                
                If DirectionalIterationCount = 1 And DirectionalIterationScore = -1 Then
                    Exit Do ' in case the counting or scoring mechanism failed in generating a proper move, e.g. gameover scenario
                End If
                
                ScoreArray(initialDirection) = ScoreArray(initialDirection) + DirectionalIterationScore ' storing actual score, to be used later to set weight for a direction
                
                ' if desired game value (e.g. 256) is reached, simulation complete
                If [index1] = desiredNumber Or [index2] = desiredNumber Or [index3] = desiredNumber Or [index4] = desiredNumber Or [index5] = desiredNumber Or [index6] = desiredNumber Or [index7] = desiredNumber Or [index8] = desiredNumber Or [index9] = desiredNumber Then
                    MsgBox "Hit!"
                    Exit Function
                End If
                
                ' if all cells filled, game over - at least this branch within the simulation will be closed off and the next will open
                If [index1] <> "" And [index2] <> "" And [index3] <> "" And [index4] <> "" And [index5] <> "" And [index6] <> "" And [index7] <> "" And [index8] <> "" And [index9] <> "" Then
                    Exit Do
                End If
                
                Direction = Int((4 - 1 + 1) * Rnd + 1)
                
                ' ensuring no crashes
                refreshRate = refreshRate + 1
                If refreshRate Mod 1000 = 0 Then
                    ActiveWorkbook.Save
                End If
                
            Loop
        Next
        ScoreSumofWeights = ScoreSumofWeights + ScoreArray(initialDirection) ' aggregating total weight
    Next
    
    If ScoreSumofWeights > 0 Then
        MaxWeight = 0
        For initialDirection = 1 To 4
            DirectionalWeights(initialDirection) = CLng(100 * CDbl(ScoreArray(initialDirection)) / CDbl(ScoreSumofWeights)) ' Weight(i) = score_i / total_scores * 100%
            If DirectionalWeights(initialDirection) > MaxWeight Then ' identify heaviest weighted direction to be used / recommended to user in Hint_viaAI
                MaxWeight = DirectionalWeights(initialDirection)
            End If
            Select Case initialDirection
                Case 1: dirString = "N11"
                Case 2: dirString = "M12"
                Case 3: dirString = "O12"
                Case 4: dirString = "N13"
            End Select
            Range(dirString) = DirectionalWeights(initialDirection) ' setting weight values
        Next

        For initialDirection = 1 To 4
            Select Case initialDirection
                Case 1: dirString = "N11"
                Case 2: dirString = "M12"
                Case 3: dirString = "O12"
                Case 4: dirString = "N13"
            End Select
            If DirectionalWeights(initialDirection) = MaxWeight Then ' setting max weight's color as indicator
                DefaultColor = RGB(255, 79, 79)
                Range(dirString).Interior.Color = DefaultColor
            End If
        Next
    Else 'ScoreSumofWeights = 0
        For initialDirection = 1 To 4
            Select Case initialDirection
                Case 1: dirString = "N11"
                Case 2: dirString = "M12"
                Case 3: dirString = "O12"
                Case 4: dirString = "N13"
            End Select
            DefaultColor = RGB(217, 223, 242) ' restoring default colors
            Range(dirString) = ""
            Range(dirString).Interior.Color = DefaultColor
        Next
    End If
    
End Function

' comments and logic explained in MonteCarlo module

Function maxValue(rownumber) As Long

    Dim max_to_beat As Long
    Dim iterator As Long
        
    max_to_beat = 0
    
    iterator = 0
    Do Until iterator = 9
        If Sheets("AISim").Range("A2").Offset(rownumber, iterator) > max_to_beat Then
            max_to_beat = Sheets("AISim").Range("A2").Offset(rownumber, iterator)
        End If
        iterator = iterator + 1
    Loop
    
    maxValue = max_to_beat

End Function

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
            Score = Score + Range(cellA).Value: Range(scoreCell).Value = Score
            moveCell = True
        End If
    End If
    
    If Range(cellC).Value <> "" Then
        If Range(cellC).Value = Range(cellB).Value Then
            If Range(cellB).Value = Range(cellA).Value Then
                Range(cellA).Value = Range(cellA).Value * 2
                Range(cellC).Value = ""
                Score = Score + Range(cellA).Value: Range(scoreCell).Value = Score
                moveCell = True
            Else
                Range(cellB).Value = Range(cellB).Value * 2
                Range(cellC).Value = ""
                Score = Score + Range(cellB).Value: Range(scoreCell).Value = Score
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
    
    If gameOverScore = 4 Then Sheets("AISim").Range("A2").Offset(rownumber, 9) = "Lose"
    If gameOverScore = 5 Then Sheets("AISim").Range("A2").Offset(rownumber, 9) = "Win"
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


