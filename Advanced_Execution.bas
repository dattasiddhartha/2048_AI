Attribute VB_Name = "Advanced_Execution"
Option Explicit

' old variables
Dim Score As Integer ' score value to be displayed
Dim gameOverScore As Integer ' counter that determines whether game is won or lost
Dim randomCell As Integer
Dim countCell As Integer
Dim twoFour As Integer
Dim moveCell As Boolean

' cells to store values from the board
Dim cell_index1 As Integer, cell_index2 As Integer, cell_index3 As Integer
Dim cell_index4 As Integer, cell_index5 As Integer, cell_index6 As Integer
Dim cell_index7 As Integer, cell_index8 As Integer, cell_index9 As Integer

Public Const defaultcell As String = "E4" ' setting default cell as board box index 1
Dim Row, Column As Integer

Dim highScore As Integer
Dim moveCount As Integer

' ActiveX labels
Dim scoreLabel As OLEObject
Dim highscoreLabel As OLEObject

' row and columns to be stored in the User Moves List sheet
Dim rowUML As Long
Dim colUML As Long
Dim historicalEntry As Long

' user moves list default parameters
Const attemptCell As String = "A2"
Const scoreCell As String = "B2"
Const displayscoreCell As String = "C2"
Const movesCell As String = "D2"

Dim utilityValue As Long ' objective function value to be used in pruning

' Dim arrowKeyToggle As Long ' may consider using it as boolean

' row value for User Value Position List
Dim rowUVL As Long

' This Module contains the complete workings on the interface of the Advanced game
'' UserMovesList: This list keeps track of all moves taken by the user across all previous and current games
'' UserValuePositionList: This list keeps track of the board values per move made by the user in the current game
'' AISim data sheet shows the permutations / values executed by the Random Walk Monte Carlo
'' User should click on Toggle Backend to see all the backend data sheets

Sub InitializeGame()

    ' Subroutine to create new game layout by generating 2 random positions for number 2
    
    ' board indexing layout
    ' 1 | 2 | 3
    ' 4 | 5 | 6
    ' 7 | 8 | 9

    Dim randomRow_1, randomCol_1, randomRow_2, randomCol_2 As Integer ' the distinct positions to place the random 2/4s
    Dim coinflipResult As Integer ' coin flip to determine where to place the random 2/4s
    Dim two_four_barrier As Double ' the threshold that determines whether the placed value is a 2 or 4
    Dim distinctPositions As Boolean ' ensuring the 2 values generated do not lie in the same board position
    
    ' insertions in UserMovesList
    Sheets("UserMovesList").Range("C1").Offset(rowUML, 0) = Range("score").Value
    Sheets("UserMovesList").Range("C1") = "displayscore"
    
    ' color change
    ColourMod.ColourSchemeAdjustment
    
    ' ActiveX label parameters
    Set scoreLabel = Sheets("Basic").OLEObjects("Label2"): Set highscoreLabel = Sheets("Basic").OLEObjects("Label3")
      
   ' setting high score
   If Range("score").Value > Range("high_score") Then
        Range("high_score").Value = Range("score").Value
        If ActiveSheet.Name = "Basic" Then ' ActiveX label only exists on the Basic game, so a check is needed
            highscoreLabel.Object.Caption = Range("high_score").Value
        End If
    End If
   
   ' reset score
   Score = 0
   Range("score") = 0
   scoreLabel.Object.Caption = 0
   ' highscoreLabel.Object.Caption = 0 ' need do this if completely new player want to play game from scratch
   
   ' reset move count
   moveCount = 0
   Range("moves_count") = moveCount
   
   ' wipe out all cell values
   For Row = 1 To 3  '1 refers to index 1 box, 3 refers to index 7 box
        For Column = 1 To 3 '1 refers to index 1 box, 3 refers to index 3 box
            Range(defaultcell).Offset(Row - 1, Column - 1) = "" ' reshape Row and Column into index starting from 0 and Offsetting from default cell location
        Next Column
   Next Row
   
   ' Random value generation
   '' generate either 2 or 4 with coinflip
   two_four_barrier = 0.9 ' the threshold: if the coinflip generates a value below this barrier, then returns 2
   If Rnd <= two_four_barrier Then
        coinflipResult = 2
    Else
        coinflipResult = 4
    End If
    
   '' generate distinct random positions
   randomRow_1 = Int((3 - 1 + 1) * Rnd + 1) ' generating position 1
   randomCol_1 = Int((3 - 1 + 1) * Rnd + 1)
   
   distinctPositions = False
   Do While distinctPositions = False ' Using boolean counter to continue generating position 2 till both are distinct
        randomRow_2 = Int((3 - 1 + 1) * Rnd + 1)
        randomCol_2 = Int((3 - 1 + 1) * Rnd + 1)
        
        If randomRow_2 <> randomRow_1 And randomCol_2 <> randomCol_1 Then ' ensuring the row and col positions are not identical
            distinctPositions = True
        End If
        
    Loop
   
   ' setting the positions offsetted from default cell with the value result from the coinflip
   Range(defaultcell).Offset(randomRow_1 - 1, randomCol_1 - 1) = coinflipResult
   Range(defaultcell).Offset(randomRow_2 - 1, randomCol_2 - 1) = coinflipResult
   
   ' entry in UserMovesList sheet
   rowUML = rowUML + 1
   Sheets("variableStorage").Range("B2") = rowUML ' storing variables to be used across modules
   colUML = 0
   Sheets("UserMovesList").Range(attemptCell).Offset(rowUML - 1, 0) = "Attempt" & Str(historicalEntry)
   historicalEntry = historicalEntry + 1
   
   ' reset User Value List
   Sheets("UserValuePositionList").UsedRange.ClearContents ' clearing history for new set of moves
   rowUVL = 0: Sheets("variableStorage").Range("B3") = rowUVL  ' change value in case altered by Undo
   ' add board values to UserValuePositionList
    rowUVL = rowUVL + 1: Sheets("variableStorage").Range("B3") = rowUVL  ' saving value
    StoreDataInUserValueList ' store initial board values
    
    ' default state of Hint section
    '' reset weight values
    Sheets("Advanced").Range("N11").Value = 0: Sheets("Advanced").Range("M12").Value = 0: Sheets("Advanced").Range("O12").Value = 0: Sheets("Advanced").Range("N13").Value = 0
    
    '' reset colors
    Sheets("Advanced").Range("N11").Interior.Color = RGB(217, 223, 242)
    Sheets("Advanced").Range("M12").Interior.Color = RGB(217, 223, 242)
    Sheets("Advanced").Range("O12").Interior.Color = RGB(217, 223, 242)
    Sheets("Advanced").Range("N13").Interior.Color = RGB(217, 223, 242)
   
End Sub

Sub DeleteHistory()
    
    ' This allows users to delete move history from all prior games
    
    '' deletes history
    Sheets("UserMovesList").UsedRange.ClearContents
    historicalEntry = 0
    rowUML = 0
    
    '' reinstantiates the column names
    Sheets("UserMovesList").Range("A1") = "attemptCount"
    Sheets("UserMovesList").Range("B1") = "utilityScore"
    Sheets("UserMovesList").Range("C1") = "displayscore"
    Sheets("UserMovesList").Range("D1") = "movesOrder"

End Sub

Function CellShifting(cellA, cellB, cellC, scoreCell)

    ' This function facilitates shifting of cells, either vertically or horizontally, depending on cell indices inputted
    
    ' sample situation: shifting left (A:1, B:4, C:7)
    '' cellA
    '' cellB
    '' cellC
    
    ' reassigning values from the shift, that when they are not empty, the next cell would possess the previous cell's value
    If Range(cellB).Value <> "" And Range(cellA).Value = "" Then
        Range(cellA).Value = Range(cellB).Value
        Range(cellB).Value = ""
        moveCell = True
    End If
    
    ' reassigning values from the shift, that when they are not empty, the next cell would possess the previous cell's value
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
        If Range(cellB).Value = Range(cellA).Value Then ' combining cells that are of equivalent value
            Range(cellA).Value = Range(cellA).Value * 2
            Range(cellB).Value = Range(cellC).Value
            Range(cellC).Value = ""
            Score = Score + Range(cellA).Value: Range(scoreCell).Value = Score ' adding up scores
            If ActiveSheet.Name = "Basic" Then scoreLabel.Object.Caption = Score ' setting score in case it is Basic sheet
            moveCell = True
        End If
    End If
    
    If Range(cellC).Value <> "" Then
        If Range(cellC).Value = Range(cellB).Value Then
            If Range(cellB).Value = Range(cellA).Value Then
                Range(cellA).Value = Range(cellA).Value * 2
                Range(cellC).Value = ""
                Score = Score + Range(cellA).Value: Range(scoreCell).Value = Score
                If ActiveSheet.Name = "Basic" Then scoreLabel.Object.Caption = Score
                moveCell = True
            Else
                Range(cellB).Value = Range(cellB).Value * 2
                Range(cellC).Value = ""
                Score = Score + Range(cellB).Value: Range(scoreCell).Value = Score
                If ActiveSheet.Name = "Basic" Then scoreLabel.Object.Caption = Score
                moveCell = True
            End If
        End If
    End If
    
End Function

Function Clearing()

    ' The purpose of this function is "clearing" of the board situation;
    ' if the gameover counter hits 4 or 5, it determines game completion or failure

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
    If Range("index1").Value = [difficulty] Then gameOverScore = 5 ' use difficulty cell in sheet to permit dynamic difficulty setting
    If Range("index2").Value = [difficulty] Then gameOverScore = 5
    If Range("index3").Value = [difficulty] Then gameOverScore = 5
    If Range("index4").Value = [difficulty] Then gameOverScore = 5
    If Range("index5").Value = [difficulty] Then gameOverScore = 5
    If Range("index6").Value = [difficulty] Then gameOverScore = 5
    If Range("index7").Value = [difficulty] Then gameOverScore = 5
    If Range("index8").Value = [difficulty] Then gameOverScore = 5
    If Range("index9").Value = [difficulty] Then gameOverScore = 5
    
    If gameOverScore < 4 Then gameOverScore = 0 ' if gameoverscore is not lose or win, then it is set to 0 (counter is reset)
    
    If gameOverScore = 4 Then ' failure condition
        MsgBox "Game Over!"
    End If
    
    If gameOverScore = 5 Then ' success condition
        MsgBox "You win!"
    End If

End Function


Function PressUp()
    
    ' This function sets model for upward movement action and results
    '' The difference between up and other directions in terms of function code is the cell indices
    
    Dim output As Variant ' a placeholder to execute the CellShifting function along the UP cell indices
    
    ' initializing default board parameters
    moveCell = False
    countCell = 0
    cell_index1 = 0: cell_index2 = 0: cell_index3 = 0
    cell_index4 = 0: cell_index5 = 0: cell_index6 = 0
    cell_index7 = 0: cell_index8 = 0: cell_index9 = 0
    
    ' setting game completion parameters
    If gameOverScore = 4 Or gameOverScore = 5 Then Exit Function
    
    ' executing UP cell shifts
    output = CellShifting("index1", "index4", "index7", "score")
    output = CellShifting("index2", "index5", "index8", "score")
    output = CellShifting("index3", "index6", "index9", "score")
    
    ' checking and clearing board situation
    Clearing
    
End Function

Function PressDown()
    
    ' same layout as UP
    
    Dim output As Variant

    moveCell = False
    countCell = 0
    cell_index1 = 0: cell_index2 = 0: cell_index3 = 0
    cell_index4 = 0: cell_index5 = 0: cell_index6 = 0
    cell_index7 = 0: cell_index8 = 0: cell_index9 = 0
    
    If gameOverScore = 4 Or gameOverScore = 5 Then Exit Function
    
    ' compare A to B, B to C, C to A
    output = CellShifting("index7", "index4", "index1", "score")
    output = CellShifting("index8", "index5", "index2", "score")
    output = CellShifting("index9", "index6", "index3", "score")
    
    Clearing
    
End Function

Function PressLeft()

    ' same layout as UP

    Dim output As Variant

    moveCell = False
    countCell = 0
    cell_index1 = 0: cell_index2 = 0: cell_index3 = 0
    cell_index4 = 0: cell_index5 = 0: cell_index6 = 0
    cell_index7 = 0: cell_index8 = 0: cell_index9 = 0
    
    If gameOverScore = 4 Or gameOverScore = 5 Then Exit Function
    
    ' compare A to B, B to C, C to A
    output = CellShifting("index1", "index2", "index3", "score")
    output = CellShifting("index4", "index5", "index6", "score")
    output = CellShifting("index7", "index8", "index9", "score")
    
    Clearing
    
End Function

Function PressRight()
    
    ' same layout as UP
    
    Dim output As Variant
    
    moveCell = False
    countCell = 0
    cell_index1 = 0: cell_index2 = 0: cell_index3 = 0
    cell_index4 = 0: cell_index5 = 0: cell_index6 = 0
    cell_index7 = 0: cell_index8 = 0: cell_index9 = 0
    
    If gameOverScore = 4 Or gameOverScore = 5 Then Exit Function

    ' compare A to B, B to C, C to A
    output = CellShifting("index3", "index2", "index1", "score")
    output = CellShifting("index6", "index5", "index4", "score")
    output = CellShifting("index9", "index8", "index7", "score")
    
    Clearing
    
End Function

Sub ButtonPressUp()

    ' This subroutine performs the UP function and adds results from the action to corresponding data sheets

    ' up action
    PressUp
    
    ' updating moves count
    moveCount = moveCount + 1
    Range("moves_count").Value = moveCount
    
    ' add move into UserMovesList
    Sheets("UserMovesList").Range(movesCell).Offset(rowUML - 1, colUML) = "U"
    colUML = colUML + 1 ' when a user/bot is playing the game, the history of moves is horizontal; vertical is each game started
    ' could keep a log of games to track progress, and show chart of progress
    
    ' add board values to UserValuePositionList
    rowUVL = Sheets("variableStorage").Range("B3") ' change value in case altered by Undo
    rowUVL = rowUVL + 1
    Sheets("variableStorage").Range("B3") = rowUVL ' saving value
    StoreDataInUserValueList
    
    CellColorModifier ' ensuring cells follow incremental color adjustment based on its value
    
End Sub

Sub ButtonPressDown()

    PressDown
    moveCount = moveCount + 1
    Range("moves_count").Value = moveCount
    
    ' add move into UserMovesList
    Sheets("UserMovesList").Range(movesCell).Offset(rowUML - 1, colUML) = "D"
    colUML = colUML + 1
    
    ' add board values to UserValuePositionList
    rowUVL = Sheets("variableStorage").Range("B3") ' change value in case altered by Undo
    rowUVL = rowUVL + 1
    Sheets("variableStorage").Range("B3") = rowUVL ' saving value
    StoreDataInUserValueList
    
    CellColorModifier
    
End Sub

Sub ButtonPressLeft()

    PressLeft
    moveCount = moveCount + 1
    Range("moves_count").Value = moveCount
    
    ' add move into UserMovesList
    Sheets("UserMovesList").Range(movesCell).Offset(rowUML - 1, colUML) = "L"
    colUML = colUML + 1
    
    ' add board values to UserValuePositionList
    rowUVL = Sheets("variableStorage").Range("B3") ' change value in case altered by Undo
    rowUVL = rowUVL + 1
    Sheets("variableStorage").Range("B3") = rowUVL ' saving value
    StoreDataInUserValueList
    
    CellColorModifier
    
End Sub

Sub ButtonPressRight()

    PressRight
    moveCount = moveCount + 1
    Range("moves_count").Value = moveCount
    
    ' add move into UserMovesList
    Sheets("UserMovesList").Range(movesCell).Offset(rowUML - 1, colUML) = "R"
    colUML = colUML + 1
    
    ' add board values to UserValuePositionList
    rowUVL = Sheets("variableStorage").Range("B3") ' change value in case altered by Undo
    rowUVL = rowUVL + 1
    Sheets("variableStorage").Range("B3") = rowUVL ' saving value
    StoreDataInUserValueList
    
    CellColorModifier
    
End Sub

Function StoreDataInUserValueList()
    
    ' This function stores board values into the UserValuePositionList
    
    Sheets("UserValuePositionList").Range("A1").Offset(rowUVL, 0) = Range("index1")
    Sheets("UserValuePositionList").Range("A1").Offset(rowUVL, 1) = Range("index2")
    Sheets("UserValuePositionList").Range("A1").Offset(rowUVL, 2) = Range("index3")
    Sheets("UserValuePositionList").Range("A1").Offset(rowUVL, 3) = Range("index4")
    Sheets("UserValuePositionList").Range("A1").Offset(rowUVL, 4) = Range("index5")
    Sheets("UserValuePositionList").Range("A1").Offset(rowUVL, 5) = Range("index6")
    Sheets("UserValuePositionList").Range("A1").Offset(rowUVL, 6) = Range("index7")
    Sheets("UserValuePositionList").Range("A1").Offset(rowUVL, 7) = Range("index8")
    Sheets("UserValuePositionList").Range("A1").Offset(rowUVL, 8) = Range("index9")
    Sheets("UserValuePositionList").Range("A1").Offset(rowUVL, 9) = Range("score")

End Function


Function CellColorModifier()

    ' logic: value of each cell will have corresponding color based on magnitude of value
    ' i.e. 2 will have default board color, 4 will have a higher saturation of that color, and so on
    
    '' default: RGB(252, 228, 214)
    '' red: RGB(255, 79, 79)
    '' green: RGB(169, 208, 142)
    '' blue: RGB(155, 194, 230)
    
    Dim cell As Range
    Dim incrementalValue As Long ' value to be used in adjusting the RGB color values
    
    incrementalValue = 0
    For Each cell In Range("E4:G6")
        incrementalValue = cell.Value / 2 * 2 ' using 2 as default seed value to adjust the incrementalValue adjustor for the color
        ' ensuring color values do not go beyond 0 or 255 range using Max/Min function
        If Sheets("Advanced").Range("color").Value = "default" Then
            cell.Interior.Color = RGB(WorksheetFunction.Max(0, 252 - incrementalValue), WorksheetFunction.Max(0, 228 - incrementalValue), WorksheetFunction.Max(0, 214 - incrementalValue))
        End If
        If Sheets("Advanced").Range("color").Value = "red" Then
            cell.Interior.Color = RGB(WorksheetFunction.Max(0, 255 - incrementalValue), WorksheetFunction.Min(255, 79 + incrementalValue), WorksheetFunction.Max(0, 79 - incrementalValue))
        End If
        If Sheets("Advanced").Range("color").Value = "green" Then
            cell.Interior.Color = RGB(WorksheetFunction.Min(255, 169 + incrementalValue), WorksheetFunction.Max(0, 208 - incrementalValue), WorksheetFunction.Max(0, 142 - incrementalValue))
        End If
        If Sheets("Advanced").Range("color").Value = "blue" Then
            cell.Interior.Color = RGB(WorksheetFunction.Max(0, 155 - incrementalValue), WorksheetFunction.Min(255, 194 + incrementalValue), WorksheetFunction.Max(0, 230 - incrementalValue))
        End If
    Next cell
    
End Function


' unused code for making arrow key movements

'Function TrueArrowKeyMovement()

'    Application.OnKey "{UP}", "ButtonPressUp"
'    Application.OnKey "{DOWN}", "ButtonPressDown"
'    Application.OnKey "{LEFT}", "ButtonPressLeft"
'    Application.OnKey "{RIGHT}", "ButtonPressRight"

'End Function

'Function FalseArrowKeyMovement()

'    Application.OnKey "{UP}"
'    Application.OnKey "{DOWN}"
'    Application.OnKey "{LEFT}"
'    Application.OnKey "{RIGHT}"

'End Function

'Sub ToggleArrowKey()

    
    ' instead of using booleans and keep changing between true and false,
    ' i am incrementally increasing it throughout program just based on modulus
'    TrueArrowKeyMovement
    'arrowKeyToggle = arrowKeyToggle + 1
    'If arrowKeyToggle Mod 2 = 0 Then ' TrueArrowKeyMovement
    '    Sheets("variableStorage").Range("B4") = "True"
    '    ActiveCell = Sheets("Advanced").Range("K5")
    'End If
    
    'If arrowKeyToggle Mod 2 <> 0 Then
    '    Sheets("variableStorage").Range("B4") = "False"
    'End If
        'TrueArrowKeyMovement
        
    'If arrowKeyToggle Mod 2 <> 0 Then FalseArrowKeyMovement
    
    'Debug.Print arrowKeyToggle
    
'End Sub
