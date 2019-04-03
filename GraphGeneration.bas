Attribute VB_Name = "GraphGeneration"
Option Explicit

' This module contains all graph related subroutines

Sub generateScoreChart()

    ' This subroutine uses all historical data stored on prior user moves, and generates a line chart of score progression

    Dim scoreChart As ChartObject
    Dim rowUML As Integer

    ' setting position of chart on spreadsheet display
    Set scoreChart = ActiveSheet.ChartObjects.Add(Top:=Range("R2").Top, Left:=Range("R2").Left, Width:=50 * 10, Height:=50 * 4)
        
    ' Find last row (non-zero)
    rowUML = CInt(Sheets("variableStorage").Range("B2")) ' calling last row variable stored from across another module
    
    ' accessing chart data from UserMovesList
    scoreChart.Chart.SetSourceData Sheets("UserMovesList").Range("C2", Sheets("UserMovesList").Range("C2").Offset(rowUML, 0))
    
    scoreChart.Chart.ChartType = xlLine ' setting as line chart
    
    ' Title, legend, axes
    scoreChart.Chart.HasTitle = True
    scoreChart.Chart.ChartTitle.Text = "Score over time"
    
    scoreChart.Chart.HasLegend = False
    
    scoreChart.Chart.Axes(xlCategory).HasTitle = True
    scoreChart.Chart.Axes(xlCategory).AxisTitle.Caption = "Game"
    
    scoreChart.Chart.Axes(xlValue).HasTitle = True
    scoreChart.Chart.Axes(xlValue).AxisTitle.Caption = "Score"
    
    
End Sub

' generate pattern of U L R D
Sub generatePatternOverviewChart()
    
    ' This subroutine will tell users their move selection pattern of Up, Down, Left or Right
    
    Dim patternChart As ChartObject

    Set patternChart = ActiveSheet.ChartObjects.Add(Top:=Range("R7").Top, Left:=Range("R7").Left, Width:=50 * 10, Height:=50 * 5)
     
    ' count all occurrences of UDLR & Storing it on a worksheet to use as data for plotting
    Sheets("patternPlot").Range("B1") = Application.CountIf(Worksheets("UserMovesList").Cells, "U")
    Sheets("patternPlot").Range("B2") = Application.CountIf(Worksheets("UserMovesList").Cells, "D")
    Sheets("patternPlot").Range("B3") = Application.CountIf(Worksheets("UserMovesList").Cells, "L")
    Sheets("patternPlot").Range("B4") = Application.CountIf(Worksheets("UserMovesList").Cells, "R")
    
    Sheets("patternPlot").Range("A1") = "U"
    Sheets("patternPlot").Range("A2") = "D"
    Sheets("patternPlot").Range("A3") = "L"
    Sheets("patternPlot").Range("A4") = "R"
     
     ' plot using data stored
    patternChart.Chart.SetSourceData Sheets("patternPlot").Range("A1:B4")
    
    
    ' Title, legend, axes
    patternChart.Chart.HasTitle = True
    patternChart.Chart.ChartTitle.Text = "Moves occurrence"
    
    patternChart.Chart.HasLegend = False
    
    patternChart.Chart.Axes(xlCategory).HasTitle = True
    patternChart.Chart.Axes(xlCategory).AxisTitle.Caption = "Move"
    
    patternChart.Chart.Axes(xlValue).HasTitle = True
    patternChart.Chart.Axes(xlValue).AxisTitle.Caption = "Frequency"
    
    
End Sub
