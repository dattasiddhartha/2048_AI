Attribute VB_Name = "ColourMod"
Option Explicit

' This macro will change the colour scheme of the Advanced game layout (only basic colours for now, not including charts, etc)

Function ColourSchemeAdjustment()

    ' Case will try to match text in the color range cell to return a predefined color scheme
    Select Case Sheets("Advanced").Range("color").Value
        Case "default"
            ' color change of main board
            Range("index1").Interior.Color = RGB(252, 228, 214)
            Range("index2").Interior.Color = RGB(252, 228, 214)
            Range("index3").Interior.Color = RGB(252, 228, 214)
            Range("index4").Interior.Color = RGB(252, 228, 214)
            Range("index5").Interior.Color = RGB(252, 228, 214)
            Range("index6").Interior.Color = RGB(252, 228, 214)
            Range("index7").Interior.Color = RGB(252, 228, 214)
            Range("index8").Interior.Color = RGB(252, 228, 214)
            Range("index9").Interior.Color = RGB(252, 228, 214)
            
            ' color change of stats board
            Sheets("Advanced").Range("B9:E15").Interior.Color = RGB(221, 235, 247)
            
            ' color change of changeables board
            Sheets("Advanced").Range("G9:J15").Interior.Color = RGB(226, 239, 218)
            
        Case "red"
            ' color change of main board
            Range("index1").Interior.Color = RGB(255, 79, 79)
            Range("index2").Interior.Color = RGB(255, 79, 79)
            Range("index3").Interior.Color = RGB(255, 79, 79)
            Range("index4").Interior.Color = RGB(255, 79, 79)
            Range("index5").Interior.Color = RGB(255, 79, 79)
            Range("index6").Interior.Color = RGB(255, 79, 79)
            Range("index7").Interior.Color = RGB(255, 79, 79)
            Range("index8").Interior.Color = RGB(255, 79, 79)
            Range("index9").Interior.Color = RGB(255, 79, 79)
            
            ' color change of stats board
            Sheets("Advanced").Range("B9:E15").Interior.Color = RGB(212, 98, 98)
            
            ' color change of changeables board
            Sheets("Advanced").Range("G9:J15").Interior.Color = RGB(246, 222, 222)
            
        Case "green"
            ' color change of main board
            Range("index1").Interior.Color = RGB(169, 208, 142)
            Range("index2").Interior.Color = RGB(169, 208, 142)
            Range("index3").Interior.Color = RGB(169, 208, 142)
            Range("index4").Interior.Color = RGB(169, 208, 142)
            Range("index5").Interior.Color = RGB(169, 208, 142)
            Range("index6").Interior.Color = RGB(169, 208, 142)
            Range("index7").Interior.Color = RGB(169, 208, 142)
            Range("index8").Interior.Color = RGB(169, 208, 142)
            Range("index9").Interior.Color = RGB(169, 208, 142)
            
            ' color change of stats board
            Sheets("Advanced").Range("B9:E15").Interior.Color = RGB(84, 130, 53)
            
            ' color change of changeables board
            Sheets("Advanced").Range("G9:J15").Interior.Color = RGB(198, 224, 180)
            
        Case "blue"
            ' color change of main board
            Range("index1").Interior.Color = RGB(155, 194, 230)
            Range("index2").Interior.Color = RGB(155, 194, 230)
            Range("index3").Interior.Color = RGB(155, 194, 230)
            Range("index4").Interior.Color = RGB(155, 194, 230)
            Range("index5").Interior.Color = RGB(155, 194, 230)
            Range("index6").Interior.Color = RGB(155, 194, 230)
            Range("index7").Interior.Color = RGB(155, 194, 230)
            Range("index8").Interior.Color = RGB(155, 194, 230)
            Range("index9").Interior.Color = RGB(155, 194, 230)
            
            ' color change of stats board
            Sheets("Advanced").Range("B9:E15").Interior.Color = RGB(47, 117, 181)
            
            ' color change of changeables board
            Sheets("Advanced").Range("G9:J15").Interior.Color = RGB(189, 215, 238)
        End Select
        
End Function
