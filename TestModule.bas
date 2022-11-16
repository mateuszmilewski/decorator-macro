Attribute VB_Name = "TestModule"
'The MIT License (MIT)
'
'Copyright (c) 2022 FORREST
' Mateusz Milewski mateusz@stellantis.com aka FORREST
'
'Permission is hereby granted, free of charge, to any person obtaining a copy
'of this software and associated documentation files (the "Software"), to deal
'in the Software without restriction, including without limitation the rights
'to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
'copies of the Software, and to permit persons to whom the Software is
'furnished to do so, subject to the following conditions:
'
'The above copyright notice and this permission notice shall be included in all
'copies or substantial portions of the Software.
'
'THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
'IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
'FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
'AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
'LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
'OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
'SOFTWARE.





' RECORDINGS FROM THE MAIN FLOW MODULE
' ====================================================================================
' ====================================================================================



' this one below is just pure macro record


Private Sub test_makeShapeTest1()
'
' makeShapeTest1 Macro
'

'
    ActiveSheet.Shapes.AddShape(msoShapeRectangle, 183.75, 134.25, 123.75, 64.5). _
        Select
    Selection.ShapeRange.name = "Corail1355"
    Selection.name = "Corail1355"
    Selection.ShapeRange.TextFrame2.VerticalAnchor = msoAnchorMiddle
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "data from 1355"
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 14). _
        ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignCenter
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 14).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 16
        .name = "+mn-lt"
    End With
    Selection.ShapeRange.IncrementLeft -33.75
    Selection.ShapeRange.IncrementTop -26.25
    With Selection.ShapeRange.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 167, 166)
        .Transparency = 0
        .Solid
    End With
    Selection.ShapeRange.TextFrame2.TextRange.Font.Bold = msoTrue
    Range("H11").Select
End Sub






Private Sub test__construct()


    Dim flowSh As Worksheet
    Set flowSh = ThisWorkbook.Sheets(G_SH_FLOW)

    Decorator.ClearModule.clearDashboard
    
    Dim shs As Decorator.ShapesHandler
    Set shs = New Decorator.ShapesHandler
    
    Dim s As Decorator.IShape
    Set s = New Decorator.ShapeHandler
    With s
        .predefine
        .create
    End With
    
    
    shs.update
    
    shs.quickLog
    
    Set shs = Nothing
End Sub

' ====================================================================================
' ====================================================================================






Private Sub GroupByTest()
'
' GroupByTest Macro
'

'
    ActiveSheet.Shapes.Range(Array("corailInput1355")).Select
    ActiveSheet.Shapes.Range(Array("corailInput1355", "corailInput1320")). _
        Select
    Selection.ShapeRange.group.Select
    Selection.ShapeRange.Ungroup.Select
    Range("H13").Select
End Sub



Private Sub testConnectWithArrow()
'
' testConnectWithArrow Macro
'

'
    ActiveSheet.Shapes.AddConnector(msoConnectorStraight, 269, 121, 500, 180). _
        Select
        Selection.ShapeRange.Line.EndArrowheadStyle = msoArrowheadTriangle
    Selection.ShapeRange.ConnectorFormat.BeginConnect ActiveSheet.Shapes( _
        "corailInput1355"), 4
    Selection.ShapeRange.ConnectorFormat.EndConnect ActiveSheet.Shapes( _
        "DECORATOR_TOOL"), 2
    Selection.ShapeRange.ShapeStyle = msoLineStylePreset15
    Range("L2:M2").Select
End Sub


Private Sub Macro1()
'
' Macro1 Macro
'

'
    ActiveSheet.Shapes.Range(Array("G_corailInput1355")).Select
    Selection.OnAction = "forAssignTest"
    Range("L17").Select
End Sub



Private Sub assignMacroToSpecificShape()
    
    Dim sh As Worksheet, s As ShapeRange
    Set sh = ThisWorkbook.Sheets("FLOW")
    Set s = sh.Shapes.Range(Array("G_corailInput1355"))
    s.Select
    Selection.OnAction = "forAssignTest"
    sh.Range("A1").Select
End Sub



Private Sub forAssignTest()
    ' MsgBox "you should see this after click on shape", vbInformation
    
    With ShapeForm
        .Caption = "Corail 1355 data"
        .show
        
    End With
End Sub



Private Sub property__TEST()
    
    Dim a As IFile
    Set a = New Data
    Set a.sSh = ThisWorkbook.ActiveSheet
    Dim s As Worksheet
    Set s = a.gSh
    Debug.Print s.name & " " & a.gSh.name
    
End Sub


