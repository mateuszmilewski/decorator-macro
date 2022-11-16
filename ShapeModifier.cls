VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapeModifier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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



' i think this should be simple : -> ' what we want to modify?
Public name As String
Public top As Single
Public left As Single
Public width As Single
Public height As Single
Public shapeType As MsoAutoShapeType
Public txt As String
Public shapeBgColor As Long





Public Function useful() As Boolean
    
    useful = False
    
    If name <> "" Then
    
        If top > 0 And left > 0 And width > 0 And height > 0 And shapeType > 0 Then
            
            useful = True
        End If
    End If
End Function


Public Sub modForAnimationTarget(e1 As EMOD)
    
    If e1 = emodRectCorail1355 Then
        
        top = 80
        left = 120
    ElseIf e1 = emodRectCorail1320 Then
        top = 160
        left = 120
    ElseIf e1 = emodRectCorailNN Then
        top = 240
        left = 120
    End If
End Sub


Public Sub partialpredef(Optional e1 As EMOD)
    
    
    If e1 = emodRectCorail Then
        
        ' corail color: 255,167,166
    ElseIf e1 = emodRectCorail1355 Then
    
        
        Me.shapeType = msoShapeRectangle
        name = G_1355
        top = 100
        left = 100
        width = 140
        height = 60
        shapeBgColor = RGB(255, 167, 166)
        txt = "data fom 1355 screen"
        
    ElseIf e1 = emodRectCorail1320 Then
    
        Me.shapeType = msoShapeRectangle
        name = G_1320
        top = 200
        left = 100
        width = 140
        height = 60
        shapeBgColor = RGB(255, 167, 166)
        txt = "data fom 1320 screen"
        
    ElseIf e1 = emodRectCorailNN Then
    
        Me.shapeType = msoShapeRectangle
        name = Decorator.G_NN
        top = 300
        left = 140
        width = 240
        height = 60
        shapeBgColor = RGB(255, 167, 166)
        txt = "custom export from Corail based on FBB film and avail. stock"
        
    ElseIf e1 = emoddecorator Then
    
        Me.shapeType = msoShapeRectangle
        name = Decorator.G_DECORATOR
        top = 40
        left = 400
        width = 300
        height = 100
        shapeBgColor = RGB(102, 0, 102)
        txt = "DECORATOR"
        
    ElseIf e1 = emodpcv Then
        
        Me.shapeType = msoShapeRectangle
        name = Decorator.G_PCV
        top = 400
        left = 100
        width = 140
        height = 60
        shapeBgColor = RGB(0, 152, 69) ' PCV GREEN
        txt = "exported data from PCV"
        
    ElseIf e1 = emodlook Then
        
        Me.shapeType = msoShapeRectangle
        name = Decorator.G_LOOK
        top = 500
        left = 100
        width = 140
        height = 60
        shapeBgColor = RGB(0, 102, 0) ' look color
        txt = "exported data from LOOK"
        
    ElseIf e1 = emodRectWhite Then
        ' ok - let be default
        Me.shapeType = msoShapeRectangle
        '
        ' this one I have a doubts - excel allows to have shapes with exactly the same name
        ' no validation from excel site
        name = "default0white" ' - this is will be a default name but there is a necessity to have a second step of checking and potential re-asign of the shape name
        top = 100
        left = 100
        width = 123
        height = 66
        shapeBgColor = RGB(240, 240, 240)
    Else
        ' ok - let be default
        Me.shapeType = msoShapeRectangle
        '
        ' this one I have a doubts - excel allows to have shapes with exactly the same name
        ' no validation from excel site
        name = "default0" ' - this is will be a default name but there is a necessity to have a second step of checking and potential re-asign of the shape name
        top = 100
        left = 100
        width = 123
        height = 66
        shapeBgColor = RGB(200, 200, 200)
    End If
    
End Sub



Private Sub nop()
    
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
