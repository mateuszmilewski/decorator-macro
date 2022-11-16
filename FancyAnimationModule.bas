Attribute VB_Name = "FancyAnimationModule"
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




Public Sub makeFancyAnimationOnShapes()
    
    Application.Run "internalForCall_1355"
    Application.Run "internalForCall_1320"
    Application.Run "internalForCall_NN"
    
End Sub

Private Sub internalForCall_1355()
    internalForCall Decorator.G_1355, emodRectCorail1355
End Sub

Private Sub internalForCall_1320()
    internalForCall Decorator.G_1320, emodRectCorail1320
End Sub

Private Sub internalForCall_NN()
    
    internalForCall Decorator.G_NN, emodRectCorailNN
End Sub

Private Sub internalForCall(nm As String, e1 As EMOD)

    Dim m As ShapeModifier, sh1 As Shape
    Set m = New ShapeModifier
    m.modForAnimationTarget e1
    Set sh1 = ThisWorkbook.Sheets("FLOW").Shapes("G_" & nm)
        
        
    Do
        
        If sh1.top > m.top Then
            sh1.top = sh1.top - 1
        Else
            sh1.top = sh1.top + 1
        End If
        
        
        If sh1.left > m.left Then
            sh1.left = sh1.left - 1
        Else
            sh1.left = sh1.left + 1
        End If
        
        Sleep 1
        DoEvents
    Loop Until targetReached(m, sh1)
End Sub


Private Function targetReached(m As ShapeModifier, s As Shape) As Boolean

    targetReached = False

    If Math.Abs(m.top - s.top) < 2 Then
        If Math.Abs(m.left - s.left) < 2 Then
            targetReached = True
        End If
    End If
End Function
