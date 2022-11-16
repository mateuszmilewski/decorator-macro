Attribute VB_Name = "ClearModule"
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



' DIRECT SH FLOW CLEARING
' ------------------------------
' ------------------------------

' this 2 subs are binded - this all about graphical FLOW sh
' at first version this was a demo to represent the connection between the data
' so we have a simple shapes controlled basically by the instances of the ShapesHandler
' clear dashboard means in fact clear FLOW sh from those shapes
' this sub are directly working with shapes
' maybe in future to consider put it inside some more abstract object !
' -----------------------------------------------------------------------------------------------
Public Sub ribbonClearDashboard(ictrl As IRibbonControl)
    clearDashboard
End Sub


Public Sub clearDashboard()
    
    Dim flowSh As Worksheet
    Set flowSh = ThisWorkbook.Sheets(G_SH_FLOW)
    
    ' not really fancy - just directly clearing FLOW sheet from the shapes that was created 'in previous action
    ' no any IF statement no nothing - clear and kill everybody
    
    Dim sh1 As Shape
    For Each sh1 In flowSh.Shapes
        sh1.Delete
    Next sh1
    
    
End Sub
' -----------------------------------------------------------------------------------------------


' ------------------------------
' ------------------------------





' those subs below are resp for MAIN sh - might be that in future this will be obsolete

Public Sub clearBuffer(t As Range)
    Debug.Print "running clear buffer logic"
    
    Dim r As Range
    Set r = ThisWorkbook.Sheets("MAIN").Range("macroClear")
    
    
    r.Interior.Color = RGB(200, 200, 200)
    
    Sleep 1000
    
    r.Interior.Color = RGB(0, 255, 0)
    
    
    
    changeToGreyForAllUplaods
    changeToGreyForAdapt
    
    
    MsgBox "BUFFER CLEARED!", vbInformation
End Sub


Private Sub changeToGreyForAdapt()
    Dim r As Range
    Set r = ThisWorkbook.Sheets("MAIN").Range("adaptCorailData")
    
    r.Interior.Color = RGB(200, 200, 200)
End Sub


Private Sub changeToGreyForAllUplaods()
    
    Dim r As Range
    Set r = ThisWorkbook.Sheets("MAIN").Range(ThisWorkbook.Sheets("MAIN").Range("macro1355"), ThisWorkbook.Sheets("MAIN").Range("macroNetNeeds"))
    
    r.Interior.Color = RGB(200, 200, 200)
End Sub
