VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "EcheancesMOHandler"
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


Private topTxt As String


' helper
Private ir As Range


Private Sub Class_Initialize()
    topTxt = "Situation livraison par échéance par article (SECHEL)" & Chr(10) & Chr(10)
    
    Dim moLabels As Range
    Set moLabels = ThisWorkbook.Sheets("register").Range(ThisWorkbook.Sheets("register").Range("A2"), _
        ThisWorkbook.Sheets("register").Range("A2").End(xlDown))
        
    For Each ir In moLabels
        
        
        If ThisWorkbook.Sheets("register").Range("A2").End(xlDown).Row > ir.Row Then
            topTxt = topTxt & ir.Value & vbTab
        Else
            topTxt = topTxt & ir.Value & Chr(10) & Chr(10)
        End If
    Next ir
    
    Debug.Print topTxt
    
End Sub

Private Sub Class_Terminate()
    topTxt = ""
End Sub


Public Sub connectWith_(w As Workbook)
    
    If Not w Is Nothing Then
        
        ' do some stuff
    End If
End Sub
