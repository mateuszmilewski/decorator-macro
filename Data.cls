VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Data"
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



' this class is just a demo without any specific business logic
' just to see how interface is interacting!

'
'

Implements IFile




' for now those fields are examples what most probably is important
Private r As Range
Private sh As Worksheet
Private wrk As Workbook

' key data pairted within custom key
Private d As Dictionary




Private Sub Class_Initialize()
    Set r = Nothing
    Set sh = Nothing
    Set wrk = Nothing
    Set d = Nothing
End Sub

Private Sub Class_Terminate()
    Set r = Nothing
    Set sh = Nothing
    Set wrk = Nothing
    Set d = Nothing
End Sub



Private Function IFile_gD() As Scripting.IDictionary
    Set IFile_gD = d
End Function

Private Function IFile_gRng() As Range
    Set IFile_gRng = r
End Function

Private Property Get IFile_gSh() As Worksheet
    Set r = r1
End Property

Private Property Get IFile_gWrk() As Workbook
    Set IFile_gWrk = wrk
End Property

Private Sub IFile_sD(d1 As Scripting.IDictionary)
    Set d = d1
End Sub

Private Sub IFile_sRng(r1 As Range)
    Set r = r1
End Sub

Private Property Set IFile_sSh(RHS As Worksheet)
    Set sh = RHS
End Property

Private Property Set IFile_sWrk(RHS As Workbook)
    Set wrk = RHS
End Property
