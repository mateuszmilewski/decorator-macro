VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UploadHandler"
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


' this class defines object that working only as visual without real business logic here
' it is independent demo version

Private f As Workbook
Private sh As Worksheet


Private corailScreen As String
Private cellButton As Range

Public Sub cfg(mcorailScreen As String, mcellButton As Range)
    
    corailScreen = mcorailScreen
    Set cellButton = mcellButton
    cellButton.Interior.Color = RGB(200, 200, 200)
End Sub


Private Sub Class_Initialize()
    Set file = Nothing
    Set sh = Nothing
End Sub

Private Sub Class_Terminate()
    Set sh = Nothing
    Set file = Nothing
End Sub


Public Sub TEST_GREEN_uploadFile()
    
    Debug.Print "uploading test file"
    cellButton.Interior.Color = RGB(255, 255, 0)
    Sleep 1000
    cellButton.Interior.Color = RGB(0, 255, 0)
End Sub

Public Sub TEST_RED_uploadFile()
    
    Debug.Print "uploading test file"
    cellButton.Interior.Color = RGB(255, 255, 0)
    Sleep 1000
    cellButton.Interior.Color = RGB(255, 0, 0)
End Sub
