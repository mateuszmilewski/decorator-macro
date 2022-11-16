Attribute VB_Name = "TestMainModule"
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



' this module having some proc and functions
' that was initally used for making some visualisation on MAIN sheet -> use cells (resized to be squares) - i think this will be obsolete - to be studied
' which are

' TEST MAIN
Private Sub adaptCorailData(t As Range)
    
    If checkIfAllInputFilesAreGreen() Then
        
        t.Interior.Color = RGB(0, 255, 0)
        
        ' some operations for adaptation / decoration
        Sleep 1000
        
        
        MsgBox "Adaptation COMPLETED!", vbInformation
        
        
    Else
        t.Interior.Color = RGB(255, 0, 0)
        
        
        MsgBox "Adaptation FAILED!", vbCritical
    End If
    
End Sub


Private Function checkIfAllInputFilesAreGreen() As Boolean
    
    checkIfAllInputFilesAreGreen = TEST_checkColors()
End Function


Private Function TEST_checkColors() As Boolean
    
    TEST_checkColors = False
    Sleep 1000
    
    
    Dim b1355 As Boolean
    Dim b1320 As Boolean
    Dim bnet As Boolean
    
    ' check 1355
    If ThisWorkbook.Sheets("MAIN").Range("macro1355").Interior.Color = RGB(0, 255, 0) Then
        b1355 = True
    Else
        b1355 = False
    End If
    
    
    ' check 1320
    If ThisWorkbook.Sheets("MAIN").Range("macro1320").Interior.Color = RGB(0, 255, 0) Then
        b1320 = True
    Else
        b1320 = False
    End If
    
    
    ' check net-needs
    If ThisWorkbook.Sheets("MAIN").Range("macroNetNeeds").Interior.Color = RGB(0, 255, 0) Then
        bnet = True
    Else
        bnet = False
    End If
    
    Sleep 1000
    
    If b1355 And b1320 And bnet Then
        TEST_checkColors = True
    Else
        TEST_checkColors = False
        
    End If
    
End Function
