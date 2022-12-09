Attribute VB_Name = "MainModule"
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




' very first step ! - open some files
' ==================================================

Private Sub tryToOpenFile()
    
    ' general opener - no need to specify anything - macro should recognize file by itself
    Dim fo As FileOpener, f As New Factory, fi As IFile, fis As Variant
    Set fo = f.newFileOpener()
    fo.openOneFile
    
    If fo.dataAvailable Then
        Set fi = fo.passOpenedFiles
    End If
End Sub

Private Sub tryToOpenManyFiles()
    
    ' maybe just take all file at once
    Dim fo As FileOpener, f As New Factory, fi As IFile
    Set fo = f.newFileOpener()
    fo.openManyFiles
    
    If fo.dataAvailable Then
        Set fis = fo.passOpenedFiles
    End If
End Sub



' ==================================================
