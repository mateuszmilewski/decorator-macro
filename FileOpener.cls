VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FileOpener"
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



Private filename As String
Private filenames As Variant ' is array usually!
Private separator As String


Private fso As FileSystemObject
Private fldPth As String
Private fld As Folder
Private flds As Folders


Private wrk As Workbook


Public Property Get getFolder() As Folder
    Set getFolder = fld
End Property


Public Property Get getFolders() As Folders
    Set getFolders = flds
End Property

Public Property Get getFolderPath() As String
    getFolderPath = CStr(fldPth)
End Property



Public Sub constructor()
    
    ' ---------------------------
    ' no operation required!
    ' ---------------------------
End Sub


Public Function dataAvailable() As Boolean
    
    dataAvailable = False
End Function


Public Sub openOneFile()
    ChDir ThisWorkbook.Path
    filename = Application.GetOpenFilename(, , "get the file", "get input file", False)
    DoEvents
    
    If filename <> False Then
    End If

End Sub


Public Sub openManyFiles()
    ChDir ThisWorkbook.Path
    filenames = Application.GetOpenFilename(, , "get the files", "get input files", MultiSelect:=True)
    DoEvents
    
    If IsArray(filenames) Then
        ' Debug.Print filenames(1)
        Dim x As Variant
        For x = LBound(filenames) To UBound(filenames)
            ' Debug.Print filenames(x) ' OK!
        Next x
    End If
End Sub


Public Sub setupFolder()
    
    ChDir ThisWorkbook.Path
    
    
    Dim sFolder As String
    
    ' Open the select folder prompt
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select a Folder"
        If .show = -1 Then ' if OK is pressed
            sFolder = .SelectedItems(1)
        End If
    End With
    
    If sFolder <> "" Then ' if a file was chosen
        ' *********************
        
        fldPth = sFolder
        Set f = fso.getFolder(fldPth)
        
        ' *********************
    End If
    
    
End Sub



Public Function passOpenedFiles() As Variant
    Set passOpenedFiles = Nothing
End Function


Private Sub Class_Initialize()
    filename = ""
    separator = Application.PathSeparator
    Set filenames = New Dictionary
    Set fso = New FileSystemObject
    Set fld = Nothing
    Set flds = Nothing
    fldPth = ""
End Sub

Private Sub Class_Terminate()
    filename = ""
    Set filenames = Nothing
    Set fso = Nothing
    
    Set fld = Nothing
    Set flds = Nothing
    
    fldPth = ""
End Sub
