VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PCVData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IData




Private fso As FileSystemObject
Private f As Folder
Private pcvPath1rng As Range
Private fo As FileOpener


Private cnn1 As Connection1




Private Sub Class_Initialize()
    Set f = Nothing
    Set fso = New FileSystemObject
    Set fo = New FileOpener
    Set pcvPath1rng = ThisWorkbook.Sheets("register").Range("pcvPath1")
    
    Set cnn1 = New Connection1
    cnn1.connectWith "rawPCV"
End Sub

Private Sub Class_Terminate()
    Set f = Nothing
    Set fso = Nothing
    Set fo = Nothing
    Set pcvPath1rng = Nothing
    Set cnn1 = Nothing
End Sub

Private Sub IData_refreshConnection()
    If Not cnn1 Is Nothing Then
        cnn1.refreshConnectionWith
    End If
End Sub

Private Sub IData_setFolder()

    
    If pcvPath1rng.Value <> "" Then
    
        Dim ans As Variant
        ans = MsgBox("Do you want to continue with source data placed in: " & pcvPath1rng.Value, vbYesNo)
        
        If ans = vbYes Then
            innerSetFolder
        Else
            chooseNewLocation
            innerSetFolder
        End If
        
    Else
        ' you need to make a new one
        chooseNewLocation
        innerSetFolder
    End If
    

End Sub


Private Sub innerSetFolder()
    Set f = fso.getFolder(pcvPath1rng.Value)
End Sub


Private Sub chooseNewLocation()
    
    fo.setupFolder
    If Not pcvPath1rng Is Nothing Then
        pcvPath1rng.Value = fo.getFolderPath
    Else
        MsgBox "sth went wrong while trying add a new location to PCV files!", vbCritical
    End If
    
End Sub

Private Property Get IData_sourceFolder() As Scripting.IFolder

End Property
