VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} DECORATOR_FORM 
   Caption         =   "DECORATOR"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4440
   OleObjectBlob   =   "DECORATOR_FORM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "DECORATOR_FORM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub BtnGenerateDarwinFiles_Click()
    hide
    makeExports
    
    
    MsgBox "READY!", vbInformation
End Sub
