VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapesHandler"
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


Private szejps As Dictionary
Private flowSh As Worksheet
Private sh1 As Shape


Private shape1355 As Decorator.IShape
Private shape1320 As Decorator.IShape
Private shapeNetNeeds As Decorator.IShape
Private shapePCV As Decorator.IShape
Private shapeLOOK As Decorator.IShape
Private theDecoratorShape As Decorator.IShape
Private exportForPusShape As Decorator.IShape
Private exportForProcure As Decorator.IShape
Private exportForTeardown As Decorator.IShape




Private Sub Class_Initialize()
    Set szejps = New Dictionary
    Set flowSh = ThisWorkbook.Sheets(Decorator.G_SH_FLOW)
    ' connect with all available shapes right now
    With flowSh
        If .Shapes.Count > 0 Then
            For Each sh1 In .Shapes
                If sh1.name Like "G_*" Then
                    szejps.Add sh1.name, sh1
                Else
                    sh1.Delete
                    DoEvents
                End If
            Next sh1
        End If
    End With
    
End Sub

Private Sub Class_Terminate()

    Set szejps = Nothing
    Set flowSh = Nothing
End Sub



Public Sub quickLog()
    
    Dim kij As Variant
    
    If szejps.Count > 0 Then
        
        For Each kij In szejps.Keys
            Debug.Print szejps(kij).name
        Next
    Else
        Debug.Print "no avial. data"
    End If
End Sub




Public Sub update()
    Set szejps = New Dictionary
    Set flowSh = ThisWorkbook.Sheets(Decorator.G_SH_FLOW)
    
    ' connect with all available shapes right now
    
    With flowSh
    
        If .Shapes.Count > 0 Then
    
            For Each sh1 In .Shapes
                
                If sh1.name Like "G_*" Then
                
                    szejps.Add sh1.name, sh1
                Else
                    'sh1.Delete
                    DoEvents
                End If
            Next sh1
            
        End If
    End With
    
End Sub

Public Sub makePCV()
    
    Set shapePCV = New Decorator.ShapeHandler
    With shapePCV
        .predefine emodpcv
        
        Dim pcv1 As IData
        Set pcv1 = New PCVData
        pcv1.setFolder
        pcv1.refreshConnection
        
        
        If checkIfThisShapeIsAlreadyExisting(emodpcv) Then
            .connect2 Decorator.G_PCV
        Else
            .create
        End If
        
        
        ' ap
        .validate
        .group
        .defineOnAction
    End With
    update
End Sub

Public Sub makeLOOK()
    
    Set shapeLOOK = New Decorator.ShapeHandler
    With shapeLOOK
        .predefine emodlook
        
        Dim d As IFile
        Set d = New Data
        d.sRng ThisWorkbook.Sheets("register").Range("h2")
        .setData d
        
        If checkIfThisShapeIsAlreadyExisting(emodlook) Then
            .connect2 Decorator.G_LOOK
        Else
            .create
        End If
        
        
        ap
        .validate
        .group
        .defineOnAction
    End With
    
    
    
    update
End Sub


Public Sub makeDECORATOR()
    
    Set theDecoratorShape = New Decorator.ShapeHandler
    With theDecoratorShape
        .predefine emoddecorator
        
        
        If checkIfDecoratorHaveAvailableResourceForCorailScenario() Then
        
        
            Dim d As IFile
            Set d = New Data
            d.sD szejps
            .setData d
        
            ' visualisation that decorator handler now have power on shapes
            makeFancyAnimationOnShapes
        
            If checkIfThisShapeIsAlreadyExisting(emoddecorator) Then
                .connect2 Decorator.G_DECORATOR
            Else
                .create
            End If
            
            
            ap
            .validate
            .group
            .defineOnAction
            
            .establishConnections
            
        Else
            MsgBox "Not all data is available!", vbCritical
        End If
        
        

    End With
    
    
    
    update
End Sub


Public Sub make1355()
    
    Set shape1355 = New Decorator.ShapeHandler
    With shape1355
        .predefine emodRectCorail1355
        
        Dim d As IFile
        Set d = New Data
        d.sRng ThisWorkbook.Sheets("register").Range("h2")
        .setData d
        
        If checkIfThisShapeIsAlreadyExisting(emodRectCorail1355) Then
            .connect2 Decorator.G_1355
        Else
            .create
        End If
        
        
        ap
        .validate
        .group
        .defineOnAction
    End With
    
    
    
    update
End Sub


Public Sub make1320()
    
    Set shape1320 = New Decorator.ShapeHandler
    With shape1320
        .predefine emodRectCorail1320
        
        Dim d As IFile
        Set d = New Data
        d.sRng ThisWorkbook.Sheets("register").Range("h2")
        .setData d
        
        If checkIfThisShapeIsAlreadyExisting(emodRectCorail1320) Then
            .connect2 Decorator.G_1320
        Else
            .create
        End If
        
        ap
        .validate
        .group
        .defineOnAction
    End With
    update
End Sub

Public Sub makeNN()
    
    Set shapeNetNeeds = New Decorator.ShapeHandler
    With shapeNetNeeds
        .predefine emodRectCorailNN
        
        Dim d As IFile
        Set d = New Data
        d.sRng ThisWorkbook.Sheets("register").Range("h2")
        .setData d
        
        If checkIfThisShapeIsAlreadyExisting(emodRectCorailNN) Then
            .connect2 Decorator.G_NN
        Else
            .create
        End If
        
        ap
        .validate
        .group
        .defineOnAction
    End With
    update
End Sub


Public Sub makeNN__NOK()
    
    Set shapeNetNeeds = New Decorator.ShapeHandler
    With shapeNetNeeds
        .predefine emodRectCorailNN
        
        Dim d As IFile
        Set d = New Data
        d.sRng ThisWorkbook.Sheets("register").Range("h1")
        .setData d
        
        If checkIfThisShapeIsAlreadyExisting(emodRectCorailNN) Then
            .connect2 Decorator.G_NN
        Else
            .create
        End If
        
    
        
        ap
        .validate
        .group
        .defineOnAction
    End With
    update
End Sub



Private Function checkIfDecoratorHaveAvailableResourceForCorailScenario()
    checkIfDecoratorHaveAvailableResourceForCorailScenario = True
    
    checkIfDecoratorHaveAvailableResourceForCorailScenario = checkIfDecoratorHaveAvailableResourceForCorailScenario And _
        szejps.Exists("G_" & Decorator.G_1355) And _
        szejps.Exists("G_" & Decorator.G_1320) And _
        szejps.Exists("G_" & Decorator.G_NN)
    
End Function






Private Function checkIfThisShapeIsAlreadyExisting(e1 As EMOD) As Boolean
    checkIfThisShapeIsAlreadyExisting = False
    
    Dim kij As Variant
    If e1 = emodRectCorail1355 Then
    
        
        For Each kij In szejps.Keys
            
            If szejps(kij).name = "G_" & Decorator.G_1355 Then
                checkIfThisShapeIsAlreadyExisting = True
                Exit Function
            End If
        Next
    ElseIf e1 = emodRectCorail1320 Then
        For Each kij In szejps.Keys
            If szejps(kij).name = "G_" & Decorator.G_1320 Then
                checkIfThisShapeIsAlreadyExisting = True
                Exit Function
            End If
        Next
    ElseIf e1 = emodRectCorailNN Then
        For Each kij In szejps.Keys
            If szejps(kij).name = "G_" & Decorator.G_NN Then
                checkIfThisShapeIsAlreadyExisting = True
                Exit Function
            End If
        Next
        
    ElseIf e1 = emodpcv Then
        For Each kij In szejps.Keys
            If szejps(kij).name = "G_" & Decorator.G_PCV Then
                checkIfThisShapeIsAlreadyExisting = True
                Exit Function
            End If
        Next
    ElseIf e1 = emodlook Then
        For Each kij In szejps.Keys
            If szejps(kij).name = "G_" & Decorator.G_LOOK Then
                checkIfThisShapeIsAlreadyExisting = True
                Exit Function
            End If
        Next
        
    ElseIf e1 = emoddecorator Then
        For Each kij In szejps.Keys
            If szejps(kij).name = "G_" & Decorator.G_DECORATOR Then
                checkIfThisShapeIsAlreadyExisting = True
                Exit Function
            End If
        Next
        
    End If
End Function
