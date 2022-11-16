VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ShapeHandler"
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



Implements Decorator.IShape


Private m1 As Decorator.ShapeModifier
Private s As Shape
Private sr As ShapeRange

' validation shape
Private vs As Shape
Private vsr As ShapeRange


' grouped shapes
Private gs As Shape
Private gsr As ShapeRange


Private ESCENARIO As EMOD


Private data1 As IFile





Private Sub Class_Initialize()
    Set m1 = Nothing
    Set s = Nothing
    Set sr = Nothing
    
    
    Set gs = Nothing
    Set gsr = Nothing
End Sub

Private Sub Class_Terminate()
    Set m1 = Nothing
    Set s = Nothing
    Set sr = Nothing
    
    Set gs = Nothing
    Set gsr = Nothing
End Sub







Private Sub IShape_define(m As Decorator.ShapeModifier)

End Sub



Private Sub IShape_connect2(name As String)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(Decorator.G_SH_FLOW)
    
    Set s = sh.Shapes(name)
    Set sr = sh.Shapes.Range(s.name)
    
    Set gs = sh.Shapes("G_" & name)
    Set gsr = sh.Shapes.Range(gs.name)
    
    
    
    gsr.Ungroup
    DoEvents
    
    On Error Resume Next
    sh.Shapes.Range("V_" & name).Delete
    DoEvents
    s.Visible = msoCTrue
    
    
    ' when moving shapes during re-connect they might have different places!
    s.top = m1.top
    s.left = m1.left
    DoEvents
End Sub


' obsolete - will not work properly in future
Private Sub IShape_connect()

    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(Decorator.G_SH_FLOW)
    
    If ESCENARIO = emodRectCorail1320 Then
        Set s = sh.Shapes(Decorator.G_1320)
        Set sr = sh.Shapes.Range(s.name)
        On Error Resume Next
        sh.Shapes.Range("V_" & Decorator.G_1320).Delete
        
        
    ElseIf ESCENARIO = emodRectCorail1355 Then
        Set s = sh.Shapes(Decorator.G_1355)
        Set sr = sh.Shapes.Range(s.name)
        On Error Resume Next
        sh.Shapes.Range("V_" & Decorator.G_1355).Delete
        
    ElseIf ESCENARIO = emodRectCorailNN Then
        Set s = sh.Shapes(Decorator.G_NN)
        Set sr = sh.Shapes.Range(s.name)
        On Error Resume Next
        sh.Shapes.Range("V_" & Decorator.G_NN).Delete
    End If
    
    
    s.top = s.top + 10
End Sub

Private Sub IShape_create()



    ' if it is already created - no need to create it again
    If s Is Nothing Then

        If Not m1 Is Nothing Then
            ' becuase i decided that creation is without params - it is important that in prev step i aquire all imprortant params for what kind of shape i want to create
            ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
            
            ' second step of validation - even if modifier is available it might be that params are not completed due to tactics of partial pre-def in objects of type Modifier
            If m1.useful() Then
                
                ' for now static
                Dim sh As Worksheet
                Set sh = ThisWorkbook.Sheets(Decorator.G_SH_FLOW)
                Set s = sh.Shapes.AddShape(m1.shapeType, m1.left, m1.top, m1.width, m1.height)
                DoEvents
                s.Visible = msoCTrue
                
                With s
                    .name = m1.name
                    Set sr = sh.Shapes.Range(.name)
                End With
                
                With sr.Fill
                    .ForeColor.RGB = m1.shapeBgColor
                    .Transparency = 0
                    .Solid
                End With
                
                Sleep 100
                DoEvents
                
                
                sr.TextFrame2.TextRange.Characters.Text = m1.txt
                sr.TextFrame2.TextRange.Characters().ParagraphFormat.Alignment = msoAlignCenter
                sr.TextFrame2.VerticalAnchor = msoAnchorMiddle
                With sr.TextFrame2.TextRange.Characters.Font
                
                
                    If Decorator.G_DECORATOR Like m1.txt & "*" Then
                        .Size = 24
                    Else
                        .Size = 16
                    End If
                    .Fill.Solid
                    .Fill.Transparency = 0
                    .Fill.Visible = msoCTrue
                    
                End With
                
                DoEvents
                
                ' Debug.Print TypeName(s) & " " & TypeName(s.Parent) -> parent is worksheet ->
                
                'sr.top = 150 ' this test working on same shape
                's.top = 100 ' this test working on same shape
                
            End If
            
            
            ' -------------------------------------------------------------------------------------------------------------------------------------------------------------
        End If
        
    End If

End Sub





Private Sub IShape_establishConnections()


    
    If Not data1.gD() Is Nothing Then
        
        If data1.gD().Count > 0 Then
        
            ' one more required if all in dic are green
            ' but for now let stay like this
            ' GREEN
            With vsr.Fill
                .ForeColor.RGB = RGB(0, 255, 0)
                .Transparency = 0
                .Solid
            End With

            connectWithArrows data1
        End If
    End If
End Sub

Private Function IShape_getModifier() As ShapeModifier
    Set IShape_getModifier = m1
End Function



Private Sub IShape_modify(m As Decorator.ShapeModifier)

End Sub

Private Sub IShape_predefine(Optional e1 As EMOD)

    ESCENARIO = e1
    Set m1 = New Decorator.ShapeModifier
    m1.partialpredef e1
    
    
End Sub

Private Sub IShape_setData(d As IFile)
    Set data1 = d
End Sub

Private Sub IShape_update()

End Sub

Private Sub IShape_validate()

    ' validation shape
    Set vs = Nothing
    Set vsr = Nothing


    ' we have modifier - so we can use same data
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(Decorator.G_SH_FLOW)
    Set vs = sh.Shapes.AddShape(msoShapeCross, m1.left - 10, m1.top - 10, 20, 20)
    
    With vs
        .name = "V_" & m1.name
        Set vsr = sh.Shapes.Range(.name)
    End With
    
    If Not data1 Is Nothing Then
    
        ' if getting range then currently the data is starting (Corail)
        If Not data1.gRng() Is Nothing Then
    
            If data1.gRng.Value = 1 Then
            
                ' GREEN
                With vsr.Fill
                    .ForeColor.RGB = RGB(0, 255, 0)
                    .Transparency = 0
                    .Solid
                End With
            Else
            
                ' RED
                With vsr.Fill
                    .ForeColor.RGB = RGB(255, 0, 0)
                    .Transparency = 0
                    .Solid
                End With
                
                MsgBox "" & m1.txt & " FAILED!", vbCritical
            End If
        Else
            
            
            ' we are getting dictionary with iterated G_ shapes
            If Not data1.gD() Is Nothing Then
                
                If data1.gD().Count > 0 Then
                    ' one more required if all in dic are green
                    ' but for now let stay like this
                    ' GREEN
                    With vsr.Fill
                        .ForeColor.RGB = RGB(0, 255, 0)
                        .Transparency = 0
                        .Solid
                    End With
                    
                    
                    ' idk here - this is specific for decorator stuff -> I have some source data from different env
                    ' and how I want to establish connections! -> this should not to inside of the validation process but alter
                    ' connectWithArrows data1 -> not use here! -> there is new one -> establish connections -> IShape have it as well!
                     
                    
                Else
                     ' RED
                    With vsr.Fill
                        .ForeColor.RGB = RGB(255, 0, 0)
                        .Transparency = 0
                        .Solid
                    End With
                End If
            Else
                MsgBox "" & m1.txt & " FAILED (LACK OF DATA)!", vbCritical
            End If
            
        End If
    Else
    
        ' no data avail. at all -> also red
        With vsr.Fill
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
            .Solid
        End With
        
        MsgBox "" & m1.txt & " FAILED (LACK OF DATA)!", vbCritical
    End If
    '


End Sub

Private Sub IShape_group()

    ''
    '    ActiveSheet.Shapes.Range(Array("corailInput1355")).Select
    '    ActiveSheet.Shapes.Range(Array("corailInput1355", "corailInput1320")). _
    '        Select
    '    Selection.ShapeRange.Group.Select
    '    Selection.ShapeRange.Ungroup.Select
    '    Range("H13").Select
    
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(Decorator.G_SH_FLOW)
    ThisWorkbook.Sheets(Decorator.G_SH_FLOW).Activate
    
    sh.Shapes.Range(Array(vs.name, s.name)).Select
    Selection.ShapeRange.group.Select
    Selection.name = "G_" & m1.name
    
    ' Set gsr = Selection
    ' Set gsr = sh.Shapes.Range("G_" & m1.name)
End Sub


Private Sub IShape_defineOnAction(Optional e1 As EMOD)



    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(Decorator.G_SH_FLOW)
    ThisWorkbook.Sheets(Decorator.G_SH_FLOW).Activate
    
    
    Set gsr = sh.Shapes.Range(Array("G_" & m1.name))
    gsr.Select
    
    If m1.name = Decorator.G_DECORATOR Then
        
        Selection.OnAction = "assign2"
    Else
    
        Selection.OnAction = "assign1"
    End If
    sh.Range("A1").Select
    
    
End Sub


' this one is not connected with required interface - internal process of the instance
Private Sub connectWithArrows(d As IFile)

    ' Debug.Print s.name -> OK name of this shape

    Dim is1 As Variant, sh As ShapeRange
    For Each is1 In d.gD().Keys
    
    
        If "G_" & s.name <> CStr(is1) Then
    
            Debug.Print TypeName(is1) & " " & is1 '  -> string key
            inner__connectWithArrows d, Replace(CStr(is1), "G_", "")
            DoEvents
        End If
    Next
    

    
    
End Sub

Private Sub inner__connectWithArrows(d As IFile, shName As String)
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Sheets(Decorator.G_SH_FLOW)
    Dim arrow1 As Shape, arrow2 As Shape, arrow3 As Shape
    Dim r As ShapeRange, gr As Shape
    
    Set arrow1 = sh.Shapes.AddConnector(msoConnectorStraight, 100, 100, 200, 200)
    Set r = sh.Shapes.Range(arrow1.name)
    r.Line.EndArrowheadStyle = msoArrowheadTriangle
    Set gr = d.gD()("G_" & shName)
    Set gr = gr.GroupItems(shName)
    r.ConnectorFormat.BeginConnect gr, 4
    r.ConnectorFormat.EndConnect s, 2
    r.ShapeStyle = msoLineStylePreset15
    r.name = "ARROW_" & shName
End Sub


