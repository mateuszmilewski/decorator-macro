Attribute VB_Name = "MainFlowVisModule"
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





Public Sub step001(ictrl As IRibbonControl)
    make1355
    make1320
    makeNN__NOK
End Sub

Public Sub step001_1()
    makeNN
End Sub

Public Sub step002(ictrl As IRibbonControl)
    
    makeDECORATOR
End Sub


Public Sub step003(ictrl As IRibbonControl)
    
    makeExports
End Sub

Public Sub makeExports()
    
    ap
End Sub


Public Sub ribbonMakeDECORATOR(ictrl As IRibbonControl)
    makeDECORATOR
End Sub

Public Sub makeDECORATOR()
    
    ' artificial progress
    ap


    Dim flowSh As Worksheet
    Set flowSh = ThisWorkbook.Sheets(G_SH_FLOW)

    ' Decorator.FlowVisModule.clearDashboard ' leave old data
    
    Dim shs As Decorator.ShapesHandler
    Set shs = New Decorator.ShapesHandler
    
    shs.makeDECORATOR
    
    Set shs = Nothing
End Sub



Public Sub ribbonMakePCV(ictrl As IRibbonControl)
    makePCV
End Sub


Public Sub makePCV()

    ' artificial progress
    ' ap


    Dim flowSh As Worksheet
    Set flowSh = ThisWorkbook.Sheets(G_SH_FLOW)

    ' Decorator.FlowVisModule.clearDashboard ' leave old data
    
    Dim shs As Decorator.ShapesHandler
    Set shs = New Decorator.ShapesHandler
    
    shs.makePCV
    
    Set shs = Nothing
    
    
    MsgBox "PCV data processing done!", vbInformation
End Sub



Public Sub ribbonMakeLOOK(ictrl As IRibbonControl)
    makeLOOK
End Sub


Public Sub makeLOOK()

    ' artificial progress
    ap


    Dim flowSh As Worksheet
    Set flowSh = ThisWorkbook.Sheets(G_SH_FLOW)

    ' Decorator.FlowVisModule.clearDashboard ' leave old data
    
    Dim shs As Decorator.ShapesHandler
    Set shs = New Decorator.ShapesHandler
    
    shs.makeLOOK
    
    Set shs = Nothing
End Sub




Public Sub ribbonMake1355(ictrl As IRibbonControl)
    make1355
End Sub


Public Sub make1355()

    ' artificial progress
    ap


    Dim flowSh As Worksheet
    Set flowSh = ThisWorkbook.Sheets(G_SH_FLOW)

    ' Decorator.FlowVisModule.clearDashboard ' leave old data
    
    Dim shs As Decorator.ShapesHandler
    Set shs = New Decorator.ShapesHandler
    
    shs.make1355
    
    Set shs = Nothing
    
End Sub


Public Sub ribbonMake1320(ictrl As IRibbonControl)
    make1320
End Sub

Public Sub make1320()

    ' artificial progress
    ap


    Dim flowSh As Worksheet
    Set flowSh = ThisWorkbook.Sheets(G_SH_FLOW)

    ' Decorator.FlowVisModule.clearDashboard ' leave old data
    
    Dim shs As Decorator.ShapesHandler
    Set shs = New Decorator.ShapesHandler
    
    shs.make1320
    
    Set shs = Nothing
    
End Sub

Public Sub makeNN__NOK()


    ' artificial progress
    ap


    Dim flowSh As Worksheet
    Set flowSh = ThisWorkbook.Sheets(G_SH_FLOW)

    ' Decorator.FlowVisModule.clearDashboard ' leave old data
    
    Dim shs As Decorator.ShapesHandler
    Set shs = New Decorator.ShapesHandler
    
    shs.makeNN__NOK
    
    Set shs = Nothing
    
End Sub


Public Sub ribbonMakeNN(ictrl As IRibbonControl)
    makeNN
End Sub


Public Sub makeNN()


    ' artificial progress
    ap


    Dim flowSh As Worksheet
    Set flowSh = ThisWorkbook.Sheets(G_SH_FLOW)

    ' Decorator.FlowVisModule.clearDashboard ' leave old data
    
    Dim shs As Decorator.ShapesHandler
    Set shs = New Decorator.ShapesHandler
    
    shs.makeNN
    
    Set shs = Nothing
    
End Sub



