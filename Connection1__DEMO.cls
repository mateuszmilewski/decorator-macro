VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Connection1__DEMO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private connName As String
Private connDesc As String
Private connStrType As String
Private connStrProvider As String
Private connStrDataSource As String
Private connLocation As String
Private connExtendedProperties As String
Private connString As String
Private sqlQuery As String

Private queryName As String
Private queryDesc As String
Private queryFormula As String


Private wrk As Workbook
Private qs As Queries
Private qt As QueryTable
Private q As WorkbookQuery
Private cnn As Connections


Private sh As Worksheet
Private lo As ListObject
Private wrkCnn As WorkbookConnection
Private oleDbCnn As OLEDBConnection




Private Sub Class_Initialize()
    connName = ""
    connDesc = ""
    connStrType = ""
    connStrProvider = ""
    connStrDataSource = ""
    connLocation = ""
    connExtendedProperties = ""
    connString = ""
    sqlQuery = ""
    
    
    queryName = "nm3"
    queryDesc = "desc1"
    
    'let
    '    Source = Excel.Workbook(File.Contents("C:\WORKSPACE\Corail in PRS\power-query-test\exports\corail1355_002.xlsx"), null, true),
    '    Sheet1_Sheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    '    #"PH" = Table.PromoteHeaders(Sheet1_Sheet, [PromoteAllScalars=true])
    'in
    '    #"PH"
    
    
    queryFormula = "let " & vbNewLine & _
        "Source = Excel.Workbook(File.Contents(""C:\WORKSPACE\Corail in PRS\power-query-test\exports\corail1355_002.xlsx""), null, true)," & vbNewLine & _
        "Sheet1_Sheet = Source{[Item=""Sheet1"",Kind=""Sheet""]}[Data]," & vbNewLine & _
        "#""PH"" = Table.PromoteHeaders(Sheet1_Sheet, [PromoteAllScalars=true])" & vbNewLine & _
        "in" & vbNewLine & _
        "#""PH"""
        
    Set wrk = ThisWorkbook
    Set cnn = wrk.Connections
    
    
    Set qs = wrk.Queries
    
    ' re-establish connection! ooo!
    For Each q In qs
        ' Debug.Print TypeName(q)
        If q.name = "nm3" Then
            q.Delete
            Exit For
        End If
    Next q
    
    qs.Add name:=queryName, Formula:=queryFormula, Description:=queryDesc
    ' def conn type
    connStrType = "OLEDB"
    connStrProvider = "Microsoft.Mashup.OleDb.1"
    'def datasource
    connStrDataSource = "$Workbook$"
    ' def location of data source
    connLocation = queryName
    ' def extended properties
    connExtendedProperties = """"""
    ' build the connection string
    connString = "" & connStrType & ";Provider=" & connStrProvider & _
        ";Data Source=" & connStrDataSource & ";Location=" & connLocation & _
        ";Extended Properties=" & connExtendedProperties + ";"
        
    ' Debug.Print connString ' OK
    
    connName = "cnm1"
    connDesc = "cndesc1"
    
    ' SQL!
    sqlQuery = "SELECT * FROM [" & queryName & "]"
    
    
    For Each wrkCnn In cnn
        ' Debug.Print TypeName(wrkCnn) ' OK
        If wrkCnn.name Like connName & "*" Then
            wrkCnn.Delete
        End If
    Next wrkCnn
    
    cnn.Add2 _
        name:=connName, Description:=connDesc, _
        ConnectionString:=connString, _
        CommandText:=sqlQuery, lCmdType:=xlCmdSql, _
        CreateModelConnection:=False, ImportRelationships:=False
        
    Set wrkCnn = cnn.Item("cnm1")
    Set oleDbCnn = wrkCnn.OLEDBConnection
    
    
    Set sh = Nothing
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets("test1")
    
    If Not sh Is Nothing Then
        Application.DisplayAlerts = False
        ThisWorkbook.Worksheets("test1").Delete
        Application.DisplayAlerts = True
    End If
    
    
    Set sh = ThisWorkbook.Worksheets.Add
    sh.name = "test1"
    ' helper here
    ' sh.ListObjects.Add(SourceType:=xlSrcExternal,Source:wrkCnn,Destination:=sh.Range("A1"))
    Set lo = sh.ListObjects.Add(xlSrcExternal, wrkCnn, , xlGuess, sh.Cells(1, 1))
    Set qt = lo.QueryTable
    With qt
        .CommandType = oleDbCnn.CommandType
        .CommandText = oleDbCnn.CommandText
        .Refresh False
        .AdjustColumnWidth = False
        ' .ListObject.TableStyle = ""
        .Refresh
    End With
        
    
    
    
End Sub

Private Sub Class_Terminate()
    connName = ""
    connDesc = ""
    connStrType = ""
    connStrProvider = ""
    connStrDataSource = ""
    connLocation = ""
    connExtendedProperties = ""
    connString = ""
    sqlQuery = ""
    
    
    queryName = "nm1"
    queryDesc = "desc1"
    queryFormula = ""
    
    
    Set wrk = Nothing
    Set cnn = Nothing
End Sub


Public Sub import()

End Sub
