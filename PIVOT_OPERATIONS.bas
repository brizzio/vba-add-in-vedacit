Attribute VB_Name = "PIVOT_OPERATIONS"


Public Sub CRIA_CONEXAO_COM_TABELA_NO_BANCO_DE_DADOS(connectionName As String, localeNomeDoArquivo As String, nomeDaTabela As String)

    'If ActiveWorkbook.Connections(connectionName).Name = vbNullString Then
        
            ActiveWorkbook.Connections.Add2 connectionName, "", Array( _
            "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" & localeNomeDoArquivo & ";Mode=Share Deny W" _
            , _
            "rite;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Eng" _
            , _
            "ine Type=5;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:" _
            , _
            "New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on " _
            , _
            "Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:" _
            , _
            "Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False" _
            ), nomeDaTabela, 3
            
    'End If





End Sub




Public Function lista(pesquisa As listagemDe, _
            Optional filtroAno As Integer, _
            Optional filtroMes As Integer, _
            Optional filtroRegional As String, _
            Optional filtroRepresentante As String) As String

    Dim s As Worksheet
    Dim pvt As PivotTable
    Dim pvfld As PivotField
    Dim rng As Range
    Dim agente As String
    Dim toString As Boolean
    Dim pos As Integer
    Dim w As Workbook
    
    Set w = ActiveWorkbook
    
    separador = ","
    
    Const itensPesquisa As String = "Regional,Representante,Cidade,Cliente"
      
    toString = True
    strResult = vbNullString
    
    Set pvt = PVT_HISTORICO(w)
    pvt.ClearTable
    
    pos = 0
    
    If Not filtroAno = 0 Then pvt.PivotFields("Ano").Orientation = xlPageField: _
                                        pvt.PivotFields("Ano").CurrentPage = filtroAno ': _
                                        'pvt.PivotFields("Ano").Position = pos: pos = pos + 1
                                        
    If Not filtroMes = 0 Then pvt.PivotFields("Mes").Orientation = xlPageField: _
                                        pvt.PivotFields("Mes").CurrentPage = filtroMes ': _
                                        pvt.PivotFields("Mes").Position = pos: pos = pos + 1

    If Not filtroRegional = vbNullString Then pvt.PivotFields("Regional").Orientation = xlPageField: _
                                        pvt.PivotFields("Regional").CurrentPage = filtroRegional ': _
                                        pvt.PivotFields("Regional").Position = pos: pos = pos + 1

    If Not filtroRepresentante = vbNullString Then pvt.PivotFields("Representante").Orientation = xlPageField: _
                                        pvt.PivotFields("Representante").CurrentPage = filtroRepresentante ': _
                                        pvt.PivotFields("Representante").Position = pos: pos = pos + 1

   
    agente = Split(itensPesquisa, ",")(pesquisa)
    
    Set pvfld = pvt.PivotFields(agente)
    
    pvfld.Orientation = xlRowField
    itemsCount = pvfld.PivotItems.Count
    
    itensVisiveis = 0
    strResult = vbNullString
    
    Set rng = pvfld.DataRange
    
    rng.Select
    
    For Each c In rng.Cells
        
        strResult = strResult & separador & c.Text
    
    Next c
    
    
     
     rng.Select
     
     If toString Then lista = Right(strResult, Len(strResult) - 1)
    
End Function

Sub trerreeste()

    aaa = GET_FATURAMENTO_TOTAL(2016, 4, , "J L P")

End Sub

Public Function GET_FATURAMENTO_TOTAL(Optional filtroAno As Integer, _
            Optional filtroMes As Integer, _
            Optional filtroRegional As String, _
            Optional filtroRepresentante As String) As String
    
    Dim s As Worksheet
    Dim pvt As PivotTable
    Dim pvfld As PivotField
    Dim rng As Range
    Dim result As Range
    Dim agente As String
      
    
    Set pvt = PVT_HISTORICO
    pvt.ClearTable
    
    pvt.PivotFields("Ano").Orientation = xlPageField
    If filtroAno <> 0 Then: pvt.PivotFields("Ano").CurrentPage = filtroAno
    
    pvt.PivotFields("Mes").Orientation = xlPageField
    If Not filtroMes = 0 Then: pvt.PivotFields("Mes").CurrentPage = filtroMes
    
    pvt.PivotFields("Regional").Orientation = xlPageField
    If Not filtroRegional = vbNullString Then: pvt.PivotFields("Regional").CurrentPage = filtroRegional
    
    pvt.PivotFields("Representante").Orientation = xlPageField
    If Not filtroRepresentante = vbNullString Then: pvt.PivotFields("Representante").CurrentPage = filtroRepresentante

   
    
    Set pvfld = pvt.PivotFields("Faturamento")
    
    pvfld.Orientation = xlDataField
    pvfld.NumberFormat = "#,##0"
    
    
    Set rng = pvfld.DataRange
    
    rng.Select
    
    GET_FATURAMENTO_TOTAL = pvfld.DataRange
    
    Set rng = Nothing
    Set pvfld = Nothing
    Set pvt = Nothing
    
End Function

Sub trerrrerereeste()

    aaa = GET_SUB_TABELA_REPRESENTANTES_X_REGIONAL("REGIONAL BELO HORIZONTE") ', 2016, 3)

End Sub

Public Function GET_SUB_TABELA_REPRESENTANTES_X_REGIONAL(filtroRegional As String, Optional filtroAno As Integer, _
            Optional filtroMes As Integer) As Range
    
    Dim s As Worksheet
    Dim pvt As PivotTable
    Dim pvfld As PivotField
    Dim pivotRepresentantes As PivotField
    Dim rng As Range
    Dim result As Range
    Dim agente As String
      
    
    Set pvt = PVT_HISTORICO
    pvt.ClearTable
    
    pvt.PivotFields("Ano").Orientation = xlPageField
    If filtroAno <> 0 Then: pvt.PivotFields("Ano").CurrentPage = filtroAno
    
    pvt.PivotFields("Mes").Orientation = xlPageField
    If Not filtroMes = 0 Then: pvt.PivotFields("Mes").CurrentPage = filtroMes
    
    pvt.PivotFields("Regional").Orientation = xlPageField
    If Not filtroRegional = vbNullString Then: pvt.PivotFields("Regional").CurrentPage = filtroRegional
    
   
    Set pivotRepresentantes = pvt.PivotFields("Representante")
    pivotRepresentantes.Orientation = xlRowField
    
                 For i = 1 To pivotRepresentantes.PivotItems.Count
                     aaaa = pivotRepresentantes.PivotItems(i).Name
                     If Trim(pivotRepresentantes.PivotItems(i).Name) = "-" Or Trim(pivotRepresentantes.PivotItems(i).Name) = vbNullString Then
                             pivotRepresentantes.PivotItems(i).Visible = False
                     End If
                
                 Next i
              
    Set pvfld = pvt.PivotFields("Faturamento")
    
    pvfld.Orientation = xlDataField
    pvfld.NumberFormat = "#,##0"
    
    
    Set rng = Union(pivotRepresentantes.DataRange, pvfld.DataRange)
    
    rng.Select
    
    GET_SUB_TABELA_REPRESENTANTES_X_REGIONAL = Union(pivotRepresentantes.DataRange, pvfld.DataRange)
    
    Set rng = Nothing
    Set pvfld = Nothing
    Set pvt = Nothing
    
End Function


Sub IMPORTA_DADOS_HISTORICO(w As Workbook)
       
       Dim db As DAO.DATABASE
       Dim rs As DAO.Recordset
       Dim s As Worksheet
       Dim rdest As Range
       Dim dados As Range
      
       Set s = w.Worksheets.Add(After:=w.Sheets(w.Sheets.Count)) '(After:=w.Sheets(w.Sheets.Count))
        
       s.Name = "HISTORICO"
       
       Set db = OpenDatabase(ACCESS_DATABASE)
        
       Set rs = db.OpenRecordset("Select * From [HISTORICO]")
        
       Set rdest = s.Cells(1, 1)
            
       For lThisField = 0 To rs.Fields.Count - 1
         rdest.Offset(0, lThisField) = "'" & UCase(rs.Fields(lThisField).Name)
       Next
                
        With rdest
           .Select
           .EntireRow.WrapText = False
           .EntireRow.ShrinkToFit = False
           .EntireRow.ColumnWidth = 14
           .EntireRow.RowHeight = 55
           .EntireRow.Font.Bold = True
           .EntireRow.HorizontalAlignment = xlCenter
           .EntireRow.VerticalAlignment = xlCenter
           .EntireRow.Font.Size = 12
        End With
             
       Set dados = rdest.Offset(1, 0)
                 
       dados.CopyFromRecordset rs
       rdest.EntireRow.Find("Regional").EntireColumn.AutoFit
       'rdest.EntireRow.Find("COD").EntireColumn.HorizontalAlignment = xlCenter
       
     
End Sub




Function PVT_HISTORICO(Optional w As Workbook) As PivotTable
Attribute PVT_HISTORICO.VB_ProcData.VB_Invoke_Func = " \n14"
'
' retorna uma planilha com uma pivot table contendo os dados da tabela - HISTORICO - do banco de dados no access
'  para esta função, assumimos que a tabela se encontre dentro do banco de dados dbVedacit.mdb no Dropbox
'

'
Const nomeDaTabelaNoBancoDeDadosParaBasearApivotTable = "HISTORICO"


Dim aux As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim pvf As PivotField

Dim rngSource As Range
Dim sName As String

On Error GoTo errorHandler

If w Is Nothing Then Set w = ActiveWorkbook
    
    
    sName = "PVT_" & nomeDaTabelaNoBancoDeDadosParaBasearApivotTable
    
    If Not Evaluate("ISREF('" & sName & "'!A1)") Then
             
             'busca os dados do historico e salva no arquivo
             IMPORTA_DADOS_HISTORICO w
             Set rngSource = w.Sheets("HISTORICO").UsedRange
             
             'insere a planilha da pivot table
             w.Sheets.Add After:=w.Sheets(w.Sheets.Count)
             
             Set aux = ActiveSheet
             
             aux.Name = sName
             localPlanilha = aux.Name & "!" & aux.Cells(3, 1).Address(ReferenceStyle:=xlR1C1)
             
             Application.ScreenUpdating = True
             Set pvtCache = ActiveWorkbook.PivotCaches.Create(xlDatabase, rngSource, xlPivotTableVersion14)
             
             'Create Pivot table from Pivot Cache
             Set pvt = pvtCache.CreatePivotTable( _
             TableDestination:=localPlanilha, _
             TableName:=nomeDaTabelaNoBancoDeDadosParaBasearApivotTable)
             
             aux.Select
             
             'w.ShowPivotTableFieldList = False
             
            
             With pvt.PivotFields("Ano")
                 .Orientation = xlPageField
                 .Position = 1
             End With
             
             With pvt.PivotFields("Regional")
                 .Orientation = xlPageField
                 .Position = 2
             End With
             With pvt.PivotFields("Representante")
                 .Orientation = xlRowField
                 .Position = 1
             End With
             'pvt.AddDataField pvt.PivotFields("Faturamento"), "FATURAMENTO", xlSum
                                    
  
    End If
    

finaliza:
    Set aux = w.Sheets(sName)
    Set PVT_HISTORICO = aux.PivotTables(nomeDaTabelaNoBancoDeDadosParaBasearApivotTable)

Exit Function


errorHandler:
    codErr = Err.Number
    descErr = Err.Description
        
    
    Err.Clear
    
    Resume 0
    GoTo finaliza
End Function




Sub IMPORTA_DADOS_REALIZADO_2017(w As Workbook)
       
       Dim db As DAO.DATABASE
       Dim rs As DAO.Recordset
       Dim s As Worksheet
       Dim rdest As Range
       Dim dados As Range
      
       Set s = w.Worksheets.Add(After:=w.Sheets(w.Sheets.Count)) '(After:=w.Sheets(w.Sheets.Count))
        
       s.Name = "REALIZADO_2017"
       
       Set db = OpenDatabase(ACCESS_DATABASE)
        
       Set rs = db.OpenRecordset("Select * From [REALIZADO_2017]")
        
       Set rdest = s.Cells(1, 1)
            
       For lThisField = 0 To rs.Fields.Count - 1
         rdest.Offset(0, lThisField) = "'" & UCase(rs.Fields(lThisField).Name)
       Next
                
        With rdest
           .Select
           .EntireRow.WrapText = False
           .EntireRow.ShrinkToFit = False
           .EntireRow.ColumnWidth = 14
           .EntireRow.RowHeight = 55
           .EntireRow.Font.Bold = True
           .EntireRow.HorizontalAlignment = xlCenter
           .EntireRow.VerticalAlignment = xlCenter
           .EntireRow.Font.Size = 12
        End With
             
       Set dados = rdest.Offset(1, 0)
                 
       dados.CopyFromRecordset rs
       rdest.EntireRow.Find("Regional").EntireColumn.AutoFit
       'rdest.EntireRow.Find("COD").EntireColumn.HorizontalAlignment = xlCenter
       
     
End Sub

Function PVT_REALIZADO_2017(Optional w As Workbook) As PivotTable
'
' retorna uma planilha com uma pivot table contendo os dados da tabela - HISTORICO - do banco de dados no access
'  para esta função, assumimos que a tabela se encontre dentro do banco de dados dbVedacit.mdb no Dropbox
'

'
Const nomeDaTabelaNoBancoDeDadosParaBasearApivotTable = "REALIZADO_2017"


Dim aux As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim pvf As PivotField

Dim rngSource As Range
Dim sName As String

On Error GoTo errorHandler

If w Is Nothing Then Set w = ActiveWorkbook
    
    
    sName = "PVT_" & nomeDaTabelaNoBancoDeDadosParaBasearApivotTable
    
    If Not Evaluate("ISREF('" & sName & "'!A1)") Then
             
             'busca os dados do historico e salva no arquivo
             IMPORTA_DADOS_REALIZADO_2017 w
             Set rngSource = w.Sheets("REALIZADO_2017").UsedRange
             
             'insere a planilha da pivot table
             w.Sheets.Add After:=w.Sheets(w.Sheets.Count)
             
             Set aux = ActiveSheet
             
             aux.Name = sName
             localPlanilha = aux.Name & "!" & aux.Cells(3, 1).Address(ReferenceStyle:=xlR1C1)
             
             Application.ScreenUpdating = True
             Set pvtCache = ActiveWorkbook.PivotCaches.Create(xlDatabase, rngSource, xlPivotTableVersion14)
             
             'Create Pivot table from Pivot Cache
             Set pvt = pvtCache.CreatePivotTable( _
             TableDestination:=localPlanilha, _
             TableName:=nomeDaTabelaNoBancoDeDadosParaBasearApivotTable)
             
             aux.Select
             
             
             With pvt.PivotFields("Regional")
                 .Orientation = xlPageField
                 
             End With
                                    
  
    End If
    

finaliza:
    Set aux = w.Sheets(sName)
    Set PVT_REALIZADO_2017 = aux.PivotTables(nomeDaTabelaNoBancoDeDadosParaBasearApivotTable)

Exit Function


errorHandler:
    codErr = Err.Number
    descErr = Err.Description
        
    
    Err.Clear
    
    'Resume 0
    GoTo finaliza
End Function


