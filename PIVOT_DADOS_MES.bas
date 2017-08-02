Attribute VB_Name = "PIVOT_DADOS_MES"


Public Function PVT_REALIZADO(w As Workbook) As PivotTable
'
' retorna uma planilha com uma pivot table contendo os dados da tabela referente a produção no periodo
'  para esta função, assumimos que a tabela ja tenha sido importada para o active workbook atraves da macro
' IMPORTA_DADOS e a planilha tenha por nome ---DADOS

Const nomeDaPlanilha = "DADOS"


Dim aux As Worksheet
Dim pvtCache As PivotCache
Dim pvt As PivotTable
Dim pvf As PivotField
Dim rngSource As Range

Dim sName As String

On Error GoTo errorHandler
 
   
    sName = "PVT_" & nomeDaPlanilha
    
    If Not Evaluate("ISREF('" & sName & "'!A1)") Then
    
             
             w.Sheets.Add After:=w.Sheets(w.Sheets.Count)
             
             Set aux = ActiveSheet
             
             aux.Name = sName
             localPlanilha = aux.Name & "!" & aux.Cells(3, 1).Address(ReferenceStyle:=xlR1C1)
             Set rngSource = w.Sheets(nomeDaPlanilha).UsedRange
             Set pvtCache = ActiveWorkbook.PivotCaches.Create(xlDatabase, rngSource, xlPivotTableVersion14)
             
             'Create Pivot table from Pivot Cache
             Set pvt = pvtCache.CreatePivotTable( _
             TableDestination:=localPlanilha, _
             TableName:=nomeDaPlanilha)
             
             aux.Select
             
             'w.ShowPivotTableFieldList = False
             
            
            
             
             With pvt.PivotFields("Regional")
                 .Orientation = xlPageField
                 '.Position = 2
             End With
             With pvt.PivotFields("Representante")
                 .Orientation = xlRowField
                 .Position = 1
             End With
             strCampo = rngSource.Rows(1).Find("Faturamento", lookat:=xlPart).Text
             pvt.AddDataField pvt.PivotFields(strCampo), , xlSum
                                    
  
    End If
    

finaliza:
    Set aux = w.Sheets(sName)
    Set PVT_REALIZADO = aux.PivotTables(nomeDaPlanilha)

Exit Function


errorHandler:
    codErr = Err.Number
    descErr = Err.Description
       
    Err.Clear
    
    Resume Next
    GoTo finaliza
End Function


