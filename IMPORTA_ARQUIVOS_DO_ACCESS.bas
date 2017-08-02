Attribute VB_Name = "IMPORTA_ARQUIVOS_DO_ACCESS"
Sub IMPORTA_TABELA_HISTORICO()
       
       Dim db As DAO.DATABASE
       Dim rs As DAO.Recordset
       Dim s As Worksheet
       Dim rdest As Range
       Dim dados As Range
       Dim w As Workbook
      
       Set w = ActiveWorkbook
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
       rdest.EntireRow.Find("REGIONAL").EntireColumn.AutoFit
       rdest.EntireRow.Find("COD").EntireColumn.HorizontalAlignment = xlCenter
       
     
End Sub

