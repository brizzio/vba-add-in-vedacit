Attribute VB_Name = "PARTICIPACAO"
Sub SUBSTITUI_NOME_REPRESENTANTE_PEO_ID()

      Dim representantes As Worksheet
      Dim participacao As Worksheet
      Dim nomes As Range
      Dim nome As Range
      
      
      
      Set representantes = ThisWorkbook.Sheets("representantes")
      Set participacao = ThisWorkbook.Sheets("participacao")
      
      Set nomes = participacao.Range("C1:EC1")
      
      For Each nome In nomes
            
            nome.Select
            
      Next
        



End Sub
