Attribute VB_Name = "PIVOT_MAIN"

#If VBA7 Then
      Private Declare PtrSafe Function SetCurrentDirectory _
                                            Lib "kernel32" _
                                            Alias "SetCurrentDirectoryA" ( _
                                            ByVal lpPathName As String) _
                                            As Long
#Else
       Private Declare Function SetCurrentDirectory _
                                            Lib "kernel32" _
                                            Alias "SetCurrentDirectoryA" ( _
                                            ByVal lpPathName As String) _
                                            As Long
#End If






Private w As Workbook


Public Sub CALCULA_PONTUACAO_GERAL()
    
    Dim s As Worksheet
    Dim y As Worksheet
    Dim strRegionais
    Dim mes, ano
    Dim totReg As Worksheet
    Dim totRep As Worksheet
    Dim vr As Range
    Dim rng As Range
    Dim tempPath As String
    Dim wbAux As Workbook
   
    Dim hdr As Range
    Dim ahdr As Range
    Dim dirPath As String
    
    Const pastaRelatoriosNoDropbox = "C:\Dropbox\VEDACIT\RELATORIOS"
    Const dropbox As String = "C:\Dropbox\VEDACIT\"
    
    
    Set addin = ThisWorkbook
    
    msg = "Deseja calcular os pontos da campanha VEDATEAM?"  ' Define message.
        
        Title = "VEDATEAM"   ' Define title.
        ' Display message.
        response = MsgBox(msg, vbYesNo, Title)
        If response = vbNo Then Exit Sub
    
    
    If Dir(dropbox) = vbNullString Then MsgBox ("Esta maquina nao tem acesso aos arquivos de dados no Dropbox. Não é possivel executar a rotina!"): Exit Sub
    inputMes = InputBox("Entre o mês a ser calculado...", "VEDATEAM", Month(Date) - 1)
    inputAno = InputBox("Entre o ano...", "VEDATEAM", Year(Date))
    SetCurrentDirectory dropbox
   
    strPath = Application.GetOpenFilename(, , "Selecione o arquivo com os resultados do mês", "Selecione")
    If strPath = False Then Exit Sub
    'strPath = ARQUIVO_DADOS
    astr = Dir(strPath)
   
    mesExtenso = Split(meses, ",")(inputMes - 1)
    
    '=============================== prepara o diretorio para armazenar os arquivos
    
    dirPath = pastaRelatoriosNoDropbox & "\" & mesExtenso
    If Dir(dirPath, vbDirectory) = vbNullString Then MkDir dirPath Else Delete_Folder dirPath: MkDir dirPath
    
   
    '------------------------------------------------------------------------------
    
    If Right(dirPath, 1) <> "\" Then dirPath = dirPath & "\"
    mes = inputMes
    ano = inputAno
    
    'Set w = ActiveWorkbook
   Set w = Workbooks.Add()
    If (GetAttr(dirPath) And vbDirectory) = vbDirectory Then
             wkbName = dirPath & "Pontuacao_" & mesExtenso & "_" & Right(CStr(ano), 2) & ".xlsx"
             Debug.Print wkbName
        Else
            dskTop = Environ("USERPROFILE") & "\Desktop"
            wkbName = dskTop & "\" & "Pontuacao Referente a " & astr
    End If
    
    'Application.ScreenUpdating = False
    
    PVT_HISTORICO w
    
    IMPORTA_TABELA_CAPILARIDADE w
    
    IMPORTA_TABELA_CLIENTES_ATIVOS w
    
    BUSCA_PARAMETROS w
    
    INSERE_DADOS_REALIZADO CStr(strPath), w
    
    INSERE_PLANILHA_PARTICIPACAO w
    
    
    strRegionais = RECUPERA_REGIONAIS_NO_ARQUIVO_METAS
    
        For i = 0 To UBound(Split(strRegionais, ","))
            aRegional = Split(strRegionais, ",")(i)
                
                pula = False
                
                If InStr(1, UCase(aRegional), "CONSTRUTOR") > 0 Then pula = True
                If InStr(1, UCase(aRegional), "EXPORT") > 0 Then pula = True
                
                If Not pula Then
                        Set s = w.Sheets.Add(w.Sheets(1))
                            s.Name = aRegional
                            PREENCHE_PLANILHA s, mes, ano
                         
                 End If
            

        Next i
     Application.ScreenUpdating = True
    'escreve as planilhas de total
     sheetname = "TOTAIS REGIONAIS"
     If Not Evaluate("ISREF('" & sheetname & "'!A1)") Then
        w.Sheets.Add Before:=w.Sheets(1): ActiveSheet.Name = sheetname
     End If
    
    Set totReg = w.Sheets(sheetname)
       totReg.Activate
       totReg.Cells(1, 1) = "REGIONAL"
       totReg.Cells(1, 2) = "PONTOS"
                          
       Set vr = totReg.Cells(1, 1)
       
             For i = 0 To UBound(Split(strRegionais, ","))
                
                aRegional = Split(strRegionais, ",")(i)
                
                pula = False
                
                If InStr(1, UCase(aRegional), "CONSTRUTOR") > 0 Then pula = True
                If InStr(1, UCase(aRegional), "EXPORT") > 0 Then pula = True
                
                If Not pula Then
                        Set s = w.Sheets(aRegional)
                            's.Activate
                            
                            Set vr = vr.Offset(1)
                            vr = s.Cells(2, 1)
                            dddd = "=" & s.Name & "!" & s.Columns(1).Find("REGIONAL TOTAL:").Offset(0, 2).Address
                            vr.Offset(0, 1).FormulaLocal = "= '" & s.Name & "'!" & s.Columns(1).Find("REGIONAL TOTAL:").Offset(0, 2).Address
                         
                 End If
            

            Next i
            totReg.Columns.AutoFit
            Set hdr = totReg.Range(s.Cells(1, 1).Address & ":" & totReg.Cells(1, totReg.UsedRange.Columns.Count).Address)
            hdr.Font.Bold = True
            hdr.WrapText = True
            hdr.VerticalAlignment = xlCenter
            hdr.HorizontalAlignment = xlCenter
            hdr.RowHeight = 40
                       
           ' s.Columns("A:B").Sort key1:=Range("B2"), order1:=xlDescending, Header:=xlYes
                       
            'escreve o total de representantes
                 sheetname = "TOTAIS REPRESENTANTES"
                 If Not Evaluate("ISREF('" & sheetname & "'!A1)") Then
                    w.Sheets.Add Before:=w.Sheets(1): ActiveSheet.Name = sheetname
                 End If
                Set totRep = w.Sheets(sheetname)
                   totRep.Activate
                   totRep.Cells(1, 1) = "REGIONAL"
                   totRep.Cells(1, 2) = "REPRESENTANTE"
                   totRep.Cells(1, 3) = "PONTOS"
                   
                                      
                   Set vr = totRep.Cells(1, 1)
                   
                         For i = 0 To UBound(Split(strRegionais, ","))
                            
                            aRegional = Split(strRegionais, ",")(i)
                            
                            pula = False
                            
                            If InStr(1, UCase(aRegional), "CONSTRUTOR") > 0 Then pula = True
                            If InStr(1, UCase(aRegional), "EXPORT") > 0 Then pula = True
                            
                            If Not pula Then
                                    Set s = w.Sheets(aRegional)
                                        's.Activate
                                        
                                         Set rng = s.Cells(1, s.Rows(1).Find("Representantes").Column)
                                            Set celIni = rng.Offset(1)
                                            Set celFim = rng.End(xlDown)
                                            
                                         Set rng = s.Range(celIni, celFim)
                                         'rng.Select
                                        
                                        For Each c In rng.Cells
                                        Set vr = vr.Offset(1)
                                        vr = s.Name
                                        vr.Offset(0, 1).FormulaLocal = "= '" & s.Name & "'!" & c.Address
                                        ssss = s.Rows(1).Find("TOTAL REPRESENTANTE").Column
                                        vr.Offset(0, 2).FormulaLocal = "= '" & s.Name & "'!" & s.Cells(c.Row, s.Rows(1).Find("TOTAL REPRESENTANTE").Column).Address
                                        
                                       
                                        Next c
                             End If
                        
            
                        Next i
                        
                        
                        totRep.Columns.AutoFit
                        Set hdr = totRep.Range(s.Cells(1, 1).Address & ":" & totRep.Cells(1, totRep.UsedRange.Columns.Count).Address)
                        hdr.Font.Bold = True
                        hdr.WrapText = True
                        hdr.VerticalAlignment = xlCenter
                        hdr.HorizontalAlignment = xlCenter
                        hdr.RowHeight = 40
                        
   
   ' w.SaveAs FileName:=wkbName
    'w.SaveAs FileName:=Environ("USERPROFILE") & "\Desktop\testeBruno.xlsx"
    
    
    
    '======================  INICIA O RELATORIO DAS REGIONAIS em novo workbook +++++++++++++++++++++++++++++++++++++++++++++++++++
    '+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    Set wbAux = Workbooks.Add()
    If (GetAttr(dirPath) And vbDirectory) = vbDirectory Then wbAuxName = dirPath & "RELATORIO_REGIONAIS_" & mesExtenso & "_" & Right(CStr(ano), 2) & ".xlsx"
    
    Set y = wbAux.Sheets(1)
    y.Name = "REGIONAIS x REPRESENTANTE"
    
    Dim REP As Range
    Dim FAT As Range
    Dim CAP As Range
    Dim REN As Range
    Dim FATR As Range
    Dim CAT As Range
    Dim MIX As Range
    
       y.Activate
                   y.Cells(1, 1) = "REGIONAL": Set REG = y.Cells(1, 1)
                   y.Cells(1, 2) = "REPRESENTANTE": Set REP = y.Cells(1, 2)
                   y.Cells(1, 3) = "FATURAMENTO": Set FAT = y.Cells(1, 3)
                   y.Cells(1, 4) = "CAPILARIDADE": Set CAP = y.Cells(1, 4)
                   y.Cells(1, 5) = "RENTABILIDADE": Set REN = y.Cells(1, 5)
                   y.Cells(1, 6) = "FAT. REP.": Set FATR = y.Cells(1, 6)
                   y.Cells(1, 7) = "CLIENTES ATIVOS": Set CAT = y.Cells(1, 7)
                   y.Cells(1, 8) = "MIX": Set MIX = y.Cells(1, 8)
                   y.Cells(1, 9) = "PONTOS REFERENTES AO FATURAMENTO DOS REPRESENTANTES": Set HF = y.Cells(1, 9)
                   y.Cells(1, 10) = "PONTOS REFERENTES AOS CLIENTES ATIVOS DOS REPRESENTANTES": Set HC = y.Cells(1, 10)
                   y.Cells(1, 11) = "PONTOS REFERENTES AO MIX DOS REPRESENTANTES": Set HM = y.Cells(1, 11)
                   
                   y.Cells(1, 12) = "TOTAL": Set TOT = y.Cells(1, 12)
                   
                                      
                   linha = 1
                   
                         For i = 0 To UBound(Split(strRegionais, ","))
                            
                            aRegional = Split(strRegionais, ",")(i)
                            
                            pula = False
                            
                            If InStr(1, UCase(aRegional), "CONSTRUTOR") > 0 Then pula = True
                            If InStr(1, UCase(aRegional), "EXPORT") > 0 Then pula = True
                            
                            If Not pula Then
                                    Set s = w.Sheets(aRegional)
                                        s.Activate
                                        
                                        linha = linha + 1
                                        y.Activate
                                        y.Cells(linha, REG.Column).Select
                                        y.Cells(linha, REG.Column) = s.Cells(2, s.Rows(1).Find("REGIONAL", lookat:=xlWhole).Column)
                                        
                                        y.Cells(linha, FAT.Column) = s.Cells(2, s.Rows(1).Find("PONTOS FATURAMENTO REGIONAL", lookat:=xlWhole).Column)
                                        y.Cells(linha, CAP.Column) = "0"
                                        y.Cells(linha, CAP.Column) = CInt(s.Cells(2, s.Rows(1).Find("CAPILARIDADE", lookat:=xlWhole).Column))
                                        y.Cells(linha, REN.Column) = 0
                                        y.Cells(linha, REN.Column) = CInt(s.Cells(2, s.Rows(1).Find("RENTABILIDADE", lookat:=xlWhole).Column))
                                        
                                        
                                        
                                        
                                         Set rng = s.Cells(1, s.Rows(1).Find("Representantes").Column)
                                            Set celIni = rng.Offset(1)
                                            Set celFim = rng.End(xlDown)
                                            
                                         Set rng = s.Range(celIni, celFim)
                                         'rng.Select
                                        
                                        linReps = linha
                                        For Each c In rng.Cells
                                        
                                        linReps = linReps + 1
                                        
                                        y.Cells(linReps, REP.Column) = c.Text
                                        
                                        y.Cells(linReps, FATR.Column) = s.Cells(c.Row, Range("N1").Column)
                                        y.Cells(linReps, CAT.Column) = "0"
                                        y.Cells(linReps, CAT.Column) = CInt(s.Cells(c.Row, Range("P1").Column))
                                        y.Cells(linReps, MIX.Column) = 0
                                        y.Cells(linReps, MIX.Column) = CInt(s.Cells(c.Row, Range("R1").Column))
                                        
                                        ' escreve as heranças
                                         
                                        y.Cells(linReps, HF.Column) = 0
                                        y.Cells(linReps, HC.Column) = 0
                                        y.Cells(linReps, HM.Column) = 0
                                        If Not y.Cells(linReps, FATR.Column) = 0 Then
                                            y.Cells(linReps, HF.Column) = CInt(s.Cells(c.Row, Range("O1").Column))
                                            y.Cells(linReps, HC.Column) = CInt(s.Cells(c.Row, Range("Q1").Column))
                                            y.Cells(linReps, HM.Column) = CInt(s.Cells(c.Row, Range("S1").Column))
                                        Else
                                            Range(y.Cells(linReps, HF.Column), y.Cells(linReps, HM.Column)).Font.ColorIndex = 15
                                        End If
                                        
                                        y.Cells(linReps, TOT.Column) = "=" & y.Cells(linReps, FATR.Column).Address & "+" & y.Cells(linReps, CAT.Column).Address & "+" & y.Cells(linReps, MIX.Column).Address
                                        If y.Cells(linReps, FATR.Column) = 0 Then y.Cells(linReps, TOT.Column) = 0
                                        
                                        
                                        Next c
                                        
                                        'totaliza os representantes da regional
                                        
                                        y.Cells(linha, HF.Column).FormulaLocal = "=SOMA(" & y.Cells(linha + 1, HF.Column).Address & ":" & y.Cells(linReps, HF.Column).Address & ")"
                                        y.Cells(linha, HC.Column).FormulaLocal = "=SOMA(" & y.Cells(linha + 1, HC.Column).Address & ":" & y.Cells(linReps, HC.Column).Address & ")"
                                        y.Cells(linha, HM.Column).FormulaLocal = "=SOMA(" & y.Cells(linha + 1, HM.Column).Address & ":" & y.Cells(linReps, HM.Column).Address & ")"
                                        
                                        If y.Cells(linha, FAT.Column) = 0 Then
                                            y.Cells(linha, TOT.Column).FormulaLocal = "=SOMA(" & y.Cells(linha, FAT.Column).Address & ":" & y.Cells(linha, REN.Column).Address & ")"
                                            Else
                                            y.Cells(linha, TOT.Column).FormulaLocal = "=SOMA(" & y.Cells(linha, FAT.Column).Address & ":" & y.Cells(linha, HM.Column).Address & ")"
                                        End If
                                        
                                        linha = linReps
                                        
                                        
                                        
                             End If
                        
            
                        Next i
                        
                        
                        y.Columns.AutoFit
                        Set hdr = y.Range(s.Cells(1, 1).Address & ":" & y.Cells(1, y.UsedRange.Columns.Count).Address)
                        hdr.Font.Bold = True
                        hdr.WrapText = True
                        hdr.VerticalAlignment = xlCenter
                        hdr.HorizontalAlignment = xlCenter
                        hdr.RowHeight = 60
                        
                        'ALINHA LARGURA DA COLUNA
                        CAT.EntireColumn.ColumnWidth = 9
                        MIX.EntireColumn.ColumnWidth = 6
                        HF.EntireColumn.ColumnWidth = 22
                        HC.EntireColumn.ColumnWidth = 22
                        HM.EntireColumn.ColumnWidth = 22

                        
                        
        'esconde as regionais
        primeira = True
        Set vr = Nothing
        For z = 2 To y.UsedRange.Rows.Count
            
            Set rng = y.Cells(z, 1)
            
            If rng.Text <> vbNullString Then
                If Not primeira Then
                    vr.Select
                    vr.EntireRow.Group
                    Set vr = Nothing
                    primeira = True
                End If
            
            Else
                
                If primeira Then
                    Set vr = rng
                    primeira = False
                Else
                    Set vr = Union(vr, rng)
                End If
            
            End If
        
        Next z
        vr.EntireRow.Group
        y.Outline.ShowLevels RowLevels:=1
        y.Cells(1, 1).Select
    
     wbAux.SaveAs fileName:=wbAuxName
     wbAux.Close
    
    '------------------ CRIA OS SLIDES COM OS RESULTADOS PARA OS REPRESENTANTES ------------------------------------
    geraPPTresultados 'dirPath & "\flyers.ppt"
    
    
    w.SaveAs fileName:=wkbName
    
    
End Sub

Sub INSERE_DADOS_REALIZADO(caminhoDoArquivo As String, Optional wb As Workbook)
    
    Dim ORIGEM As Workbook
    Dim sh As Worksheet
            
            If wb Is Nothing Then Set wb = ActiveWorkbook

            Set ORIGEM = Workbooks.Open(caminhoDoArquivo) 'pega o arquivo q contem a planilha a ser copiada
            ORIGEM.Sheets(1).Copy After:= _
                         wb.Sheets(wb.Sheets.Count)
            ActiveSheet.Name = "DADOS" 'nomeia a nova planilha
            If ActiveSheet.AutoFilterMode Then Cells.AutoFilter
            ORIGEM.Close (False)
End Sub


Sub INSERE_PLANILHA_PARTICIPACAO(Optional wb As Workbook)
    
    Dim ORIGEM As Workbook
    Dim sh As Worksheet
            
            
            If wb Is Nothing Then Set wb = ActiveWorkbook
            
            Set ORIGEM = Workbooks.Open(ARQUIVO_METAS, False, True) 'pega o arquivo q contem a planilha a ser copiada
            ORIGEM.Sheets("PARTICIPACAO").Copy After:= _
                         wb.Sheets(wb.Sheets.Count)
            ActiveSheet.Name = "PARTICIPACAO" 'nomeia a nova planilha
            If ActiveSheet.AutoFilterMode Then Cells.AutoFilter
            ORIGEM.Close (False)
End Sub


Sub PREENCHE_PLANILHA(s As Worksheet, inputMes, inputAno)


    Dim ORIGEM As Workbook
    Dim sh As Worksheet
    Dim t As Range
    Dim c As Range
    Dim raux As Range
    Dim rowCell As Range
    Dim rRep As Range
    Dim wkb As Workbook
    Dim nomeRegional As String
    Dim rangeBonus As Range
    Dim bonusRegional As Range
    
    Const pointSize As Integer = 16
    
    Set wkb = s.Parent
    
    Coluna = 1
    
                    s.Cells(1, Coluna) = "REGIONAL": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "META": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "REALIZADO": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "%REG": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "CIDADES NOVAS": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "PONTOS FATURAMENTO REGIONAL": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "CAPILARIDADE": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "RENTABILIDADE": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "REPRESENTANTES": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "META REPRESENTANTE": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "REALIZADO": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "% REP": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "CLIENTES NOVOS": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "PONTOS FATURAMENTO REPRESENTANTE": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "HERANCA PONTOS FATURAMENTO": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "CLIENTES ATIVOS": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "HERANCA CLIENTES ATIVOS": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "MIX": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "HERANCA MIX": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "TOTAL REPRESENTANTE": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "TOTAL REGIONAL": Coluna = Coluna + 1
                    s.Cells(1, Coluna) = "BONUS REPRESENTANTE"
                    
    s.Columns.AutoFit
    
          'abre o arquivo metas e preenche os dados
            Set ORIGEM = Workbooks.Open(ARQUIVO_METAS, UpdateLinks:=0, ReadOnly:=True)
     
                Set sh = ORIGEM.Sheets(s.Name)
                'sh.Select
                
                If InStr(1, UCase(s.Name), "HOME") > 0 Then Set t = sh.Cells.Find("Fator de Conversão Utilizado (Bruto/Merc)") Else Set t = sh.Cells.Find("Incremento para a Campanha 2017")
                        
                        Set regio = t.Offset(3)
                        Set REPRE = t.EntireColumn.Find(what:="Represent", After:=t, lookat:=xlPart, SearchOrder:=xlByRows).Offset(1)
                        colMes = regio.Offset(-1).EntireRow.Find(CDate("1/" & inputMes & "/" & Right(CStr(inputAno), 2))).Column
                        
                        metaRegional = sh.Cells(regio.Row, colMes)
                        
                        'escreve regional
                       
                        s.Activate
                        
                        Select Case s.Name
                        
                        Case Is = "Reg Belo Horizonte": nomeRegional = "REGIONAL BELO HORIZONTE"
                        Case Is = "Reg São Paulo Interior": nomeRegional = "REGIONAL SAO PAULO INTERIOR"
                        Case Is = "Conta Construtoras": nomeRegional = "CONTA CONSTRUTORA"
                        Case Is = "Reg Belém": nomeRegional = "REGIONAL BELEM"
                        Case Is = "Reg Florianópolis": nomeRegional = "REGIONAL FLORIANOPOLIS"
                        Case Is = "Reg Ribeirão Preto": nomeRegional = "REGIONAL RIBEIRAO PRETO"
                        Case Is = "Reg Salvador": nomeRegional = "REGIONAL SALVADOR"
                        Case Is = "Reg São Paulo Capital": nomeRegional = "REGIONAL SAO PAULO CAPITAL"
                        Case Is = "Reg Recife": nomeRegional = "REGIONAL RECIFE"
                        Case Is = "Reg Rio de Janeiro": nomeRegional = "REGIONAL RIO DE JANEIRO"
                        Case Is = "Conta Home Center": nomeRegional = "CONTA HOMECENTER"
                        Case Is = "Reg Curitiba": nomeRegional = "REGIONAL CURITIBA"
                        Case Is = "Reg Fortaleza": nomeRegional = "REGIONAL FORTALEZA"
                        Case Is = "Reg Cuiabá": nomeRegional = "REGIONAL CUIABA"
                        Case Is = "Exportações": nomeRegional = "CONTA EXPORTACAO"
                        Case Is = "0": nomeRegional = "-"
                        Case Is = "Reg Goiânia": nomeRegional = "REGIONAL GOIANIA"
                        Case Is = "Reg Porto Alegre": nomeRegional = "REGIONAL PORTO ALEGRE"

                        End Select
                        lRow = s.Cells(Rows.Count, 1).End(xlUp).Row + 1
                        s.Cells(lRow, 1) = nomeRegional
                        s.Cells(lRow, 2) = metaRegional
                            s.Cells(lRow, 2).NumberFormat = "#,##0"
                            
                       
                        
                        
                            'BUSCA REPRESENTANTES---------------------------------------------
                            '-----------------------------------------------------------------
                                    sh.Outline.ShowLevels RowLevels:=1
                                    
                                    Set t = REPRE.Resize(60000, 1)
                                    't.Select
                                    Dim rng1 As Range
                                    Dim rng2 As Range
                                    
                                    Set rng1 = Nothing
                                    Set rng2 = Nothing
                                        
                                    On Error Resume Next
                                    Set rng1 = t.SpecialCells(xlCellTypeConstants)
                                    Set rng2 = t.SpecialCells(xlCellTypeFormulas)
                                    On Error GoTo 0
                                    
                                   
                                    
                                    If rng2 Is Nothing And Not rng1 Is Nothing Then
                                      Set t = Intersect(t.SpecialCells(xlCellTypeVisible), rng1)
                                    ElseIf rng1 Is Nothing And Not rng2 Is Nothing Then
                                      Set t = Intersect(t.SpecialCells(xlCellTypeVisible), rng2)
                                    ElseIf Not rng1 Is Nothing And Not rng2 Is Nothing Then
                                      Set t = Intersect(t.SpecialCells(xlCellTypeVisible), Union(rng1, rng2))
                                    End If
                                                          
                                   
                                    't.Select
                                            
                                        For Each REPRE In t.Cells
                                    
                                        
                                         'escreve representante
                                         If sh.Cells(REPRE.Row, colMes) = 0 Then GoTo skip_rep
                                          s.Cells(lRow, s.Rows(1).Find("REPRESENTANTES").Column) = REPRE.Text
                                            s.Cells(lRow, s.Rows(1).Find("META REPRESENTANTE").Column) = sh.Cells(REPRE.Row, colMes)
                                               s.Cells(lRow, s.Rows(1).Find("META REPRESENTANTE").Column).NumberFormat = "#,##0"
                                         
                                        'incrementa a linha na planilha metas
                                        lRow = s.Cells(Rows.Count, 9).End(xlUp).Row + 1
skip_rep:
                                        Next REPRE
                                        
                                        
           'merge as celulas da regional
           
            For z = 1 To s.Rows(1).Find("REPRESENTANTES").Column - 1
                                        
                Set rMerge = s.Range(s.Cells(2, z).Address & ":" & s.Cells(lRow - 1, z).Address)
                    rMerge.Select
                    rMerge.Merge
                    rMerge.VerticalAlignment = xlCenter
                    rMerge.Columns.AutoFit
            
            Next z
            
            'escreve o faturamento realizado
            
            s.Cells(2, 3) = GET_FATURAMENTO_REALIZADO(CStr(inputAno), CStr(inputMes), nomeRegional)
                s.Cells(2, 3).NumberFormat = "#,##0"
                
            s.Cells(2, s.Rows(1).Find("%REG").Column) = s.Cells(2, s.Rows(1).Find("REALIZADO").Column) / s.Cells(2, s.Rows(1).Find("META").Column)
                s.Cells(2, s.Rows(1).Find("%REG").Column).NumberFormat = "0.00%"
                'pontos de faturamento da regional
                 With s.Cells(2, s.Rows(1).Find("PONTOS FATURAMENTO REGIONAL").Column)
                    .Value = GET_PONTOS_FATURAMENTO_REGIONAL(s.Cells(2, s.Rows(1).Find("%REG").Column))
                    .Font.Bold = True
                    .Font.Size = pointSize
                    .HorizontalAlignment = xlCenter
                 End With
                 
                 'preenche pontos de Capilaridade
                s.Cells(2, s.Rows(1).Find("CAPILARIDADE").Column) = 0
                If s.Cells(2, s.Rows(1).Find("PONTOS FATURAMENTO REGIONAL").Column) <> 0 Then
                    With s.Cells(2, s.Rows(1).Find("CAPILARIDADE").Column)
                        .Value = calcCapilaridade(nomeRegional, inputMes, inputAno)
                        .Font.Bold = True
                        .Font.Size = pointSize
                        .HorizontalAlignment = xlCenter
                    End With
                End If
                 
                 
                'preenche pontos de rentabilidade
                s.Cells(2, s.Rows(1).Find("RENTABILIDADE").Column) = 0
                If s.Cells(2, s.Rows(1).Find("PONTOS FATURAMENTO REGIONAL").Column) <> 0 Then
                    With s.Cells(2, s.Rows(1).Find("RENTABILIDADE").Column)
                        .Value = 50
                        .Font.Bold = True
                        .Font.Size = pointSize
                        .HorizontalAlignment = xlCenter
                    End With
                End If
                'poe cor na aba da planilha segundo o desempenho do faturamento da regional
                If s.Cells(2, s.Rows(1).Find("PONTOS FATURAMENTO REGIONAL").Column) >= 53 Then s.Tab.Color = vbGreen
                If s.Cells(2, s.Rows(1).Find("PONTOS FATURAMENTO REGIONAL").Column) = 38 Then s.Tab.Color = vbYellow
                If s.Cells(2, s.Rows(1).Find("PONTOS FATURAMENTO REGIONAL").Column) = 45 Then s.Tab.Color = vbYellow
                If s.Cells(2, s.Rows(1).Find("PONTOS FATURAMENTO REGIONAL").Column) = 0 Then s.Tab.Color = vbRed
                    
                
                
            Set rRep = s.Range(s.Cells(2, s.Rows(1).Find("REPRESENTANTES").Column).Address & ":" & s.Cells(lRow - 1, s.Rows(1).Find("REPRESENTANTES").Column).Address)
            regBonus = 0
            
            
            For Each c In rRep.Cells
                
                'calcula os pontos de faturamento
                s.Cells(c.Row, rRep.Offset(0, 2).Column) = GET_FATURAMENTO_REALIZADO(CStr(inputAno), CStr(inputMes), nomeRegional, c.Text)
                    s.Cells(c.Row, rRep.Offset(0, 2).Column).NumberFormat = "#,##0"
                s.Cells(c.Row, rRep.Offset(0, 3).Column) = 0
                If s.Cells(c.Row, rRep.Offset(0, 1).Column) <> 0 Then s.Cells(c.Row, rRep.Offset(0, 3).Column) = s.Cells(c.Row, rRep.Offset(0, 2).Column) / s.Cells(c.Row, rRep.Offset(0, 1).Column)
                    s.Cells(c.Row, rRep.Offset(0, 3).Column).NumberFormat = "0.00%"
                
                'escreve os pontos
                s.Cells(c.Row, s.Rows(1).Find("PONTOS FATURAMENTO REPRESENTANTE").Column) = GET_PONTOS_FATURAMENTO(s.Cells(c.Row, rRep.Offset(0, 3).Column))
                s.Cells(c.Row, s.Rows(1).Find("HERANCA PONTOS FATURAMENTO").Column) = GET_PONTOS_FATURAMENTO_REGIONAL(s.Cells(c.Row, rRep.Offset(0, 3).Column))
                
                'calcula clientes ativos
                    cliA = RECUPERA_CLIENTES_ATIVOS(c.Text, CStr(inputMes), CStr(inputAno))
                    's.Activate
                    s.Cells(c.Row, s.Rows(1).Find("CLIENTES NOVOS").Column) = cliA
                    
                        s.Cells(c.Row, s.Rows(1).Find("CLIENTES ATIVOS").Column) = 0
                        s.Cells(c.Row, s.Rows(1).Find("HERANCA CLIENTES ATIVOS").Column) = 0
                        
                        If cliA = -1 Then GoTo skipClientesAtivos
                        Set raux = w.Sheets("PARAMETROS").Range("1:1").Find("CLIENTES ATIVOS")
                        'raux.Select
                            If cliA >= 10 Then cliA = 10
                            Set rowCell = raux.Offset(cliA + 1, 0)
                            If rowCell Is Nothing Then GoTo skipClientesAtivos
                                
                                s.Cells(c.Row, s.Rows(1).Find("CLIENTES ATIVOS").Column) = rowCell.Offset(0, 3)
                                s.Cells(c.Row, s.Rows(1).Find("HERANCA CLIENTES ATIVOS").Column) = rowCell.Offset(0, 2)
                                s.Activate
                    
                    
skipClientesAtivos:
                ' calcula mix da forma antiga
                'If s.Cells(c.Row, s.Rows(1).Find("PONTOS FATURAMENTO REPRESENTANTE").Column) = 0 Then
                    
                '    s.Cells(c.Row, s.Rows(1).Find("MIX").Column) = 0
                '    s.Cells(c.Row, s.Rows(1).Find("HERANCA MIX").Column) = 0
                
                
                'Else
                '    s.Cells(c.Row, s.Rows(1).Find("MIX").Column).Select
                '    pMix = CALCULA_PONTOS_MIX(c.Text, c.Offset(0, 1), inputAno, inputMes)
                '    s.Activate
                '    s.Cells(c.Row, s.Rows(1).Find("MIX").Column) = Split(CStr(pMix), ",")(0)
                '    s.Cells(c.Row, s.Rows(1).Find("HERANCA MIX").Column) = Split(CStr(pMix), ",")(1)
                'End If
                
                'calcula o mix levando em conta as familias com bonus... se quiser voltar atras é so apagar isso e
                'tirar as aspas da rotina de cima
                
                    s.Cells(c.Row, s.Rows(1).Find("MIX").Column).Select
                    pMix = CALCULA_PONTOS_MIX(c.Text, c.Offset(0, 1), inputAno, inputMes)
                    s.Activate
                    s.Cells(c.Row, s.Rows(1).Find("MIX").Column) = Split(CStr(pMix), ",")(0)
                    s.Cells(c.Row, s.Rows(1).Find("HERANCA MIX").Column) = Split(CStr(pMix), ",")(1)
                        
                    obonus = 0
                    obonus = Split(CStr(pMix), ",")(2)
                    
                    'escreve o bonus de faturamento do representante:
                    Set rangeBonus = Nothing
                    Set rangeBonus = s.Cells(c.Row, Coluna)
                    rangeBonus = obonus
                                'verifica se o bonus ajuda o representante a bater meta...
                                If s.Cells(c.Row, rRep.Offset(0, 1).Column) <> 0 Then caluloFaturamentoComBonus = (s.Cells(c.Row, rRep.Offset(0, 2).Column) + obonus) / s.Cells(c.Row, rRep.Offset(0, 1).Column)
                                    If caluloFaturamentoComBonus >= 0.9 Then
                                    Dim aRng As Range
                                        Set aRng = s.Cells(c.Row, s.Rows(1).Find("PONTOS FATURAMENTO REPRESENTANTE").Column)
                                        rangeBonus.Offset(0, 1) = caluloFaturamentoComBonus
                                        If aRng = 0 Then rangeBonus.Interior.Color = vbGreen
                                        'se ajudou, recalculamos os pontos de faturamento...
                                        newRealizado = (s.Cells(c.Row, rRep.Offset(0, 2).Column) + obonus)
                                         s.Cells(c.Row, s.Rows(1).Find("PONTOS FATURAMENTO REPRESENTANTE").Column) = GET_PONTOS_FATURAMENTO(caluloFaturamentoComBonus)
                                         s.Cells(c.Row, s.Rows(1).Find("HERANCA PONTOS FATURAMENTO").Column) = GET_PONTOS_FATURAMENTO_REGIONAL(caluloFaturamentoComBonus)
                                    End If
                                
                'atualiza o bonus da regional
                regBonus = regBonus + obonus
                
                '----------------- fim da atualização do mix com bonus
'                qqq = t.Cells.Count
'                Set rMix = t.Cells.Find(Trim(c.Text), searchorder:=xlByRows)
'                rMix.Select
                
                
'escreve as formulas de totais

                'TOTAL DO REPRESENTANTE
                If s.Cells(c.Row, s.Rows(1).Find("PONTOS FATURAMENTO REPRESENTANTE").Column) > 0 Then
                    s.Cells(c.Row, s.Rows(1).Find("TOTAL REPRESENTANTE").Column).FormulaLocal = "=" & s.Cells(c.Row, s.Rows(1).Find("PONTOS FATURAMENTO REPRESENTANTE").Column).Address & "+" & s.Cells(c.Row, s.Rows(1).Find("CLIENTES ATIVOS").Column).Address & "+" & s.Cells(c.Row, s.Rows(1).Find("MIX").Column).Address
                Else
                    s.Cells(c.Row, s.Rows(1).Find("TOTAL REPRESENTANTE").Column) = 0
                End If
               
               'TOTAL DAS HERANCAS DA REGIONAL
                If s.Cells(c.Row, s.Rows(1).Find("PONTOS FATURAMENTO REPRESENTANTE").Column) > 0 Then
                    s.Cells(c.Row, s.Rows(1).Find("TOTAL REGIONAL").Column).FormulaLocal = "=" & s.Cells(c.Row, s.Rows(1).Find("PONTOS FATURAMENTO REGIONAL").Column).Address & "+" & s.Cells(c.Row, s.Rows(1).Find("CAPILARIDADE").Column).Address & "+" & s.Cells(c.Row, s.Rows(1).Find("RENTABILIDADE").Column).Address & "+" & s.Cells(c.Row, s.Rows(1).Find("HERANCA PONTOS FATURAMENTO").Column).Address & "+" & s.Cells(c.Row, s.Rows(1).Find("HERANCA CLIENTES ATIVOS").Column).Address & "+" & s.Cells(c.Row, s.Rows(1).Find("HERANCA MIX").Column).Address
                Else
                    s.Cells(c.Row, s.Rows(1).Find("TOTAL REGIONAL").Column) = 0
                End If
                        
                        
                
            Next c
            
            
            ORIGEM.Close (False)
            
            'formata a planilha
            Dim hdr As Range
            s.Activate
            Set hdr = s.Range(s.Cells(1, 1).Address & ":" & s.Cells(1, s.UsedRange.Columns.Count).Address)
            hdr.Font.Bold = True
            hdr.WrapText = True
            hdr.VerticalAlignment = xlCenter
            hdr.HorizontalAlignment = xlCenter
            hdr.RowHeight = 78
            
            For Each c In hdr.Cells
                Select Case c
                    Case "REGIONAL": c.ColumnWidth = 19
                    Case "PONTOS FATURAMENTO REGIONAL", "META REPRESENTANTE", _
                         "PONTOS FATURAMENTO REPRESENTANTE", "HERANCA PONTOS FATURAMENTO", "TOTAL REPRESENTANTE"
                            c.ColumnWidth = 16
                    Case "CAPILARIDADE", "RENTABILIDADE": c.Orientation = 75
                    Case "REPRESENTANTES": c.ColumnWidth = 31
                    Case "META REPRESENTANTE": c.ColumnWidth = 15
                    Case Else: c.ColumnWidth = 10
                End Select
            Next c
            
            
'-------------------------------------------------
'escreve o total do bonus embaixo do realizado

Set bonusRegional = s.Cells(s.UsedRange.Rows.Count + 1, 3)
bonusRegional = regBonus
bonusRegional.NumberFormat = "#,##0"
        
         'verifica se o bonus ajuda a regional a bater meta...
          metaReg = s.Cells(2, 2)
          newRel = s.Cells(2, 3) + regBonus
          
          'escreve o percentual projetado
          
          bonusRegional.Offset(0, 1) = newRel / metaReg
          bonusRegional.Offset(0, 1).NumberFormat = "0.00%"
          aaa = newRel / metaReg
          If newRel / metaReg >= 0.9 Then
                s.Tab.Color = vbBlue
                'altera pontuação do faturamento em função do bonus...
                s.Cells(2, s.Rows(1).Find("PONTOS FATURAMENTO REGIONAL").Column) = GET_PONTOS_FATURAMENTO_REGIONAL(newRel / metaReg)
                'refaz a capilaridade...
                s.Cells(2, s.Rows(1).Find("CAPILARIDADE").Column) = calcCapilaridade(nomeRegional, inputMes, inputAno)
                'reescreve a rentabilidade
                s.Cells(2, s.Rows(1).Find("RENTABILIDADE").Column) = 50
                
          End If
  
'------------------------------------------------

'coloca os totais na planilha
Dim qr As Range

Set qr = s.Cells(s.UsedRange.Rows.Count + 2, 1)

qr.Select

    qr = "PONTOS ORIGEM REGIONAL:"
    qr.Offset(1) = "PONTOS HERANCA REGIONAL:"
    qr.Offset(2) = "REGIONAL TOTAL:"
    pReg = s.Cells(2, s.Rows(1).Find("PONTOS FATURAMENTO REGIONAL").Column)
    
    If pReg > 0 Then
     qr.Offset(0, 2).FormulaLocal = "=" & s.Cells(2, s.Rows(1).Find("PONTOS FATURAMENTO REGIONAL").Column).Address & "+" _
                                                   & s.Cells(2, s.Rows(1).Find("CAPILARIDADE").Column).Address & "+" & _
                                                     s.Cells(2, s.Rows(1).Find("RENTABILIDADE").Column).Address
     qr.Offset(1) = "PONTOS HERANCA REGIONAL"
     Set rngSomaRegional = s.Cells(1, s.Rows(1).Find("TOTAL REGIONAL").Column)
        Set celIni = rngSomaRegional.Offset(1)
        Set celFim = rngSomaRegional.End(xlDown)
           
     
     qr.Offset(1, 2).FormulaLocal = "=SOMA(" & celIni.Address(False, False) & ":" & celFim.Address(False, False) & ")"
        
     qr.Offset(2, 2).FormulaLocal = "=" & qr.Offset(0, 2).Address(False, False) & "+" & qr.Offset(1, 2).Address(False, False)
        
    Else
     qr.Offset(0, 2) = 0
     qr.Offset(1, 2) = 0
     qr.Offset(2, 2) = 0
    End If
    
    

End Sub

Sub testeMix()

CALCULA_PONTOS_MIX "BRITO REPRES LTDA", 292735, 2017, 4


End Sub


Function CALCULA_PONTOS_MIX(nomeRep As String, vlrMeta As Variant, anoPesquisa, mesPesquisa)
        
      Dim w As Workbook
      Dim prt As Worksheet
      Dim colParticipacao As Integer
      Dim p As Range
      Dim famPromo As Range
      Dim aux As Worksheet
      Dim relatorio As Worksheet
      Dim c As Range
      Dim xx As Range
      Dim relRange As Range
      Dim pontuaDeQualquerJeito As Boolean
      
      Set w = ActiveWorkbook
      Set prt = w.Sheets("PARTICIPACAO")
        
        
      resultado = "0,0,0"
      If InStr(1, UCase(nomeRep), "REG.") > 0 Then GoTo finaliza
      'monta a range de familias promocionais
      
      Set famPromo = w.Sheets("PARAMETROS").Rows(1).Find(what:="MIX")
      'famPromo.Select
        
      Set famPromo = famPromo.Offset(1)
      Set famPromo = famPromo.Resize(famPromo.End(xlDown).Row - 1)
      'famPromo.Select
      
      If Not Evaluate("ISREF('" & "RELATORIO_MIX" & "'!A1)") Then
        w.Sheets.Add After:=w.Sheets(w.Sheets.Count): ActiveSheet.Name = "RELATORIO_MIX"
      End If
      
      Set relatorio = w.Sheets("RELATORIO_MIX")
      
redo:
      If Not Evaluate("ISREF('" & "AUXMIX" & "'!A1)") Then
        w.Sheets.Add After:=w.Sheets(w.Sheets.Count): ActiveSheet.Name = "AUXMIX"
      Else
        'Application.DisplayAlerts = False
        '    w.Sheets("AUXMIX").Delete
        'Application.DisplayAlerts = True
        'GoTo redo
      End If
      
      
      Set aux = w.Sheets("AUXMIX")
      aux.Cells.Clear
      
      lin = 1
      aux.Cells(1, lin) = UCase(nomeRep)
      
      For Each c In famPromo.Cells
          
          pontuaDeQualquerJeito = False
          lin = lin + 1
            
          aux.Cells(lin, 1) = c.Text
          
          ' busca a participacao...
          
          'prt.Activate
          
          Set xx = prt.Columns(1).Find(what:=c.Text)
          
          If Not xx Is Nothing Then
                'xx.Select
                  
                'escreve a participacao na auxiliar
                Set p = prt.Rows(1).Find(what:=nomeRep)
                If p Is Nothing Then GoTo finaliza
                'prt.Cells(xx.Row, p.Column).Select
                aux.Cells(lin, 2) = prt.Cells(xx.Row, p.Column)
                aux.Cells(lin, 2).NumberFormat = "0.00%"
                
                'calcula a meta
                
                aux.Cells(lin, 3) = vlrMeta * aux.Cells(lin, 2)
                aux.Cells(lin, 3).NumberFormat = "#,##0"
                
                'recupera o realizado
                      
                aux.Cells(lin, 4) = GET_REALIZADO_POR_FAMILIA(aux.Cells(1, 1), c.Text, CInt(anoPesquisa), CInt(mesPesquisa))
                'if aux.Cells(lin, 4) = -1
                aux.Cells(lin, 4).NumberFormat = "#,##0"
                      
                'calcula pontos representante
                      If aux.Cells(lin, 4) > 0 And aux.Cells(lin, 4) >= aux.Cells(lin, 3) Then
                        
                            aux.Cells(lin, 5) = c.Offset(0, 3)
                            aux.Cells(lin, 6) = c.Offset(0, 2)
                            aux.Cells(lin, 7) = 0
                      Else
                            aux.Cells(lin, 5) = 0
                            aux.Cells(lin, 6) = 0
                            aux.Cells(lin, 7) = 0
                      
                      End If
                      
                If c.Text = "Fita" Then
                            aux.Cells(lin, 5) = c.Offset(0, 3)
                            aux.Cells(lin, 6) = c.Offset(0, 2)
                            aux.Cells(lin, 7) = aux.Cells(lin, 3) - aux.Cells(lin, 4)
                            If aux.Cells(lin, 7) < 0 Then aux.Cells(lin, 7) = 0 'se teve faturamento nesse item.. o bonus é 0
                            
                End If
                
                If c.Text = "Vedalit" Then
                            aux.Cells(lin, 5) = c.Offset(0, 3)
                            aux.Cells(lin, 6) = c.Offset(0, 2)
                            aux.Cells(lin, 7) = aux.Cells(lin, 3) - aux.Cells(lin, 4)
                            If aux.Cells(lin, 7) < 0 Then aux.Cells(lin, 7) = 0
                End If
                
                'If c.Text = "Vedamax" Then
                '            aux.Cells(lin, 5) = c.Offset(0, 3)
                '            aux.Cells(lin, 6) = c.Offset(0, 2)
                '            aux.Cells(lin, 7) = aux.Cells(lin, 3) - aux.Cells(lin, 4)
                '            If aux.Cells(lin, 7) < 0 Then aux.Cells(lin, 7) = 0
                'End If
                
          End If
            
      Next c
      
      Set p = aux.Cells(1, 5)
      Set p = p.EntireColumn
      
      PontosRepresentante = WorksheetFunction.Sum(p)
      
      Set p = aux.Cells(1, 6)
      Set p = p.EntireColumn
      
      PontosRegional = WorksheetFunction.Sum(p)
              
      Set p = aux.Cells(1, 7)
      Set p = p.EntireColumn
      p.NumberFormat = "#,##0"
      
      bonus = Round(WorksheetFunction.Sum(p), 0)
              
      resultado = PontosRepresentante & "," & PontosRegional & "," & bonus
      
      'antes de encerrar, salva o MIX no relatorio para futura analise
      
      aux.Cells(1, 2) = "Participacao"
      aux.Cells(1, 3) = "Meta"
      aux.Cells(1, 4) = "Realizado"
      aux.Cells(1, 5) = "Ptos Repre"
      aux.Cells(1, 6) = "Ptos Regio"
      aux.Cells(1, 7) = "Bonus"
      
      
      If relatorio.UsedRange.Address = "$A$1" And relatorio.Range("A1") = "" Then
            
           Set relRange = relatorio.Range("A1")
            
                aux.UsedRange.Copy relRange
        
            Else
            
           
           Set relRange = relatorio.Cells(relatorio.UsedRange.Rows.Count + 2, 1)
           'relRange.Select
                aux.UsedRange.Copy relRange
          
      End If

finaliza:
    
    CALCULA_PONTOS_MIX = resultado

End Function




Function RECUPERA_CLIENTES_ATIVOS(strRepresentante As String, mes As String, ano As String)
    
    Dim baseCli As Worksheet
    
    Dim strClientesBase As String
    Dim strClientes As String
    Dim strClientesNovos As String
    Dim mediaAnual
    Dim resultado
    Dim rng As Range
    
        
        mesExtenso = Split(meses, ",")(mes - 1)
        
        
        
        'strClientesBase = LISTA_CLIENTES_ATIVOS_NO_ANO(strRepresentante)
        
        mediaAnual = MEDIA_CLIENTES_ATIVOS_NO_ANO(strRepresentante)
        If mediaAnual = -1 Then GoTo finaliza
        
        Set baseCli = ActiveWorkbook.Sheets("CLIENTES_ATIVOS_2016")
        strClientesBase = baseCli.Cells(baseCli.Columns(1).Find(strRepresentante).Row, baseCli.Rows(1).Find(mesExtenso).Column)
        strClientesBase = baseCli.Cells(baseCli.Columns(1).Find(strRepresentante).Row, baseCli.Rows(1).Find(mesExtenso).Column)
        
        strClientes = GET_LISTA_DO_REALIZADO(clientes, CStr(ano), CStr(mes), , strRepresentante)
        nClientes = UBound(Split(strClientes, ","))
        If nClientes = 0 Then nClientes = 1
        If nClientes = -1 Then GoTo finaliza
        
        If nClientes >= mediaAnual Then
            
              strClientesNovos = COMPARA_STRINGS_E_RETORNA_DIFERENCAS(strClientesBase, strClientes)
                      
                    nClientesNovos = UBound(Split(strClientesNovos, ","))
                    If nClientesNovos = 0 Then nClientesNovos = 1
                    resultado = nClientesNovos
       
        Else
            resultado = 0
        End If
    
encerra:
    RECUPERA_CLIENTES_ATIVOS = resultado
Exit Function


finaliza:
    resultado = -1
    GoTo encerra
    
    
End Function

Function LISTA_CLIENTES_ATIVOS_NO_ANO(repName As String) As String
        
            
    Dim s As Worksheet
    Dim pvt As PivotTable
    Dim pvfld As PivotField
    Dim rng As Range
    Dim result As Range
    Dim agente As String
      
    
    Set pvt = PVT_HISTORICO(ActiveWorkbook)
    pvt.ClearTable
    
    pvt.PivotFields("Ano").Orientation = xlPageField
    'If filtroAno <> 0 Then: pvt.PivotFields("Ano").CurrentPage = filtroAno
    
    pvt.PivotFields("Mes").Orientation = xlPageField
    'If Not filtroMes = 0 Then: pvt.PivotFields("Mes").CurrentPage = filtroMes
    
    pvt.PivotFields("Regional").Orientation = xlPageField
    'If Not filtroRegional = vbNullString Then: pvt.PivotFields("Regional").CurrentPage = filtroRegional
    
    On Error GoTo errorHandler
    pvt.PivotFields("Representante").Orientation = xlPageField
    pvt.PivotFields("Representante").CurrentPage = repName

   
    
    Set pvfld = pvt.PivotFields("Cliente")
    
    pvfld.Orientation = xlDataField
    pvfld.Function = xlCount
    'pvfld.NumberFormat = "0"
    
    
    Set rng = pvfld.DataRange
    
    'rng.Select
    resultado = Round(pvfld.DataRange / 24)
    If resultado = 0 Then resultado = 1
    

    astr = vbNullString
    For Each c In rng.Cells
    
    If astr = vbNullString Then astr = c.Text Else astr = astr & "," & c.Text
    
    Next c
    
finaliza:
    LISTA_CLIENTES_ATIVOS_NO_ANO = astr
    
    Set rng = Nothing
    Set pvfld = Nothing
    Set pvt = Nothing

Exit Function
errorHandler:
    
    astr = "erro"
    Err.Clear
    GoTo finaliza


End Function




Function MEDIA_CLIENTES_ATIVOS_NO_ANO(repName As String)
        
            
    Dim s As Worksheet
    Dim pvt As PivotTable
    Dim pvfld As PivotField
    Dim rng As Range
    Dim result As Range
    Dim agente As String
      
    
    Set pvt = PVT_HISTORICO(ActiveWorkbook)
    pvt.ClearTable
    
    pvt.PivotFields("Ano").Orientation = xlPageField
    'If filtroAno <> 0 Then: pvt.PivotFields("Ano").CurrentPage = filtroAno
    
    pvt.PivotFields("Mes").Orientation = xlPageField
    'If Not filtroMes = 0 Then: pvt.PivotFields("Mes").CurrentPage = filtroMes
    
    pvt.PivotFields("Regional").Orientation = xlPageField
    'If Not filtroRegional = vbNullString Then: pvt.PivotFields("Regional").CurrentPage = filtroRegional
    
    On Error GoTo errorHandler
    pvt.PivotFields("Representante").Orientation = xlPageField
    pvt.PivotFields("Representante").CurrentPage = repName

   
    
    Set pvfld = pvt.PivotFields("Cliente")
    
    pvfld.Orientation = xlDataField
    pvfld.Function = xlCount
    'pvfld.NumberFormat = "0"
    
    
    Set rng = pvfld.DataRange
    
    'rng.Select
    resultado = Round(pvfld.DataRange / 24)
    If resultado = 0 Then resultado = 1
    
finaliza:
    MEDIA_CLIENTES_ATIVOS_NO_ANO = resultado
    
    Set rng = Nothing
    Set pvfld = Nothing
    Set pvt = Nothing

Exit Function
errorHandler:
    
    resultado = -1
    Err.Clear
    GoTo finaliza


End Function

Public Function COMPARA_STRINGS_E_RETORNA_DIFERENCAS(referencia As String, ativos As String) As String

   
    Dim arrReferencia, arrAtivos, arrDiferenca
    Dim J As Integer
    Dim k As Integer
    Dim strNovos
    Dim clienteAtivo As String
    
    

    arrReferencia = Split(referencia, ",")
    arrAtivos = Split(ativos, ",")
        
    strNovos = ""
        
        For J = 0 To UBound(arrAtivos)
            
            clienteAtivo = arrAtivos(J)
           
            For k = 0 To UBound(arrReferencia)
            
                If clienteAtivo = arrReferencia(k) Then GoTo novoCliente
                 
            Next k
            
            If strNovos = "" Then strNovos = clienteAtivo & vbCrLf Else strNovos = strNovos & "," & clienteAtivo
            
novoCliente:
        Next J
        
        
        COMPARA_STRINGS_E_RETORNA_DIFERENCAS = strNovos
        
        Erase arrReferencia
        Erase arrAtivos
        
    


End Function


Public Function GET_FATURAMENTO_REALIZADO(Optional filtroAno As Integer, _
                                          Optional filtroMes As Integer, _
                                          Optional filtroRegional As String, _
                                          Optional filtroRepresentante As String) As Long
    
    Dim s As Worksheet
    Dim pvt As PivotTable
    Dim pvfld As PivotField
    Dim rng As Range
    Dim result As Range
    Dim agente As String
      
    
    Set pvt = PVT_REALIZADO(ActiveWorkbook)
    
    Set s = pvt.Parent
    
    pvt.ClearTable
    
    pvt.PivotFields("Ano").Orientation = xlPageField
    If filtroAno <> 0 Then: pvt.PivotFields("Ano").CurrentPage = filtroAno
    
    pvt.PivotFields("Mês").Orientation = xlPageField
    If Not filtroMes = 0 Then: pvt.PivotFields("Mês").CurrentPage = filtroMes
    
    pvt.PivotFields("Regional").Orientation = xlPageField
    If Not filtroRegional = vbNullString Then: pvt.PivotFields("Regional").CurrentPage = filtroRegional
    
    On Error GoTo errorClean
    
    If filtroRepresentante = vbNullString Then
            pvt.PivotFields("Representante").Orientation = xlPageField
      
       Else
            
            pvt.PivotFields("Representante").Orientation = xlPageField
            pvt.PivotFields("Representante").CurrentPage = filtroRepresentante
    
   End If
    
             strCampo = "Faturamento Mercadoria - R$"
             
             If filtroMes = 2 Then strCampo = "Faturamento"

             pvt.AddDataField pvt.PivotFields(strCampo), "Total", xlSum
             
             GET_FATURAMENTO_REALIZADO = pvt.PivotFields("Total").DataRange
    
    
    'Set rng = pvt.PivotFields("Representante").DataRange
    'rng.Select
    
Exit Function
errorClean:
    
    Debug.Print filtroRepresentante & " nao encontrado"
    GET_FATURAMENTO_REALIZADO = 0
    
    
    Set rng = Nothing
    Set pvfld = Nothing
    Set pvt = Nothing
    
End Function

Sub tetetetete()
aaaa = GET_REALIZADO_POR_FAMILIA("TOCANTINS REPR", "Neutrol", 2017, 2)

End Sub

Public Function GET_REALIZADO_POR_FAMILIA(filtroRepresentante As String, filtroFamilia, _
                                          Optional filtroAno As Integer, _
                                          Optional filtroMes As Integer) As Long
    
    Dim s As Worksheet
    Dim pvt As PivotTable
    Dim pvfld As PivotField
    Dim rng As Range
    Dim result As Range
    Dim agente As String
      
    On Error GoTo errorClean
    
    Set pvt = PVT_REALIZADO(ActiveWorkbook)
    'pvt.Parent.Activate
    Set s = pvt.Parent
    aaa = pvt.PageFields.Count
    'If pvt.PageFields.Count = 4 Then
    
       ' For Each pvfld In pvt.PageFields
    
            'If pvfld.Name = "Representante" And pvfld.CurrentPage <> filtroRepresentante Then pvfld.CurrentPage = filtroRepresentante
            'If pvfld.Name = "Familia" And pvfld.CurrentPage <> filtroFamilia Then pvfld.CurrentPage = filtroFamilia
            'If pvfld.Name = "Ano" And pvfld.CurrentPage <> filtroAno Then pvfld.CurrentPage = filtroAno
            'If pvfld.Name = "Mês" And pvfld.CurrentPage <> filtroMes Then pvfld.CurrentPage = filtroMes
       
        'Next pvfld
    
   'Else
    
            pvt.ClearTable
            
            pvt.PivotFields("Ano").Orientation = xlPageField
            If filtroAno <> 0 Then: pvt.PivotFields("Ano").CurrentPage = filtroAno
            
            pvt.PivotFields("Mês").Orientation = xlPageField
            If Not filtroMes = 0 Then: pvt.PivotFields("Mês").CurrentPage = filtroMes
            
            pvt.PivotFields("Representante").Orientation = xlPageField
            
            If Not filtroRepresentante = vbNullString Then: pvt.PivotFields("Representante").CurrentPage = filtroRepresentante
            
            pvt.PivotFields("Familia").Orientation = xlPageField
            If Not filtroFamilia = vbNullString Then: pvt.PivotFields("Familia").CurrentPage = filtroFamilia
            
    'End If
    
            
             strCampo = "Faturamento Mercadoria - R$"
             If filtroMes = 2 Then strCampo = "Faturamento"
             eee = pvt.DataFields.Count
            If Not pvt.DataFields.Count = 1 Then
             pvt.AddDataField pvt.PivotFields(strCampo), "Total", xlSum
            End If
             GET_REALIZADO_POR_FAMILIA = pvt.PivotFields("Total").DataRange
    
    
    'Set rng = pvt.PivotFields("Representante").DataRange
    'rng.Select
    
Exit Function
errorClean:
    
    aaa = Err.Description
    Debug.Print filtroFamilia & " " & filtroRepresentante & " nao encontrado"
    Err.Clear
    GET_REALIZADO_POR_FAMILIA = 0
    
    
    Set rng = Nothing
    Set pvfld = Nothing
    Set pvt = Nothing
    
End Function



Function RECUPERA_REGIONAIS_NO_ARQUIVO_METAS() As String

         Dim ORIGEM As Workbook
         Dim sh As Worksheet
         Dim planilhas As String
            
            
            If IsMissing(wb) Then Set wb = ActiveWorkbook

            Set ORIGEM = Workbooks.Open(ARQUIVO_METAS, UpdateLinks:=0, ReadOnly:=True) 'pega o arquivo q contem a planilha a ser copiada
            
            For Each sh In ORIGEM.Worksheets
                aa = sh.Name
                bb = sh.Tab.ColorIndex
                If sh.Tab.ColorIndex = 15 Then
                
                    planilhas = planilhas & "," & sh.Name
                    
                End If
            
            Next sh
            
            planilhas = Right(planilhas, Len(planilhas) - 1)
            Debug.Print planilhas
            
            RECUPERA_REGIONAIS_NO_ARQUIVO_METAS = planilhas
            
            ORIGEM.Close (False)
   

End Function

Sub IMPORTA_TABELA_CLIENTES_ATIVOS(w As Workbook)
       
       Dim db As DAO.DATABASE
       Dim rs As DAO.Recordset
       Dim s As Worksheet
       Dim rdest As Range
       Dim dados As Range
       
      
       Set s = w.Worksheets.Add(After:=w.Sheets(w.Sheets.Count)) '(After:=w.Sheets(w.Sheets.Count))
        
       s.Name = "CLIENTES_ATIVOS_2016"
       
       Set db = OpenDatabase(ACCESS_DATABASE)
        
       Set rs = db.OpenRecordset("Select * From [CLIENTES_ATIVOS_2016]")
        
       Set rdest = s.Cells(1, 1)
            
       For lThisField = 0 To rs.Fields.Count - 1
         rdest.Offset(0, lThisField) = "'" & UCase(rs.Fields(lThisField).Name)
       Next
                
        With rdest
           .Select
           .EntireRow.WrapText = False
           .EntireRow.ShrinkToFit = False
           .EntireRow.ColumnWidth = 12
           .EntireRow.RowHeight = 55
           .EntireRow.Font.Bold = True
           .EntireRow.HorizontalAlignment = xlCenter
           .EntireRow.VerticalAlignment = xlCenter
           .EntireRow.Font.Size = 12
        End With
             
       Set dados = rdest.Offset(1, 0)
                 
       dados.CopyFromRecordset rs
       rdest.EntireRow.Find("REPRESENTANTE").EntireColumn.AutoFit
       'rdest.EntireRow.Find("COD").EntireColumn.HorizontalAlignment = xlCenter
       
       rs.Close
       Set rs = Nothing
       db.Close
       Set db = Nothing
     
End Sub

Sub IMPORTA_TABELA_CAPILARIDADE(w As Workbook)
       
       Dim db As DAO.DATABASE
       Dim rs As DAO.Recordset
       Dim s As Worksheet
       Dim rdest As Range
       Dim dados As Range
      
       Set s = w.Worksheets.Add(After:=w.Sheets(w.Sheets.Count)) '(After:=w.Sheets(w.Sheets.Count))
        
       s.Name = "CAPILARIDADE_2016"
       
       Set db = OpenDatabase(ACCESS_DATABASE)
        
       Set rs = db.OpenRecordset("Select * From [Capilaridade_2016]")
        
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


Sub IMPORTA_TABELA_MIX()
       
       Dim db As DAO.DATABASE
       Dim rs As DAO.Recordset
       Dim w As Workbook
       Dim s As Worksheet
       Dim rdest As Range
       Dim dados As Range
       
       Set w = ActiveWorkbook
       Set s = w.Worksheets.Add(Before:=w.Sheets(1)) '(After:=w.Sheets(w.Sheets.Count))
        
       s.Name = "MIX_2016"
       
       Set db = OpenDatabase(ACCESS_DATABASE)
        
       Set rs = db.OpenRecordset("Select * From [Mix_2016]")
        
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
       rdest.EntireRow.Find("REPRESENTANTE").EntireColumn.AutoFit
       rdest.EntireRow.Find("COD").EntireColumn.HorizontalAlignment = xlCenter
       
     
End Sub

Public Function GET_LISTA_DO_REALIZADO(pesquisa As listagemDe, _
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
    
    separador = ","
    
    Const itensPesquisa As String = "Regional,Representante,Cidade,Cliente"
      
    toString = True
    strResult = vbNullString
    
    Set pvt = PVT_REALIZADO(ActiveWorkbook)
    pvt.ClearTable
    
    pos = 0
    
    If Not filtroAno = 0 Then pvt.PivotFields("Ano").Orientation = xlPageField: _
                                        pvt.PivotFields("Ano").CurrentPage = filtroAno ': _
                                        'pvt.PivotFields("Ano").Position = pos: pos = pos + 1
                                        
    If Not filtroMes = 0 Then pvt.PivotFields("Mês").Orientation = xlPageField: _
                                        pvt.PivotFields("Mês").CurrentPage = filtroMes ': _
                                        pvt.PivotFields("Mês").Position = pos: pos = pos + 1

    If Not filtroRegional = vbNullString Then pvt.PivotFields("Regional").Orientation = xlPageField: _
                                        pvt.PivotFields("Regional").CurrentPage = filtroRegional ': _
                                        pvt.PivotFields("Regional").Position = pos: pos = pos + 1
    On Error GoTo errorHandler
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
    
    
    
    For Each c In rng.Cells
        
        strResult = strResult & separador & c.Text
    
    Next c
    
    
     
     If toString Then GET_LISTA_DO_REALIZADO = Right(strResult, Len(strResult) - 1)
Exit Function
errorHandler:
    Err.Clear
        GET_LISTA_DO_REALIZADO = vbNullString
    
    
End Function

Sub BUSCA_PARAMETROS(Optional wb As Workbook)
    
    Dim ORIGEM As Workbook
    Dim sh As Worksheet
            
            If IsMissing(wb) Then Set wb = ActiveWorkbook

            Set ORIGEM = Workbooks.Open(ARQUIVO_PARAMETROS) 'pega o arquivo q contem a planilha a ser copiada
            ORIGEM.Sheets(1).Copy After:= _
                         wb.Sheets(wb.Sheets.Count)
            ActiveSheet.Name = "PARAMETROS" 'nomeia a nova planilha
            ORIGEM.Close (False)
End Sub

Public Function GET_PONTOS_FATURAMENTO(percentualDaMetaRealizado) As String

Dim PontosRepresentante As Integer
Dim PontosRegional As Integer


Select Case percentualDaMetaRealizado
    Case Is < 0.8999: PontosRepresentante = 0: PontosRegional = 0
    Case 0.9 To 0.9499: PontosRepresentante = 75: PontosRegional = 38
    Case 0.9 To 0.9499: PontosRepresentante = 75: PontosRegional = 38
    Case 0.95 To 0.9999: PontosRepresentante = 90: PontosRegional = 45
    Case 1 To 1.0999: PontosRepresentante = 105: PontosRegional = 53
    Case 1.1 To 1.1999: PontosRepresentante = 120: PontosRegional = 60
    Case 1.2 To 1.2999: PontosRepresentante = 135: PontosRegional = 68
    Case 1.3 To 1.3999: PontosRepresentante = 150: PontosRegional = 75
    Case 1.4 To 1.4999: PontosRepresentante = 180: PontosRegional = 90
    Case 1.5 To 1.6999: PontosRepresentante = 210: PontosRegional = 105
    Case 1.7 To 1.8999: PontosRepresentante = 240: PontosRegional = 120
    Case 1.9 To 1.9999: PontosRepresentante = 270: PontosRegional = 135
    Case Is > 2: PontosRepresentante = 300: PontosRegional = 150
End Select
    
    
        
    GET_PONTOS_FATURAMENTO = PontosRepresentante
        
        

End Function


Public Function GET_PONTOS_FATURAMENTO_REGIONAL(percentualDaMetaRealizado) As String

Dim PontosRepresentante As Integer
Dim PontosRegional As Integer


Select Case percentualDaMetaRealizado
    Case Is < 0.8999: PontosRepresentante = 0: PontosRegional = 0
    Case 0.9 To 0.9499: PontosRepresentante = 75: PontosRegional = 38
    Case 0.9 To 0.9499: PontosRepresentante = 75: PontosRegional = 38
    Case 0.95 To 0.9999: PontosRepresentante = 90: PontosRegional = 45
    Case 1 To 1.0999: PontosRepresentante = 105: PontosRegional = 53
    Case 1.1 To 1.1999: PontosRepresentante = 120: PontosRegional = 60
    Case 1.2 To 1.2999: PontosRepresentante = 135: PontosRegional = 68
    Case 1.3 To 1.3999: PontosRepresentante = 150: PontosRegional = 75
    Case 1.4 To 1.4999: PontosRepresentante = 180: PontosRegional = 90
    Case 1.5 To 1.6999: PontosRepresentante = 210: PontosRegional = 105
    Case 1.7 To 1.8999: PontosRepresentante = 240: PontosRegional = 120
    Case 1.9 To 1.9999: PontosRepresentante = 270: PontosRegional = 135
    Case Is > 2: PontosRepresentante = 300: PontosRegional = 150
End Select
    
    
        
    GET_PONTOS_FATURAMENTO_REGIONAL = PontosRegional
        
        

End Function



