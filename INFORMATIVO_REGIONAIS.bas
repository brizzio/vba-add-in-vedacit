Attribute VB_Name = "INFORMATIVO_REGIONAIS"
Public Type sSet
    numSet As Variant
    mesesQueCompoemSet As Variant
    mesesQueJaSePassaram As Variant
End Type



Public Sub INFORMA_REGIONAIS_CAPILARIDADE()

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
    Dim lin As Integer
    
    Dim objMsg As Outlook.MailItem
    Dim oApp As Outlook.Application
    Dim ns As Outlook.Namespace
    Dim fld As Outlook.Folder
    Dim oAccount As Outlook.Account
    
    Dim hdr As Range

    Dim ahdr As Range
    Dim clientes As Variant
    Dim agentes As Range
    Dim refData As String
    Dim cadastro As Worksheet
    
    Const dropbox As String = "C:\Dropbox\VEDACIT\"
    'Const conta_vedateam As String = "vedateam@outlook.com"
    
    Set addin = ThisWorkbook
    
    'response = MsgBox("Deseja preparar o informativo para os regionais sobre CAPILARIDADE?", vbYesNo, "VEDATEAM")
    'If response = vbNo Then Exit Sub

    If Dir(dropbox) = vbNullString Then MsgBox ("Esta maquina nao tem acesso aos arquivos de dados no Dropbox. Não é possivel executar a rotina!"): Exit Sub
    
    'inputMes = InputBox("Entre o mês para o informativo", "VEDATEAM", Month(Date))
    'inputAno = InputBox("Entre o ano...", "VEDATEAM", Year(Date))
    
    inputMes = Month(Date)
    inputAno = 2017
    
    mesExtenso = Split(meses, ",")(inputMes - 1)
               
    mes = inputMes
    ano = inputAno
    
    If ano < 2017 Then Exit Sub
    If Month(Date) < mes Then Exit Sub
    
    Set w = ActiveWorkbook
    Set s = w.Worksheets.Add(Before:=w.Sheets(1))
    titulos = Split("REGIONAL,NOME,CIDADES NO SET,CIDADES ATIVADAS,PARA,COPIA", ",")
    Set hdr = s.Cells(1, 1).Resize(1, UBound(titulos) + 1)
    hdr = titulos
    
    hdr.RowHeight = 40
    hdr.WrapText = True
    hdr.HorizontalAlignment = xlCenter
    hdr.VerticalAlignment = xlCenter
    hdr.Columns.ColumnWidth = 20
    hdr(1).ColumnWidth = 25
    hdr(2).ColumnWidth = 35
    hdr(3).ColumnWidth = 30
   
    
    'IMPORTA QUADRO DE METAS
    If Not Evaluate("ISREF('" & "METAS" & "'!A1)") Then INSERE_QUADRO_GERAL_DE_METAS w
    
    'IMPORTA CADASTRO DE REPRESENTANTES
     sheetname = "CADREG"
        If Not Evaluate("ISREF('" & sheetname & "'!A1)") Then
               Set ORIGEM = Workbooks.Open("C:\Dropbox\VEDACIT_DADOS\Contatos.xlsx") 'pega o arquivo q contem a planilha a ser copiada
               ORIGEM.Sheets("Regionais").Copy After:= _
                            w.Sheets(w.Sheets.Count)
               ORIGEM.Close (False)
               ActiveSheet.Name = sheetname
               
        End If
        
    Set cadastro = w.Sheets("CADREG")
    'Prepara o email
      Set oApp = New Outlook.Application
      Set ns = oApp.GetNamespace("MAPI")
    Set fld = ns.GetDefaultFolder(olFolderOutbox)
    
                          'seleciona a conta para envio
                          For Each obj In oApp.Session.Accounts
                             
                             If obj = conta_vedateam Then Set oAccount = obj
                             
                          Next obj
    
    Set metas = w.Sheets("METAS")
    
    'Set agentes = metas.Range(metas.Cells(2, 1).Address & ":" & metas.Cells(metas.UsedRange.Rows.Count, 1).Address)
    cadastro.Activate
     Set agentes = cadastro.Range(cadastro.Cells(2, 1).Address & ":" & cadastro.Cells(cadastro.Cells(Rows.Count, 1).End(xlUp).Row, 1).Address)
   agentes.Select
   
   
     'dados comuns para todos os emails
                
                mesRelatorio = mesExtenso
                
                    strMesesNoSet = vbNullString
                    strMesesCorridos = vbNullString
                    
                    Select Case mes
                        
                        Case 1
                            Exit Sub
                        Case 2
                            numeroDoSet = 1
                            mesesNoSet = Array(2, 3, 4)
                            mesesCorridos = Array(2)
                            strMesesNoSet = "FEVEREIRO, MARÇO e ABRIL"
                            strMesesCorridos = "FEVEREIRO"
                        Case 3
                            numeroDoSet = 1
                            mesesNoSet = Array(2, 3, 4)
                            mesesCorridos = Array(2, 3)
                            strMesesNoSet = "FEVEREIRO, MARÇO e ABRIL"
                            strMesesCorridos = "FEVEREIRO e MARÇO"
                        
                        Case 4
                            numeroDoSet = 1
                            mesesNoSet = Array(2, 3, 4)
                            mesesCorridos = Array(2, 3, 4)
                            strMesesNoSet = "FEVEREIRO, MARÇO e ABRIL"
                            strMesesCorridos = "SET COMPLETO"
                        
                        Case 5
                            numeroDoSet = 2
                            mesesNoSet = Array(5, 6, 7)
                            mesesCorridos = Array(5)
                            strMesesNoSet = "MAIO, JUNHO e JULHO"
                            
                        Case 6
                            numeroDoSet = 2
                            mesesNoSet = Array(5, 6, 7)
                            mesesCorridos = Array(5, 6)
                            strMesesNoSet = "MAIO, JUNHO e JULHO"
                            
                        Case 7
                            numeroDoSet = 2
                            mesesNoSet = Array(5, 6, 7)
                            mesesCorridos = Array(5, 6, 7)
                            strMesesNoSet = "MAIO, JUNHO e JULHO"
                        
                        Case 8
                            numeroDoSet = 3
                            mesesNoSet = Array(8, 9, 10)
                            mesesCorridos = Array(8)
                            strMesesNoSet = "AGOSTO, SETEMBRO e OUTUBRO"
                        
                        Case 9
                            numeroDoSet = 3
                            mesesNoSet = Array(8, 9, 10)
                            mesesCorridos = Array(8, 9)
                            strMesesNoSet = "AGOSTO, SETEMBRO e OUTUBRO"
                        
                        
                        Case 10
                            numeroDoSet = 3
                            mesesNoSet = Array(8, 9, 10)
                            mesesCorridos = Array(8, 9, 10)
                        
                        Case 11
                            numeroDoSet = 4
                            mesesNoSet = Array(11, 12)
                            mesesCorridos = Array(11)
                        
                        
                        Case 12
                            numeroDoSet = 4
                            mesesNoSet = Array(11, 12)
                            mesesCorridos = Array(11, 12)
                   
                    End Select
     
                
                   
     'dados especificos para cada email
            
            For Each c In agentes.Cells
            
                 If c.Interior.ColorIndex = 15 Then aRegional = c.Text: GoTo skipThat
                 
                 
                 nomeReg = c.Text
                 cidadesHist = LISTA_CIDADES_NO_HISTORICO_DO_SET(c.Text, mesesNoSet, ano)
                 cidadesReal = LISTA_CIDADES_ATIVAS_NO_SET(c.Text, mesesCorridos, ano)
                 If Not Len(Join(cidadesHist)) > 0 Then c.Interior.ColorIndex = 3: GoTo skipThat
                 
                 cidNoSet = UBound(cidadesHist) + 1
                 If Not Len(Join(cidadesReal)) > 0 Then cidReal = 0 Else cidReal = UBound(cidadesReal) + 1
                 
                 cidNovas = COMPARA_STRINGS_E_RETORNA_DIFERENCAS(Join(cidadesHist, ","), Join(cidadesReal, ","))
                
                
                 endereco = cadastro.Cells.Find(c.Text).Offset(0, 2)
                 enderecoCopia = cadastro.Cells.Find(c.Text).Offset(0, 5)
                 nomePessoal = cadastro.Cells.Find(c.Text).Offset(0, 1)
                       'seleciona a conta para envio
                       
                       
                            Set objMsg = fld.Items.Add(olMailItem)
                            
                             With objMsg
                              .To = endereco
                              .CC = enderecoCopia
                              '.BCC = "bruno.pacheco@vedacit.com.br"
                              .Subject = "INFORMATIVO CAMPANHA VEDATEAM - CAPILARIDADE - acumulado do " & numeroDoSet & " SET de " & ano
                              .Categories = "Vedateam_Informativo_Capilaridade"
                              '.VotingOptions = "Yes;No;Maybe;"
                              .BodyFormat = olFormatHTML
                              .HTMLBody = HTML_BODY_CAPILARIDADE(c.Text, _
                                                                 CStr(nomePessoal), _
                                                                 CStr(numeroDoSet), _
                                                                 CStr(mesExtenso), _
                                                                 CStr(ano), _
                                                                 mesesNoSet, _
                                                                 mesesCorridos, _
                                                                 cidadesHist, _
                                                                 cidadesReal)
                            
                              '.Display
                              .Save
                              '.SendUsingAccount = oAccount
                              
                              '.DeleteAfterSubmit = True
                            End With
                            
                              Set objMsg = Nothing
                     
skipThat:
            '"REGIONAL,NOME,CIDADES NO SET,CIDADES ATIVADAS,PARA,COPIA"
                lin = s.UsedRange.Rows.Count + 1
                's.Cells(lin, 1).Select
                s.Cells(lin, 1) = UCase(c.Text)
                s.Cells(lin, 2) = CStr(nomePessoal)
                s.Cells(lin, 3) = cidNoSet
                s.Cells(lin, 4) = cidReal
                s.Cells(lin, 5) = endereco
                s.Cells(lin, 6) = enderecoCopia
            
            'zera as variaveis
                            
                            cidNoSet = 0
                            cidReal = 0
                            endereco = vbNullString
                            enderecoCopia = vbNullString
                            nomePessoal = vbNullString
            Next c
            
        'remove linhas com email em branco e
        'envia um email para o gerente de produto com o resumo da planilha de envio
        s.Name = "RELATORIO_ENVIO"
        Set rng = hdr(6).Offset(1).Resize(s.UsedRange.Rows.Count - 1, 1)
        rng.Select
        
        For Each c In rng.Cells
            
            If c.Text = vbNullString Then If vr Is Nothing Then Set vr = c Else Set vr = Union(vr, c)
        
        Next c
        vr.Select
        vr.EntireRow.Delete xlUp
        hdr(1).Select
        
        '===========================
        htmlHead = "<br><br>DATA DE REFERÊNCIA: " & mesExtenso & " / " & CStr(ano)
        htmlHead = htmlHead & "<br><br>CAPILARIDADE NO SET: " & numeroDoSet & "<br><br>"
        
        

        Set objMsg = fld.Items.Add(olMailItem)
                            
                             With objMsg
                              .To = "bruno.pacheco@vedacit.com.br"
                              .CC = "vedateam@outlook.com"
                              
                              .Subject = "RELATORIO DE ENVIO EMAIL PARA REGIONAIS - CAPILARIDADE - " & mesExtenso & "/" & ano
                              .Categories = "Relatorio_Email_Marketing_Capilaridade"
                              '.VotingOptions = "Yes;No;Maybe;"
                              .BodyFormat = olFormatHTML
                              .HTMLBody = UCase(htmlHead)
                              .HTMLBody = .HTMLBody & RangetoHTML(s.UsedRange)
                              .HTMLBody = .HTMLBody & "<br><br>======================== FAC SIMILE DO EMAIL ENVIADO =====================================<br><br>"
                              .HTMLBody = .HTMLBody & HTML_BODY_CAPILARIDADE(" NOME REGIONAL ", _
                                                                             " ****nome***** ", _
                                                                             " ***SET*** ", _
                                                                             " ***mes*** ", _
                                                                             " ***ano*** ", _
                                                                             " ***meses do set *** ", _
                                                                             " ***meses corridos*** ", _
                                                                             " ***cidades no historico *** ", _
                                                                             " ***cidades realizadas*** ")
                            
                              '.Display
                              '.SendUsingAccount = oAccount
                              .Save
                            '  .DeleteAfterSubmit = True
                            End With
                            
                            Set objMsg = Nothing
        
         
    
End Sub

Sub INSERE_QUADRO_GERAL_DE_METAS(Optional wb As Workbook)
    
    Dim ORIGEM As Workbook
    Dim sh As Worksheet
            
            
            If wb Is Nothing Then Set wb = ActiveWorkbook
            
            Set ORIGEM = Workbooks.Open(ARQUIVO_METAS, False, True) 'pega o arquivo q contem a planilha a ser copiada
            ORIGEM.Sheets("QUADRO_METAS").Copy After:= _
                         wb.Sheets(wb.Sheets.Count)
            ActiveSheet.Name = "METAS" 'nomeia a nova planilha
            If ActiveSheet.AutoFilterMode Then Cells.AutoFilter
            ORIGEM.Close (False)
End Sub

Public Function LISTA_CIDADES_NO_HISTORICO_DO_SET(oRegional As String, periodoSet, ano) As Variant
    
    Dim cidades() As Variant
    
    Dim pvt As PivotTable
    Dim fld As PivotField
            
    Set pvt = PVT_HISTORICO
            
    pvt.ClearTable
            
    
    Set fld = pvt.PivotFields("CIDADE")
    fld.Orientation = xlRowField
    
    strCidades = vbNullString
     'pvt.ColumnGrand = False
     'pvt.RowGrand = False
     
   ReDim cidades(1)
   oItem = 0
    For i = 0 To UBound(periodoSet)
    
          pvt.ClearTable
          pvt.PivotFields("ANO").Orientation = xlPageField: pvt.PivotFields("ANO").CurrentPage = CInt(ano) - 1
          pvt.PivotFields("MES").Orientation = xlPageField: pvt.PivotFields("MES").CurrentPage = periodoSet(i)
          On Error GoTo errorsolve
          pvt.PivotFields("REGIONAL").Orientation = xlPageField: pvt.PivotFields("REGIONAL").CurrentPage = oRegional
          pvt.RowGrand = False
          Set fld = pvt.PivotFields("CIDADE")
          fld.Orientation = xlRowField
         
          For Each c In fld.DataRange
          
             If InStr(1, c.Text, "Total Geral") > 0 Then GoTo pula
             
             If cidades(0) = vbEmpty Then
                    cidades(oItem) = c.Text
             Else
             
                    ReDim Preserve cidades(oItem)
                    cidades(oItem) = c.Text
             
             End If
             oItem = oItem + 1
          Next c
pula:
    Next i
   
    LISTA_CIDADES_NO_HISTORICO_DO_SET = removeDupesInArray(cidades)
Exit Function
errorsolve:
aa = Err.Number
dd = Err.Description
Err.Clear

Erase cidades

LISTA_CIDADES_NO_HISTORICO_DO_SET = cidades

End Function


Public Function LISTA_CIDADES_ATIVAS_NO_SET(oRegional As String, acumulado, ano) As Variant
    
    Dim cidades() As Variant
    
    Dim pvt As PivotTable
    Dim fld As PivotField
            
    Set pvt = PVT_REALIZADO_2017
            
    
            
    
   
    
    strCidades = vbNullString
     'pvt.ColumnGrand = False
     'pvt.RowGrand = False
     
   ReDim cidades(1)
   oItem = 0
    For i = 0 To UBound(acumulado)
    
          pvt.ClearTable
          pvt.PivotFields("ANO").Orientation = xlPageField: pvt.PivotFields("ANO").CurrentPage = CInt(ano)
          pvt.PivotFields("MES").Orientation = xlPageField: pvt.PivotFields("MES").CurrentPage = acumulado(i)
          On Error GoTo errorsolve
          pvt.PivotFields("REGIONAL").Orientation = xlPageField: pvt.PivotFields("REGIONAL").CurrentPage = oRegional
          pvt.RowGrand = False
          Set fld = pvt.PivotFields("Cidade")
          fld.Orientation = xlRowField
         
          For Each c In fld.DataRange
          
             If InStr(1, c.Text, "Total Geral") > 0 Then GoTo pula
             
             If cidades(0) = vbEmpty Then
                    cidades(oItem) = c.Text
             Else
             
                    ReDim Preserve cidades(oItem)
                    cidades(oItem) = c.Text
             
             End If
             oItem = oItem + 1
          Next c
pula:
    Next i
encerra:
    LISTA_CIDADES_ATIVAS_NO_SET = removeDupesInArray(cidades)
Exit Function
errorsolve:
aa = Err.Number
dd = Err.Description

'Erase cidades
GoTo encerra


End Function


Public Sub ENVIA_RELATORIO_GERAL_PONTUACAO_PARA_OS_REGIONAIS()



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
    Dim lin As Integer
    
    Dim objMsg As Outlook.MailItem
    Dim oApp As Outlook.Application
    Dim ns As Outlook.Namespace
    Dim fld As Outlook.Folder
    Dim oAccount As Outlook.Account
    
    Dim hdr As Range

    Dim ahdr As Range
    Dim clientes As Variant
    Dim agentes As Range
    Dim refData As String
    Dim strPath As String
    
    
    Dim relatorio As Workbook
    Dim rel As Range
    Dim relh As Range
    Dim delRange As Range
    
    Dim pptApp As PowerPoint.Application
    Dim pptPresentation As PowerPoint.Presentation
    Dim sl As Slide
    Dim pptLayout As CustomLayout
    Dim shp As PowerPoint.Shape
    Dim tbl As PowerPoint.Table
    Dim pres As PowerPoint.Presentation
    
    On Error GoTo error_handler
    
    Const dropbox As String = "C:\Dropbox\VEDACIT\"
    Set w = ActiveWorkbook
    
    Set addin = ThisWorkbook
             
         strPath = Application.GetOpenFilename(, , "Selecione o arquivo com os resultados do mês", "Selecione")
         If InStr(1, UCase(strPath), ":") = 0 Then Exit Sub
            
          Set fs = CreateObject("Scripting.FileSystemObject")
          Set f = fs.GetFile(strPath)
          a = Split(f.ParentFolder, "\")
          strMes = a(UBound(a))
          
          If strMes = "MARCO" Then strMes = "MARÇO"
    
            
    'AA===================================================
    '   transforma o relatorio das regionais em html
        Set relatorio = Workbooks.Open(strPath)
            Set s = relatorio.Sheets(1)
            s.Copy After:=s
            Set s = ActiveSheet
            
            Set rel = s.UsedRange
            rel.Copy
            rel.PasteSpecial xlPasteValues
            DeleteFilteredOutRows s
            Set relh = rel.Rows(1)
            rel.Select
            relh.Select
            
            rel.Sort key1:=relh.Find("TOTAL", lookat:=xlWhole), order1:=xlDescending, HEADER:=xlYes
                        
            primeira = rel(2, 1)
                        
            Set relh = relh.Offset(1)
            
            For Each c In relh.Cells
                If c.Text = vbNullString Then
                If delRange Is Nothing Then Set delRange = c Else Set delRange = Union(delRange, c)
                End If
            Next c
            
            delRange.Select
            delRange.EntireColumn.Delete
            
                        
        s.Name = "MAIL_BODY"
        
        strHtmlBody = RangetoHTML(s.UsedRange)
        
    
    'AA===================================================
    '------------------------------------------------------------------------------------------------------------------
    'BB===================================================
    '   monta as listas de envio
    
        Set wbContatos = Workbooks.Open("C:\Dropbox\VEDACIT_DADOS\Contatos.xlsx") 'pega o arquivo q contem a planilha a ser copiada
        Set cadastro = wbContatos.Sheets("Regionais")
        
        Set rng = cadastro.Rows(1).Find("mailRegional").Offset(1).Resize(cadastro.UsedRange.Rows.Count - 1)
        Set rng = Union(rng, cadastro.Rows(1).Find("mailExecutivo").Offset(1).Resize(cadastro.UsedRange.Rows.Count - 1))
        
        Set d = CreateObject("Scripting.Dictionary")
            c = rng
            For i = 1 To UBound(c, 1)
              d(c(i, 1)) = 1
            Next i
        strPara = vbNullString
        For Each k In d.Keys
        
            If strPara = vbNullString Then strPara = k Else strPara = strPara & ";" & k
        
        Next k
                
        strPara = strPara & ";marcelo.bastos@vedacit.com.br;silvia.sobreira@vedacit.com.br"
        'marcelo.bastos@vedacit.com.br;aline.sendo@grupobaumgart.com.br;gustavo.leme@grupobaumgart.com.br;andreia.stadnik@vedacit.com.br
        'aline.sendo@grupobaumgart.com.br>; Gustavo Mancanares Leme gustavo.leme@grupobaumgart.com.br
        strCopia = "aline.sendo@grupobaumgart.com.br;gustavo.leme@grupobaumgart.com.br" & _
                   "joao.ximenes@vedacit.com.br;" & _
                   "mauricio.gasperini@vedacit.com.br;" & _
                   "bruno.pacheco@vedacit.com.br;" & _
                   "andreia.stadnik@vedacit.com.br"
                   
        nomePrimeira = cadastro.Columns(1).Find(primeira, lookat:=xlWhole).Offset(0, 1)
        'Set rng = rng.AdvancedFilter(Action:=xlFilterCopy, CopyToRange:=cadastro.Range("A100"), Unique:=True)
      
    'BB===================================================
    '------------------------------------------------------------------------------------------------------------------
    'CC===================================================
    '     Let's create a new PowerPoint
          imgFile = "C:\Dropbox\VEDACIT_DADOS\Fundo.jpg"
      
          Set pptApp = New PowerPoint.Application
          Set pres = pptApp.Presentations.Add
              pres.PageSetup.SlideSize = ppSlideSizeA4Paper
              pres.PageSetup.SlideOrientation = msoOrientationVertical
       
                Set sl = pres.Slides.Add(1, ppLayoutBlank)
                    sl.Select
                    sl.FollowMasterBackground = msoFalse
                    sl.Background.Fill.UserPicture imgFile
        
            'insere a tabela
            'Paste to PowerPoint and position
                rel.Copy
                sl.Shapes.PasteSpecial DataType:=2   '2 = ppPasteEnhancedMetafile
                'Set tbl = sl.Shapes(sl.Shapes.Count)
                
                  'Set position:
                    'tbl.ScaleProportionally 70
                    pptApp.ActiveWindow.Selection.ShapeRange.Top = 200
                    pptApp.ActiveWindow.Selection.ShapeRange.Left = 10
                    pptApp.ActiveWindow.Selection.ShapeRange.ScaleWidth 0.68, msoTrue, msoScaleFromTopLeft
                    h1 = Round(pptApp.ActiveWindow.Selection.ShapeRange.Top + pptApp.ActiveWindow.Selection.ShapeRange.Height, 0)
                         str1 = "A meta de faturamento é um habilitador para pontuar em todos os itens da campanha."
                         'str1 = str1 & "Vocês são guerreiros e nao tem medo de desafios." & vbCrLf
                         
                     'sl.Shapes("texto").Delete
                    Set shp = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, 5, h1 + 5, 520, 20)
                    'shp.Top = 500
                    shp.TextFrame.TextRange.Text = str1
                    shp.TextFrame.TextRange.Font.Size = 10
                    shp.TextFrame.TextRange.Font.Bold = msoFalse
                    shp.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
                    shp.Name = "texto"
                            
                    str1 = UCase(nomePrimeira & ", parabens!!!")
                    str3 = str3 & "Você é o grande destaque de " & WorksheetFunction.Proper(strMes) & "!" & vbCrLf
                    str3 = str3 & "Mantenha o foco e toda essa garra. Com isso vai garantir presença na super liga!!!!" & vbCrLf & vbCrLf & vbCrLf
                    str3 = str3 & "Para aqueles que não pontuaram, é só o começo de um grande jogo!" & vbCrLf
                    str3 = str3 & "Mas não dá pra ficar só no aquecimento." & vbCrLf
                    str3 = str3 & "Vocês são guerreiros e nao tem medo de desafios." & vbCrLf
                    str3 = str3 & "Tem a torcida do seu lado. É jogar pra ganhar!!!" & vbCrLf
           
                
                   'sl.Shapes("chamada").Delete
                    Set shp = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, 5, h1 + 25, 520, 20)
                    'shp.Top = 500
                    shp.TextFrame.TextRange.Text = str1
                    shp.TextFrame.TextRange.Font.Size = 14
                    shp.TextFrame.TextRange.Font.Bold = msoFalse
                    shp.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
                    shp.Name = "chamada"
                
                    'sl.Shapes("chamada").Delete
                    Set shp = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, 5, h1 + 50, 520, 80)
                    'shp.Top = 500
                    shp.TextFrame.TextRange.Text = str3
                    shp.TextFrame.TextRange.Font.Size = 12
                    shp.TextFrame.TextRange.Font.Bold = msoFalse
                    shp.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
                    shp.Name = "txt"
    
                    
                    imagesPath = FolderFromPath(strPath) & "MAILIMAGES\"
                    If Dir(imagesPath, vbDirectory) = vbNullString Then MkDir imagesPath Else Delete_Folder CStr(imagesPath): MkDir CStr(imagesPath)
         
                    arquivoImagem = "informativo_regionais.jpg"
                    sl.Export imagesPath & arquivoImagem, "JPG"
            
    
    
    
    
    '       Prepara o email
            Set oApp = New Outlook.Application
            Set ns = oApp.GetNamespace("MAPI")
            Set fld = ns.GetDefaultFolder(olFolderOutbox)
                
                
                 
                 strAssunto = "RELATORIO PONTUAÇÃO REGIONAIS " & strMes
                       
                 
                 strObs = "OBS: anexo os pontos das regionais e representantes"
                       
                       
                            Set objMsg = fld.Items.Add(olMailItem)
                            
                             With objMsg
                              .To = strPara
                              .CC = strCopia
                              '.BCC = "bruno.pacheco@vedacit.com.br"
                              .Subject = strAssunto
                              .Categories = "Vedateam_Informativo_Regionais_" & strMes
                              '.VotingOptions = "Yes;No;Maybe;"
                              .BodyFormat = olFormatHTML
                              .Attachments.Add strPath, olByValue, 1
'                              .HTMLBody = strHtmlBody
                              .Display
                              .Save
                              '.Send 'UsingAccount = oAccount
                              
                              '.DeleteAfterSubmit = True
                            End With
                            
                              Set objMsg = Nothing
                  
    Application.DisplayAlerts = False
        relatorio.Close False
        wbContatos.Close False
        'pres.Close
   Application.DisplayAlerts = True
                  
Exit Sub
error_handler:
Err.Clear
Resume Next

End Sub












Function HTML_BODY_CAPILARIDADE(regio As String, _
                                nomePessoal As String, _
                                strSet As String, _
                                txtMes As String, _
                                txtAno As String, _
                                mset, _
                                mcor, _
                                cidH, _
                                cidR) As String

    Dim str As String
    Dim strMesesNoSet As String
    Dim strMessage As String
    Dim ipos As Integer
    
        strMesesNoSet = vbNullString
        strMesesNoSet = Split(meses, ",")(mset(0))
        strMesesNoSet = strMesesNoSet & ", " & Split(meses, ",")(mset(1))
        strMesesNoSet = strMesesNoSet & " e " & Split(meses, ",")(mset(2))
            
            
            
        If UBound(mset) = 1 Then
            'o set tem 2 meses
            If UBound(mcor) = 1 Then ipos = 3
            
            Else
            'o set tem 3 meses
              If UBound(mcor) = 1 Then ipos = 2 Else ipos = 3
        End If
        'se so temos um mes corrido então é o primeiro mes do set
        If UBound(mcor) = 0 Then ipos = 1
        
        Select Case ipos
        
            Case 1 'estamos no primeiro mes do set
                strPeriodo = ""
            Case 2 'estamos no meio set
                strPeriodo = ""
            Case 3 ' estamos no final do set
                strPeriodo = ""
                strPeriodo = vbNullString
                strPeriodo = Split(meses, ",")(mcor(0))
                strPeriodo = strPeriodo & ", " & Split(meses, ",")(mcor(1))
                If UBound(mcor) = 1 Then
                    strPeriodo = Replace(CStr(strPeriodo), ",", " e")
                Else
                    strPeriodo = strPeriodo & " e " & Split(meses, ",")(mcor(2))
                End If
        End Select

str = "<h2 style=" & chr(34) & "text-align: center;" & chr(34) & "><strong>" & UCase(regio) & "</strong></h2>"
str = str & "<p><strong><span style=" & chr(34) & "text-decoration: underline;" & chr(34) & ">COMO PONTUAR CAPILARIDADE NO " & strSet & " SET - " & txtMes & "/" & txtAno & "</span>:</strong></p>"
str = str & "<p><strong></br></br>Como vai " & nomePessoal & "?</br></br></strong></p>"
str = str & "<p><strong>Saudações da equipe VEDATEAM!</strong></p>"
str = str & "<p><strong>Apresentamos a seguir um resumo das cidades atendidas pela regional por meio de seus representantes." & _
            " No ano passado durante o mesmo periodo, a tua equipe ativou " & UBound(cidH) + 1 & "cidades.</strong></p></br>" & _
            " No " & strSet & " set da campanha, o total de cidades ativas, durante os meses de " & strPeriodo & ", foram ativadas " & UBound(cidR) + 1 & "cidades.</strong></p></br>"
str = str & "<p><strong>Para habilitar-se a obter os pontos de CAPILARIDADE, é preciso ter atingido 90% da meta de faturamento (habilitador) e manter a mesma cobertura de cidades realizada no mesmo periodo do ano anterior.</strong></p>"
str = str & "<p><strong>Caso esse numero seja mantido, e exista fauramento para alguma nova cidade, ou seja, cidades que não sejam no historico no periodo em questão," & _
            " a pontuação ocorrerá de acordo com a seguinte tabela:</strong></p></br></br>"

str = str & "<p>Num. Cidades Novas Pontos</p></br>"
str = str & "<p>        0            -33</p></br>"
str = str & "<p>        1             33</p></br>"
str = str & "<p>        2             47</p></br>"
str = str & "<p>        3             60</p></br>"
str = str & "<p>        4             73</p></br>"
str = str & "<p>        5             87</p></br>"
str = str & "<p>        6            100</p></br>"
str = str & "<p>        7            117</p></br>"
str = str & "<p>        8            133</p></br>"
str = str & "<p>        9            150</p></br>"
str = str & "<p>       10            200</p></br>"


'str = str & "<h2>LISTAGEM DE " & numCli & " CLIENTES ATIVOS NO SEMESTRE M&Oacute;VEL:</h2>"

'For n = 0 To UBound(arr): str = str & "<h3>" & UCase(arr(n)) & "<h3>": Next n


HTML_BODY_CAPILARIDADE = str


End Function

Public Function calcCapilaridade(nomeRegional, mes, ano)

Dim aRng As Range

faltaDados = False
result = 0
                   Select Case mes
                     
                        Case 4
                            numeroDoSet = 1
                            mesesNoSet = Array(2, 3, 4)
                            mesesCorridos = Array(2, 3, 4)
                        
                        Case 7
                            numeroDoSet = 2
                            mesesNoSet = Array(5, 6, 7)
                            mesesCorridos = Array(5, 6, 7)
                       
                        Case 10
                            numeroDoSet = 3
                            mesesNoSet = Array(8, 9, 10)
                            mesesCorridos = Array(8, 9, 10)
                       
                        Case 12
                            numeroDoSet = 4
                            mesesNoSet = Array(11, 12)
                            mesesCorridos = Array(11, 12)
                      
                        Case Else
                            GoTo finaliza
                    End Select

        
                 cidadesHist = LISTA_CIDADES_NO_HISTORICO_DO_SET(CStr(nomeRegional), mesesNoSet, ano)
                 cidadesReal = LISTA_CIDADES_ATIVAS_NO_SET(CStr(nomeRegional), mesesCorridos, ano)
                 If Not Len(Join(cidadesHist)) > 0 Then faltaDados = True: GoTo finaliza
                 
                 cidNoSet = UBound(cidadesHist) + 1
                 If Not Len(Join(cidadesReal)) > 0 Then faltaDados = True: GoTo finaliza
                 
                 cidNovas = COMPARA_STRINGS_E_RETORNA_DIFERENCAS(Join(cidadesHist, ","), Join(cidadesReal, ","))
                 
                 If Len(cidNovas) > 0 Then numCidNovas = UBound(Split(cidNovas, ",")) + 1 Else numCidNovas = 0

                
                Set aRng = ActiveWorkbook.Sheets("PARAMETROS").Rows(1).Find("CAPILARIDADE").Offset(numCidNovas + 1, 1)
                
                calcCapilaridade = aRng.Value



Exit Function
finaliza:
    calcCapilaridade = 0

End Function





Public Function dadosDoSet(mes) As sSet

     Select Case mes
                        
                        Case 1
                            Exit Function
                        Case 2
                            numeroDoSet = 1
                            mesesNoSet = Array(2, 3, 4)
                            mesesCorridos = Array(2)
                            strMesesNoSet = "FEVEREIRO, MARÇO e ABRIL"
                            strMesesCorridos = "FEVEREIRO"
                        Case 3
                            numeroDoSet = 1
                            mesesNoSet = Array(2, 3, 4)
                            mesesCorridos = Array(2, 3)
                            strMesesNoSet = "FEVEREIRO, MARÇO e ABRIL"
                            strMesesCorridos = "FEVEREIRO e MARÇO"
                        
                        Case 4
                            numeroDoSet = 1
                            mesesNoSet = Array(2, 3, 4)
                            mesesCorridos = Array(2, 3, 4)
                            strMesesNoSet = "FEVEREIRO, MARÇO e ABRIL"
                            strMesesCorridos = "SET COMPLETO"
                        
                        Case 5
                            numeroDoSet = 2
                            mesesNoSet = Array(5, 6, 7)
                            mesesCorridos = Array(5)
                            strMesesNoSet = "MAIO, JUNHO e JULHO"
                            
                        Case 6
                            numeroDoSet = 2
                            mesesNoSet = Array(5, 6, 7)
                            mesesCorridos = Array(5, 6)
                            strMesesNoSet = "MAIO, JUNHO e JULHO"
                            
                        Case 7
                            numeroDoSet = 2
                            mesesNoSet = Array(5, 6, 7)
                            mesesCorridos = Array(5, 6, 7)
                            strMesesNoSet = "MAIO, JUNHO e JULHO"
                        
                        Case 8
                            numeroDoSet = 3
                            mesesNoSet = Array(8, 9, 10)
                            mesesCorridos = Array(8)
                            strMesesNoSet = "AGOSTO, SETEMBRO e OUTUBRO"
                        
                        Case 9
                            numeroDoSet = 3
                            mesesNoSet = Array(8, 9, 10)
                            mesesCorridos = Array(8, 9)
                            strMesesNoSet = "AGOSTO, SETEMBRO e OUTUBRO"
                        
                        
                        Case 10
                            numeroDoSet = 3
                            mesesNoSet = Array(8, 9, 10)
                            mesesCorridos = Array(8, 9, 10)
                        
                        Case 11
                            numeroDoSet = 4
                            mesesNoSet = Array(11, 12)
                            mesesCorridos = Array(11)
                        
                        
                        Case 12
                            numeroDoSet = 4
                            mesesNoSet = Array(11, 12)
                            mesesCorridos = Array(11, 12)
                   
                    End Select
     
                
dadosDoSet.numSet = numeroDoSet
dadosDoSet.mesesQueCompoemSet = mesesNoSet
dadosDoSet.mesesQueJaSePassaram = mesesCorridos

End Function










