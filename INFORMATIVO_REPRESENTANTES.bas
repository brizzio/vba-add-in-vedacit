Attribute VB_Name = "INFORMATIVO_REPRESENTANTES"
Public Sub INFORMA_REPRESENTANTES()

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
    
    Const dropbox As String = "C:\Dropbox\VEDACIT\"
    
    
    Set addin = ThisWorkbook
    
    'response = MsgBox("Deseja preparar o informativo para os representantes?", vbYesNo, "VEDATEAM")
    'If response = vbNo Then Exit Sub

    If Dir(dropbox) = vbNullString Then MsgBox ("Esta maquina nao tem acesso aos arquivos de dados no Dropbox. Não é possivel executar a rotina!"): Exit Sub
    
    inputMes = InputBox("Entre o mês para o informativo", "VEDATEAM", Month(Date) - 1)
    inputAno = InputBox("Entre o ano...", "VEDATEAM", Year(Date))
    
    'inputMes = 4
    'inputAno = 2017
    
    mesExtenso = Split(meses, ",")(inputMes - 1)
    
    
    '=============================== prepara o diretorio para armazenar os arquivos
   
    'dirPath = dropbox & "INFORMATIVO_REPRESENTANTES"
    'If Dir(dirPath, vbDirectory) = vbNullString Then
    '        MkDir dirPath
    '    Else
    '        If Right(dirPath, 1) <> "\" Then dirPath = dirPath & "\"
    '        ' dirPath & "*.*"
    '        RmDir dirPath
    '        MkDir dirPath
    'End If
    '------------------------------------------------------------------------------
               
    mes = inputMes
    ano = inputAno
    
    Set w = ActiveWorkbook
    Set s = w.Worksheets.Add(Before:=w.Sheets(1))
    titulos = Split("REGIONAL,REPRESENTANTE,NOME,MEDIA CLIENTES ATIVOS,CLIENTES ATIVOS NO PERIODO,PARA,COPIA", ",")
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
     sheetname = "CADREPRE"
        If Not Evaluate("ISREF('" & sheetname & "'!A1)") Then
               Set ORIGEM = Workbooks.Open("C:\Dropbox\VEDACIT_DADOS\Contatos.xlsx") 'pega o arquivo q contem a planilha a ser copiada
               ORIGEM.Sheets("Representantes").Copy After:= _
                            w.Sheets(w.Sheets.Count)
               ORIGEM.Close (False)
               ActiveSheet.Name = sheetname
        End If
    'Prepara o email
      Set oApp = New Outlook.Application
      Set ns = oApp.GetNamespace("MAPI")
    Set fld = ns.GetDefaultFolder(olFolderOutbox)
    
                          'seleciona a conta para envio
                          For Each obj In oApp.Session.Accounts
                             
                             If obj = conta_vedateam Then Set oAccount = obj
                             
                          Next obj
    
    Set metas = w.Sheets("METAS")
    
    Set agentes = metas.Range(metas.Cells(2, 1).Address & ":" & metas.Cells(metas.UsedRange.Rows.Count, 1).Address)
     
     'dados comuns para todos os emails
                refData = "01/" & mes - 1 & "/" & ano
                txtPeriodo = Split(meses, ",")(Month(DateAdd("m", -6, CDate(refData))) - 1) & " de " & Year(DateAdd("m", -6, CDate(refData))) & _
                                         " a " & Split(meses, ",")(Month(DateAdd("m", -1, CDate(refData))) - 1) & " de " & Year(DateAdd("m", -1, CDate(refData)))
                txtPeriodo = StrConv(CStr(txtPeriodo), vbLowerCase)
                   
     'dados especificos para cada email
            
            For Each c In agentes.Cells
            
                 If c.Interior.ColorIndex = 15 Then aRegional = c.Text: GoTo skipThat
                 If InStr(1, c.Text, "REG.") > 0 Then c.Interior.ColorIndex = 3: GoTo skipThat
                 If InStr(1, c.Text, "REG ") > 0 Then c.Interior.ColorIndex = 3: GoTo skipThat
                 
                 nomeRep = c.Text
                 clientes = LISTA_CLIENTES_NO_HISTORICO(c.Text, "01/" & mes - 1 & "/" & ano)
                 If Not Len(Join(clientes)) > 0 Then c.Interior.ColorIndex = 3: GoTo skipThat
                 cliNoPeriodo = UBound(clientes) + 1
                 mediaClientesAtivadosMensalmente = Round(cliNoPeriodo / 6, 0)
                 strListagem = vbNullString
                 'For n = 0 To UBound(clientes): strListagem = strListagem & UCase(clientes(n)) & vbCrLf: Next n
                 endereco = Replace(w.Sheets("CADREPRE").Cells.Find(c.Text).Offset(0, 4), "/", ";")
                 nomePessoal = w.Sheets("CADREPRE").Cells.Find(c.Text).Offset(0, 2)
                       'seleciona a conta para envio
                       
                       
                            Set objMsg = fld.Items.Add(olMailItem)
                            
                             With objMsg
                              .To = Replace(endereco, "/", ";")
                              '.CC = ""
                              '.BCC = "bruno.pacheco@vedacit.com.br"
                              .Subject = "INFORMATIVO CAMPANHA VEDATEAM - CLIENTES ATIVOS - " & mesExtenso & "/" & ano
                              .Categories = "Vedateam_Informativo_Clientes_Ativos"
                              '.VotingOptions = "Yes;No;Maybe;"
                              .BodyFormat = olFormatHTML
                              .HTMLBody = MAKE_HTML_BODY(c.Text, CStr(nomePessoal), CStr(txtPeriodo), CStr(mesExtenso), CStr(ano), CStr(mediaClientesAtivadosMensalmente), cliNoPeriodo, clientes)
                              
                              '.Display
                              .Send 'UsingAccount = oAccount
                              
                              '.DeleteAfterSubmit = True
                            End With
                            
                              Set objMsg = Nothing
                     
skipThat:
            
                lin = s.UsedRange.Rows.Count + 1
                's.Cells(lin, 1).Select
                s.Cells(lin, 1) = UCase(aRegional)
                s.Cells(lin, 2) = UCase(c.Text)
                s.Cells(lin, 3) = CStr(nomePessoal)
                
                s.Cells(lin, 4) = mediaClientesAtivadosMensalmente
                s.Cells(lin, 5) = cliNoPeriodo
                s.Cells(lin, 6) = endereco
            
            'zera as variaveis
            
                            nomeRep = vbNullString
                            cliNoPeriodo = vbNullString
                            mediaClientesAtivadosMensalmente = vbNullString
                            strListagem = vbNullString
                            endereco = vbNullString
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
        htmlHead = htmlHead & "<br><br>PERÍODO BASE: " & txtPeriodo & "<br><br>"
        
        

        Set objMsg = fld.Items.Add(olMailItem)
                            
                             With objMsg
                              .To = "bruno.pacheco@vedacit.com.br"
                              .CC = "vedateam@outlook.com"
                              
                              .Subject = "RELATORIO DE ENVIO EMAIL PARA REPRESENTANTES - CLIENTES ATIVOS - " & mesExtenso & "/" & ano
                              .Categories = "Relatorio_Email_Marketing_Clientes_Ativos"
                              '.VotingOptions = "Yes;No;Maybe;"
                              .BodyFormat = olFormatHTML
                              .HTMLBody = UCase(htmlHead)
                              .HTMLBody = .HTMLBody & RangetoHTML(s.UsedRange)
                              .HTMLBody = .HTMLBody & "<br><br>======================== FAC SIMILE DO EMAIL ENVIADO =====================================<br><br>"
                              .HTMLBody = .HTMLBody & MAKE_HTML_BODY(UCase("Razao Social Representante"), UCase("Nome do Representante"), UCase("periodo base"), "MES", "ANO", "MEDIA", "NUMERO DE CLIENTES", "XXXXXXXXX")
                            
                              '.Display
                              .SendUsingAccount = oAccount
                              
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

Public Function LISTA_CLIENTES_NO_HISTORICO(oRepresentante As String, dataDeReferencia As String) As Variant
    
    Dim clientes() As Variant
    
    Dim pvt As PivotTable
    Dim fld As PivotField
            
    Set pvt = PVT_HISTORICO
            
    pvt.ClearTable
            
    
    Set fld = pvt.PivotFields("CLIENTE")
    fld.Orientation = xlRowField
    
    strClientes = vbNullString
     'pvt.ColumnGrand = False
     'pvt.RowGrand = False
     
   ReDim clientes(1)
   oItem = 0
    For i = -6 To -1
    
         
          pvt.PivotFields("ANO").Orientation = xlPageField: pvt.PivotFields("ANO").CurrentPage = Year(DateAdd("m", i, CDate(dataDeReferencia)))
          pvt.PivotFields("MES").Orientation = xlPageField: pvt.PivotFields("MES").CurrentPage = Month(DateAdd("m", i, CDate(dataDeReferencia)))
          On Error GoTo errorsolve
          pvt.PivotFields("REPRESENTANTE").Orientation = xlPageField: pvt.PivotFields("REPRESENTANTE").CurrentPage = oRepresentante
          pvt.RowGrand = False
          'fld.DataRange.Select
         
          For Each c In fld.DataRange
          
             If InStr(1, c.Text, "Total Geral") > 0 Then GoTo pula
             
             If clientes(0) = vbEmpty Then
                    clientes(oItem) = c.Text
             Else
             
                    ReDim Preserve clientes(oItem)
                    clientes(oItem) = c.Text
             
             End If
             oItem = oItem + 1
          Next c
pula:
    Next i
   
    LISTA_CLIENTES_NO_HISTORICO = removeDupesInArray(clientes)
Exit Function
errorsolve:
aa = Err.Number
dd = Err.Description

Erase clientes

LISTA_CLIENTES_NO_HISTORICO = clientes




End Function


Function MAKE_HTML_BODY(repres As String, _
                        nomePessoal As String, _
                        periodo As String, _
                        txtMes As String, _
                        txtAno As String, _
                        txtMedia As String, _
                        numCli, _
                        Optional lista As Variant) As String

Dim str As String

arr = lista
        
str = "<h1 style=" & chr(34) & "text-align: center;" & chr(34) & "><strong>" & UCase(repres) & "</strong></h1>"
str = str & "<p><strong><span style=" & chr(34) & "text-decoration: underline;" & chr(34) & ">DICAS PARA PONTUAR CLIENTES ATIVOS EM " & txtMes & "/" & txtAno & "</span>:</strong></p>"
str = str & "<p><strong></br></br>Como vai " & nomePessoal & "?</br></br></strong></p>"
str = str & "<p><strong>Sauda&ccedil;&otilde;es da equipe VEDATEAM!</strong></p>"
str = str & "<p><strong>Apresentamos a seguir um resumo dos clientes ativos (faturados) no periodo que compreende os meses de " & periodo & _
            ". No decorrer deste periodo, a sua m&eacute;dia de clientes ativos foi&nbsp;</strong></p></br>"
str = str & "<h2><span style= " & chr(34) & "color: #ff9900;" & chr(34) & "><strong>" & txtMedia & " clientes/mes</strong></span></h2></br>"
str = str & "<p><strong>Para que voc&ecirc; possa habilitar-se a obter os pontos de CLIENTES ATIVOS, voc&ecirc; precisa ter atingido 90% da meta de faturamento (habilitador) e manter a m&eacute;dia acima de clientes faturados.</strong></p>"
str = str & "<p><strong>Caso voc&ecirc; mantenha a media e fature para algum cliente novo, ou seja, que não sejam clientes faturados no periodo em questão," & _
            "você receberá a pontuação de acordo com a seguinte tabela:</strong></p></br></br>"
str = str & "<table width=" & chr(34) & "294" & chr(34) & ">"
str = str & "<tbody>"
str = str & "<tr>"
str = str & "<td style=" & chr(34) & "text-align: center;" & chr(34) & "width=" & chr(34) & "230" & chr(34) & ">"
str = str & "<h3><strong>Acréscimo de clientes (semestre m&oacute;vel)</strong></h3>"
str = str & "</td>"
str = str & "<td width=" & chr(34) & "64" & chr(34) & ">"
str = str & "<h3>Pontos</h3>"
str = str & "</td>"
str = str & "</tr>"
str = str & "<tr>"
str = str & "<td>Redu&ccedil;&atilde;o ou 0</td>"
str = str & "<td>-50</td>"
str = str & "</tr>"
str = str & "<tr>"
str = str & "<td>+1</td>"
str = str & "<td>50</td>"
str = str & "</tr>"
str = str & "<tr>"
str = str & "<td>+2</td>"
str = str & "<td>70</td>"
str = str & "</tr>"
str = str & "<tr>"
str = str & "<td>+3</td>"
str = str & "<td>90</td>"
str = str & "</tr>"
str = str & "<tr>"
str = str & "<td>+4</td>"
str = str & "<td>110</td>"
str = str & "</tr>"
str = str & "<tr>"
str = str & "<td>+5</td>"
str = str & "<td>130</td>"
str = str & "</tr>"
str = str & "<tr>"
str = str & "<td>+6</td>"
str = str & "<td>150</td>"
str = str & "</tr>"
str = str & "<tr>"
str = str & "<td>+7</td>"
str = str & "<td>175</td>"
str = str & "</tr>"
str = str & "<tr>"
str = str & "<td>+8</td>"
str = str & "<td>200</td>"
str = str & "</tr>"
str = str & "<tr>"
str = str & "<td>+9</td>"
str = str & "<td>225</td>"
str = str & "</tr>"
str = str & "<tr>"
str = str & "<td>+10 ou acima</td>"
str = str & "<td>300</td>"
str = str & "</tr>"
str = str & "</tbody>"
str = str & "</table>"
str = str & "<p>&nbsp;</p>"
'str = str & "<h2>LISTAGEM DE " & numCli & " CLIENTES ATIVOS NO SEMESTRE M&Oacute;VEL:</h2>"

'For n = 0 To UBound(arr): str = str & "<h3>" & UCase(arr(n)) & "<h3>": Next n


MAKE_HTML_BODY = str


End Function












