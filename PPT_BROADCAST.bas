Attribute VB_Name = "PPT_BROADCAST"
Private Const strPrimeiro = ", como bom estrategista, você sabe que os primeiros pontos são fundamentais." & vbCrLf & vbCrLf & _
                           "Para o SET total fará a diferença!" & vbCrLf & vbCrLf & _
                           "Parabéns! É o grande destaque do mês!" & vbCrLf & _
                           "Mantenha o foco e toda essa garra. Com isso vai garantir presença na Super Liga, e mais, leva um Prêmio especial no final deste SET!"

'incluir o "Parabéns" &  XXXXXXX & "!" & vbCrLf & vbCrLf  antes da fase
Private Const strSegundo = "O jogo começou, são grandes os desafios, mas você é guerreiro e já está na segunda posição!" & vbCrLf & vbCrLf & _
                           "Força e Foco. O objetivo final é chegar na Super Liga. Mas, se continuar pontuando com este afinco, levará também um Prêmio especial no final deste SET!" & vbCrLf & vbCrLf & _
                           "Vamos para cima!!!"

Private Const strTerceiro = ", como é estar em os três melhores da equipe?" & vbCrLf & vbCrLf & _
                           "Parabéns!" & vbCrLf & vbCrLf & _
                           "Os pontos só confirmam sua força, coragem e determinação." & vbCrLf & _
                           "Vamos em frente, dando o nosso melhor!" & vbCrLf & _
                           "Garanta o seu lugar na Super Liga e o Prêmio especial do final deste SET!"

Private Const str4a10 = "O segundo SET mal começou e você já faz parte da lista dos 20 melhores em quadra." & vbCrLf & vbCrLf & _
                           "Parabéns!" & vbCrLf & vbCrLf & _
                           "Atenção para não desviar o foco. Os pontos iniciais fazem a diferença, podem definir quem é o Campeão." & vbCrLf & _
                           "Ponto a ponto a Super Liga fica mais próxima." & vbCrLf & _
                           "E se continuar entre os '20 mais' o Prêmio especial do final deste SET será seu!"


Private Const strDemais = "É só o começo de um grande jogo!" & vbCrLf & vbCrLf & _
                           "Mas, não dá para ficar só no aquecimento. Os primeiros pontos podem decidir a partida." & vbCrLf & vbCrLf & _
                           "Como o segundo SET mal começou, com garra e determinação, vamos para cima!!!" & vbCrLf & _
                           "Você é guerreiro e não tem medo de desafios." & vbCrLf & _
                           "Tem a torcida do seu lado." & vbCrLf & _
                           "É jogar para ganhar!!!"


Sub geraPPTresultados() '(caminhoDoArquivo As String)
    
    Dim w As Workbook
    Dim s As Worksheet
    Dim cad As Worksheet
    Dim r As Range
    Dim ORIGEM As Workbook
    Dim sh As Worksheet
    Dim sht As Worksheet
    Dim aRegional As String
    
    Dim pptApp As PowerPoint.Application
    Dim pptPresentation As PowerPoint.Presentation
    
    Dim nomePessoalRep As String
    Dim razaoSocialRep As String
    Dim aRank As Integer
    Dim ptosFat As Integer
    Dim ptosCat As Integer
    Dim ptosMix As Integer
    Dim aTotalPontos As Integer
    Dim emailRep As String
    Dim bloco As String
    
    Set w = ActiveWorkbook
    Set s = w.Sheets("TOTAIS REPRESENTANTES")
    s.Select
    txtMes = mesExtenso
    'txtMes = w.Name
    'txtMes = Mid(txtMes, InStr(1, txtMes, "_") + 1, InStr(InStr(1, txtMes, "_") + 1, txtMes, "_") - InStr(1, txtMes, "_") - 1)
    If UCase(txtMes) = "MARCO" Then txtMesRelatorio = "MARÇO" Else txtMesRelatorio = txtMes
    
    s.Columns("A:C").Sort key1:=Range("C2"), order1:=xlDescending, HEADER:=xlYes
    
     sheetname = "CADREPRE"
        If Not Evaluate("ISREF('" & sheetname & "'!A1)") Then
               Set ORIGEM = Workbooks.Open("C:\Dropbox\VEDACIT_DADOS\Contatos.xlsx") 'pega o arquivo q contem a planilha a ser copiada
               ORIGEM.Sheets("Representantes").Copy After:= _
                            w.Sheets(w.Sheets.Count)
               ORIGEM.Close (False)
               ActiveSheet.Name = sheetname
        End If
    
    
    'Let's create a new PowerPoint
    If pptApp Is Nothing Then Set pptApp = New PowerPoint.Application
    If pptApp.Presentations.Count = 0 Then Set pptPresentation = pptApp.Presentations.Add Else Set pptPresentation = pptApp.ActivePresentation
    
    
    numeroDeRepresentantes = s.UsedRange.Rows.Count
    
    For i = 2 To numeroDeRepresentantes
        Set sht = Nothing
        Set sht = w.Sheets(CStr(s.Cells(i, 1)))
        aRegional = sht.Cells(2, 1)
        aRepre = s.Cells(i, 2)
        aTotalPontos = s.Cells(i, 3)
        aRank = i - 1
    
        
        sht.Activate
        
        linha = sht.Rows(1).Find("REPRESENTANTES").EntireColumn.Find(aRepre).Row
        
        ptosFat = sht.Cells(linha, sht.Rows(1).Find("PONTOS FATURAMENTO REPRESENTANTE").Column)
        ptosCat = sht.Cells(linha, sht.Rows(1).Find("CLIENTES ATIVOS").Column)
        ptosMix = sht.Cells(linha, sht.Rows(1).Find("MIX").Column)
        
        Set cad = w.Sheets("CADREPRE")
        cad.Activate
        If InStr(1, Replace(UCase(aRepre), ".", " "), "REG ") > 0 Then GoTo pulaEsse
        If InStr(1, Replace(UCase(aRepre), ".", " "), "HOME") > 0 Then GoTo pulaEsse
        If InStr(1, Replace(UCase(aRepre), ".", " "), "BR SAMP") > 0 Then GoTo pulaEsse
        cadLin = cad.Rows(1).Find("APELIDO").EntireColumn.Find(aRepre).Row
        
        nomePessoalRep = cad.Cells(cadLin, cad.Rows(1).Find("NOME").Column)
        razaoSocialRep = cad.Cells(cadLin, cad.Rows(1).Find("RAZAOSOCIAL").Column)
        emailRep = Replace(cad.Cells(cadLin, cad.Rows(1).Find("EMAIL").Column), "/", ",")
        
        'cria um bloco para envio aos representantes com faturamento 0
        If aRank > 20 And aTotalPontos = 0 Then
            bloco = bloco & ";" & Replace(emailRep, ",", ";")
        End If
        
        If i = numeroDeRepresentantes Then
        
            bloco = Right(bloco, Len(bloco) - 1)
            makeSlide pptPresentation, aRegional, nomePessoalRep, razaoSocialRep, CInt(aRank), ptosFat, ptosCat, ptosMix, aTotalPontos, CStr(bloco), CStr(txtMesRelatorio)
        
        Else
            makeSlide pptPresentation, aRegional, nomePessoalRep, razaoSocialRep, CInt(aRank), ptosFat, ptosCat, ptosMix, aTotalPontos, Replace(emailRep, ",", ";"), CStr(txtMesRelatorio)
        End If
        
        'cria um slide para cada representante, independente da pontuação de faturamento
        'makeSlide pptPresentation, aRegional, nomePessoalRep, razaoSocialRep, CInt(aRank), ptosFat, ptosCat, ptosMix, aTotalPontos, Replace(emailRep, ",", ";"), CStr(txtMesRelatorio)
        
        


pulaEsse:
    Next i
    
    pptPresentation.SaveAs "C:\Dropbox\VEDACIT\RELATORIOS\" & txtMes & "\" & "Comunicado_Representantes_" & txtMes & ".ppt"
    pptPresentation.Close

End Sub

Sub makeSlide(pres As PowerPoint.Presentation, _
              nomeRegional As String, _
              nomeRep As String, _
              rzSocRep As String, _
              rank As Integer, _
              pontosFaturamento As Integer, _
              pontosClientesAtivos As Integer, _
              pontosMix As Integer, _
              totalPontuacao As Integer, _
              email As String, _
              txtMes As String)
    
    On Error GoTo trapError

    'Dim pptApp As PowerPoint.Application
    'Dim pres As PowerPoint.Presentation
    Dim sl As Slide
    Dim pptLayout As CustomLayout
    Dim shp As PowerPoint.Shape
    Dim tbl As PowerPoint.Table
    Dim bloco As String
            
    imgFile = "C:\Dropbox\VEDACIT_DADOS\Fundo.jpg"
    
    strTitulo = nomeRep
    strSubTitulo = "(" & rzSocRep & ")"
    
    If CInt(pontosFaturamento) > 0 Then
            pFat = CInt(pontosFaturamento)
            pCat = CInt(pontosClientesAtivos)
            pMix = CInt(pontosMix)
            pTotal = CInt(totalPontuacao)
        Else
            pFat = 0
            pCat = 0
            pMix = 0
            pTotal = 0
    End If
   
    
    If pres.Slides.Count = 0 Then
        pres.PageSetup.SlideSize = ppSlideSizeA4Paper
        pres.PageSetup.SlideOrientation = msoOrientationVertical
        'insereslide mestre
    End If
    
 
        Set sl = pres.Slides.Add(pres.Slides.Count + 1, ppLayoutBlank)
        sl.Select
        sl.FollowMasterBackground = msoFalse
        sl.Background.Fill.UserPicture imgFile
        
        'sl.Shapes("titulo").Delete
        Set shp = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, 5, 200, 500, 250)
        shp.TextFrame.TextRange.Text = strTitulo
        shp.TextFrame.TextRange.Font.Size = 26
        shp.TextFrame.TextRange.Font.Bold = msoTrue
        shp.TextFrame.HorizontalAnchor = msoAnchorCenter
        shp.Name = "titulo"
        sl.Select
        shp.Select
        
        'insere sub titulo
        'sl.Shapes("subtit").Delete
        Set shp = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, 1, 230, 535, 250)
        shp.TextFrame.TextRange.Text = strSubTitulo
        shp.TextFrame.TextRange.ChangeCase ppCaseUpper
        shp.TextFrame.TextRange.Font.Size = 22
        shp.TextFrame.TextRange.Font.Bold = msoTrue
        shp.TextFrame.HorizontalAnchor = msoAnchorCenter
        shp.Name = "subtit"
        
        'insere texto
        Select Case rank
            Case 1
                strTexto = nomeRep & strPrimeiro
            Case 2
                strTexto = "Parabéns " & nomeRep & "!" & vbCrLf & vbCrLf & strSegundo
            Case 3
                strTexto = nomeRep & strTerceiro
            Case 4 To 20
                strTexto = str4a10
            Case Else
            
                strTexto = strDemais
          
        End Select
        
        'sl.Shapes("texto").Delete
        Set shp = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, 5, 420, 530, 450)
        shp.TextFrame.TextRange.Text = strTexto
        shp.TextFrame.TextRange.Font.Size = 16
        shp.TextFrame.TextRange.Font.Bold = msoFalse
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
        shp.Name = "texto"
            
        txtPontuacao = vbNullString
        txtPontuacao = "PONTUAÇÃO - " & txtMes & ":" & vbCrLf & vbCrLf
        txtPontuacao = txtPontuacao & "Pontos Faturamento: " & pFat & vbCrLf
        txtPontuacao = txtPontuacao & "Pontos Clientes Ativos: " & pCat & vbCrLf
        txtPontuacao = txtPontuacao & "Pontos Mix de Produtos: " & pMix & vbCrLf
        txtPontuacao = txtPontuacao & "Total de Pontos Realizado: " & pTotal & vbCrLf
            
        'insere pontos faturamento
        'sl.Shapes("titfat").Delete
        Set shp = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, 1, 280, 530, 250)
        shp.TextFrame.TextRange.Text = CStr(txtPontuacao)
        shp.TextFrame.TextRange.Font.Size = 18
        shp.TextFrame.TextRange.Font.Bold = msoTrue
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
        shp.Name = "pontos"
        
        
        'insere a observação no fim do relatorio
        Set shp = sl.Shapes.AddTextbox(msoTextOrientationHorizontal, 1, 620, 530, 100)
        shp.TextFrame.TextRange.Text = "OBS: Caso seja constatada divergência na pontuação apurada devido à analises posteriores, as eventuais correções serão aplicadas retroativamente para não prejudicar o desempenho do participante na campanha."
        shp.TextFrame.TextRange.Font.Size = 10
        shp.TextFrame.TextRange.Font.Bold = msoFalse
        shp.TextFrame.TextRange.ParagraphFormat.Alignment = ppAlignCenter
        shp.Name = "observacao"
        
        
        sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.InsertAfter "<email>" & email & "</email>"
        sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.InsertAfter "<regional>" & nomeRegional & "</regional>"
        
        
        'sendMessage sl
        
finaliza:
        
        Exit Sub
        
        
trapError:
        aaa = Err.Description
        Err.Clear
        GoTo finaliza
    

End Sub


Sub sendMessage(ByRef oSlide)

Dim tempPath As String

Dim CDOmsg As CDO.Message


Set CDOmsg = New CDO.Message
 
tempPath = "C:\TEMP"
If Dir(tempPath, vbDirectory) = vbNullString Then MkDir tempPath
oSlide.Export tempPath & "\SLIDE.jpg", "JPG"


With CDOmsg.Configuration.Fields
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "fabrizio@altcom.net.br"
    .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "80sao20ver"
    .Item("http://schemas.microsoft.com/cdo/configuration/smptserverport") = 25
    .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 'NTLM method
    .Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
    .Update
End With
' build email parts

Dim strHTML As String
        strHTML = "<!DOCTYPE HTML PUBLIC ""-//IETF//DTD HTML//EN"">" & NL
        strHTML = strHTML & "<HTML>"
        strHTML = strHTML & "  <HEAD>"
        strHTML = strHTML & "    <TITLE>EXEMPLOOOOO</TITLE>"
        strHTML = strHTML & "  </HEAD>"
        strHTML = strHTML & "  <BODY>"
        strHTML = strHTML & "   <IMG src=""SLIDE.jpg"">"
        strHTML = strHTML & "  </BODY>"
        strHTML = strHTML & "</HTML>"
        
        CDOmsg.HTMLBody = strHTML
        
        CDOmsg.AddRelatedBodyPart tempPath & "\SLIDE.jpg", "SLIDE.jpg", cdoRefTypeId




        With CDOmsg
            .Subject = "teste automacao"
            .From = "fabrizio@altcom.net.br"
            .To = "fabrizio.salvade@gmail.com" '"bruno.pacheco@vedacit.com.br"
            '.cc = "fabrizio.salvade@gmail.com"
            .BCC = ""
            .Subject = "mais um teste de envio automatico!!!!"
        End With
        
        CDOmsg.Send
        
finaliza:
 
        'If Right(tempPath, 1) <> "\" Then caminho = tempPath & "\"
        'Kill caminho & "*.*"
        'RmDir "C:\TEMP"
        
        Set CDOmsg = Nothing
        
        If Right(tempPath, 1) <> "\" Then tempPath = tempPath & "\"
        Kill tempPath & "*.*"
        RmDir "C:\TEMP"

End Sub


Sub ENVIA_COMUNICADO_REPRESENTANTE()
        
        Dim wbContatos As Workbook
        Dim cadastro As Worksheet
        

         Dim pptApp As New PowerPoint.Application
         Dim pptPresentation As PowerPoint.Presentation
         Dim sl As PowerPoint.Slide
         Dim shp As PowerPoint.Shape
         
            Dim objMsg As Outlook.MailItem
            Dim oApp As Outlook.Application
            Dim ns As Outlook.Namespace
            Dim fld As Outlook.Folder
            Dim oAccount As Outlook.Account
            Dim att As Outlook.Attachment
            Dim atts As Outlook.Attachments
            Dim arquivoImagem As String
            Dim strPath As String
            Dim imagesPath As String
         
         On Error GoTo error_handler
         
         strPath = Application.GetOpenFilename(, , "Selecione o arquivo com os resultados do mês", "Selecione")
         If InStr(1, UCase(strPath), ":") = 0 Then Exit Sub
         
         imagesPath = FolderFromPath(strPath) & "IMAGES\"
         If Dir(imagesPath, vbDirectory) = vbNullString Then MkDir imagesPath Else Delete_Folder imagesPath: MkDir imagesPath
               'abre CADASTRO DE REGIONAIS
               Set wbContatos = Workbooks.Open("C:\Dropbox\VEDACIT_DADOS\Contatos.xlsx") 'pega o arquivo q contem a planilha a ser copiada
               Set cadastro = wbContatos.Sheets("Regionais")
               
                           
              'INICIA O OUTLOOK
                Set oApp = New Outlook.Application
                Set ns = oApp.GetNamespace("MAPI")
                Set fld = ns.GetDefaultFolder(olFolderOutbox)
        
                     'seleciona a conta para envio
                     For Each obj In oApp.Session.Accounts
                        
                        If obj = conta_vedateam Then Set oAccount = obj
                        
                     Next obj
                           
            'abre o arquivo
                           
            Set pptApp = New PowerPoint.Application
            Set pptPresentation = pptApp.Presentations.Open(strPath)
            
            For Each sl In pptPresentation.Slides
                    
                    txtNota = sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange
                    stag = "email"
                    tagini = InStrRev(txtNota, "<" & stag & ">") + Len(stag) + 2 ' position of start delimiter
                    tagfim = InStr(txtNota, "</" & stag & ">") ' position of end delimiter
                    endereco = Mid(txtNota, tagini, tagfim - tagini)
                    
                    stag = "regional"
                    tagini = InStrRev(txtNota, "<" & stag & ">") + Len(stag) + 2 ' position of start delimiter
                    tagfim = InStr(txtNota, "</" & stag & ">") ' position of end delimiter
                    aRegional = Mid(txtNota, tagini, tagfim - tagini)
                    mailRegional = cadastro.Columns(1).Find(aRegional, lookat:=xlWhole).Offset(0, 2)
                    nomeRepresentante = sl.Shapes(2).TextFrame.TextRange.Text
                    nomeRepresentante = cleanString(CStr(nomeRepresentante))
                    
                    
                    arquivoImagem = "slide_" & sl.SlideIndex & "_" & nomeRepresentante & ".jpg"
                    sl.Export imagesPath & arquivoImagem, "JPG"
            
                            Set objMsg = oApp.CreateItem(olMailItem)
                            'Set atts = objMsg.Attachments
                            'Set att = atts.Add(imagesPath & arquivoImagem, olByValue, 0)
                            
                            'att.PropertyAccessor.SetProperty "http://schemas.microsoft.com/mapi/proptag/0x370E001F", "image/jpeg"  'Change From 0x370eE001E
                            'att.PropertyAccessor.SetProperty "http://schemas.microsoft.com/mapi/proptag/0x3712001F", arquivoImagem     'Changed from 0x3712001E
                            'objMsg.PropertyAccessor.SetProperty "http://schemas.microsoft.com/mapi/id/{00062008-0000-0000-C000-000000000046}/8514000B", True
                            
                             With objMsg
                              .To = Replace(endereco, "/", ";")
                              .CC = mailRegional
                              .BCC = conta_vedateam
                              .Subject = "INFORMATIVO CAMPANHA VEDATEAM - PONTUAÇÃO"
                              .Categories = "Vedateam_Informativo_Pontuacao"
                              '.VotingOptions = "Yes;No;Maybe;"
                              .BodyFormat = olFormatHTML
                              
                              '.AddRelatedBodyPart tempPath & "\SLIDE.jpg", "SLIDE.jpg", cdoRefTypeId
                              
                              .Attachments.Add imagesPath & arquivoImagem, olByValue, 0
                              .HTMLBody = .HTMLBody & "<p align=""center"">" & "<img src='cid:" & arquivoImagem & "'>" & "</p>"
                              
                              .Save
                              .Display
                              .Send 'UsingAccount = oAccount
                              
                              '.DeleteAfterSubmit = True
                            End With
                            
                              Set objMsg = Nothing
                     
            Next sl

finaliza:
wbContatos.Close False
pptPresentation.Close

Exit Sub
error_handler:
eee = Err.Description
Err.Clear
Resume Next
GoTo finaliza


       
End Sub


Sub EDITA_PPT_PARA_ENVIO_MANUAL()
        
        Dim wbContatos As Workbook
        Dim cadastro As Worksheet
        

         Dim pptApp As New PowerPoint.Application
         Dim pptPresentation As PowerPoint.Presentation
         Dim sl As PowerPoint.Slide
         Dim shp As PowerPoint.Shape
         
            Dim objMsg As Outlook.MailItem
            Dim oApp As Outlook.Application
            Dim ns As Outlook.Namespace
            Dim fld As Outlook.Folder
            Dim oAccount As Outlook.Account
            Dim att As Outlook.Attachment
            Dim atts As Outlook.Attachments
            Dim arquivoImagem As String
            Dim strPath As String
            Dim imagesPath As String
         
         On Error GoTo error_handler
         
         strPath = Application.GetOpenFilename(, , "Selecione o arquivo PPT", "Selecione")
         If InStr(1, UCase(strPath), ":") = 0 Then Exit Sub
         
        
               'abre CADASTRO DE REGIONAIS
               Set wbContatos = Workbooks.Open("C:\Dropbox\VEDACIT_DADOS\Contatos.xlsx") 'pega o arquivo q contem a planilha a ser copiada
               Set cadastro = wbContatos.Sheets("Regionais")
               
            
            'abre o arquivo
                           
            Set pptApp = New PowerPoint.Application
            Set pptPresentation = pptApp.Presentations.Open(strPath)
            
            For Each sl In pptPresentation.Slides
                    
                    sl.Select
                    txtNota = sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange
                    stag = "email"
                    tagini = InStrRev(txtNota, "<" & stag & ">") + Len(stag) + 2 ' position of start delimiter
                    tagfim = InStr(txtNota, "</" & stag & ">") ' position of end delimiter
                    endereco = Mid(txtNota, tagini, tagfim - tagini)
                    
                    stag = "regional"
                    tagini = InStrRev(txtNota, "<" & stag & ">") + Len(stag) + 2 ' position of start delimiter
                    tagfim = InStr(txtNota, "</" & stag & ">") ' position of end delimiter
                    aRegional = Mid(txtNota, tagini, tagfim - tagini)
                    mailRegional = cadastro.Columns(1).Find(aRegional, lookat:=xlWhole).Offset(0, 2)
                    nomeRepresentante = sl.Shapes(2).TextFrame.TextRange.Text
                    'nomeRepresentante = cleanString(CStr(nomeRepresentante))
                    
                    sl.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange = Replace(txtNota, aRegional, mailRegional)
                   
            Next sl

finaliza:
wbContatos.Close False
pptPresentation.Save
pptPresentation.Close

Set pptApp = Nothing


Exit Sub
error_handler:
eee = Err.Description
Err.Clear
Resume Next
GoTo finaliza


       
End Sub










