Attribute VB_Name = "SENDMAIL"


Public Sub CreateNewMessage()
Dim objMsg As Outlook.MailItem
Dim oApp As Outlook.Application
Dim ns As Outlook.Namespace
Dim fld As Outlook.Folder

Dim s As Worksheet

Set s = ActiveSheet

Set oApp = New Outlook.Application
Set ns = oApp.GetNamespace("MAPI")

Set fld = ns.GetDefaultFolder(olFolderOutbox)

Set objMsg = fld.Items.Add(olMailItem)

 With objMsg
  .To = "Alias@domain.com"
  .CC = "Alias2@domain.com"
  .BCC = "Alias3@domain.com"
  .Subject = "This is the subject"
  .Categories = "Teste caixa"
  .VotingOptions = "Yes;No;Maybe;"
  .BodyFormat = olFormatHTML ' send plain text message
  .HTMLBody = RangetoHTML(s.UsedRange)
  .HTMLBody = .HTMLBody & "<br><br>======================== FAC SIMILE DO EMAIL ENVIADO =====================================<br><br>"
  .HTMLBody = .HTMLBody & MAKE_HTML_BODY("XXXXXXXXX", "XXXXXXXXX", "XXXXXXXXX", "XXXXXXXXX", "XXXXXXXXX", "XXXXXXXXX", "XXXXXXXXX", "XXXXXXXXX")
  .Importance = olImportanceHigh
  .Sensitivity = olConfidential
 ' .Attachments.Add ("path-to-file.docx")

' Calculate a date using DateAdd or enter an explicit date
  .ExpiryTime = DateAdd("m", 6, Now) '6 months from now
  .DeferredDeliveryTime = #8/1/2012 6:00:00 PM#
  
  .Display
  '.Save
  
'  .DeleteAfterSubmit = True
End With

Set objMsg = Nothing
End Sub



Public Sub SEND_TEXT_FILE_BY_MAIL(fullPathFile As String, strsubject As String, strbody As String)
    
    Dim iMsg As Object
    Dim iConf As Object
    
    Dim Flds As Variant
    'Dim mm As New CDO.Message

    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")

    iConf.Load -1    ' CDO Source Defaults
    Set Flds = iConf.Fields
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "semesp.hbp@gmail.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "semesp@123"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"

        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Update
    End With
    
    'strbody = "Cara ANA, " & vbNewLine & vbNewLine & _
    '          "segue o arquivo texto" & vbNewLine & _
    '          "This is line 2" & vbNewLine & _
    '          "This is line 3" & vbNewLine & _
    '          "This is line 4"

    With iMsg
        Set .Configuration = iConf
        .To = "semesp@bpsp.org.br"
        .CC = ""
        .BCC = ""
        ' Note: The reply address is not working if you use this Gmail example
        ' It will use your Gmail address automatic. But you can add this line
        ' to change the reply address  .ReplyTo = "Reply@something.nl"
        .From = """SISTEMA SISE"" <semesp.hbp@gmail.com>"
        .Subject = strsubject
        .TextBody = strbody
        
        .AddAttachment fullPathFile
        '.Send
    End With

End Sub

Public Sub SEND_FATURAMENTO_REPORT(WS As Worksheet)
    
    Dim iMsg As Object
    Dim iConf As Object
    Dim diaFat As String
    Dim Flds As Variant
    'Dim mm As New CDO.Message

    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")

    iConf.Load -1    ' CDO Source Defaults
    Set Flds = iConf.Fields
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "semesp.hbp@gmail.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "semesp@123"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"

        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Update
    End With
    

    
    'separa a data para poder enviar email
    diaFat = VBA.Replace(ActiveWorkbook.Name, "ACSLAC", "")
    diaFat = VBA.Replace(UCase(diaFat), ".XLS", "")
    
    With iMsg
        Set .Configuration = iConf
        .To = "joaoalmeidajr@uol.com.br"
        .CC = "jotaflaquer@hotmail.com"
        .BCC = "cetsemesp@terra.com.br"
        ' Note: The reply address is not working if you use this Gmail example
        ' It will use your Gmail address automatic. But you can add this line
        ' to change the reply address  .ReplyTo = "Reply@something.nl"
        .From = """SISTEMA SISE"" <semesp.hbp@gmail.com>"
        .Subject = "FATURAMENTO DO DIA (PERIODO): " & diaFat
        .HTMLBody = RangetoHTML(WS.UsedRange)
        
        '.AddAttachment fullPathFile
        '.Send
    End With

End Sub




'If you have a GMail account then you can try this example to use the GMail smtp server
'The example will send a small text message
'You must change four code lines before you can test the code

'.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "Full GMail mail address"
'.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "GMail password"

'Use your own mail address to test the code in this line
'.To = "Mail address receiver"

'Change YourName to the From name you want to use
'.From = """YourName"" <Reply@something.nl>"

'If you get this error : The transport failed to connect to the server
'then try to change the SMTP port from 25 to 465

'Possible that you must also enable the "Less Secure" option for GMail
'https://www.google.com/settings/security/lesssecureapps


Sub SEND_MAIL_USING_OUTLOOK()
    Dim iMsg As Object
    Dim iConf As Object
    Dim strbody As String
    Dim Flds As Variant
    Dim mm As New CDO.Message

    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")

    iConf.Load -1    ' CDO Source Defaults
    Set Flds = iConf.Fields
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "vedateam@outlook.com"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "@204720!"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp-mail.outlook.com"

        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Update
    End With
    
    strbody = "Hi there" & vbNewLine & vbNewLine & _
              "This is line 1" & vbNewLine & _
              "This is line 2" & vbNewLine & _
              "This is line 3" & vbNewLine & _
              "This is line 4"

    With iMsg
        Set .Configuration = iConf
        .To = "fabrizio@altcom.net.br"
        .CC = ""
        .BCC = ""
        ' Note: The reply address is not working if you use this Gmail example
        ' It will use your Gmail address automatic. But you can add this line
        ' to change the reply address  .ReplyTo = "Reply@something.nl"
        .From = "vedateam@outlook.com"
        .Subject = "Important message"
        .TextBody = strbody
        .Send
    End With

End Sub





Sub SEND_MAIL()
    Dim iMsg As Object
    Dim iConf As Object
    Dim strbody As String
    Dim Flds As Variant
    Dim mm As New CDO.Message

    Set iMsg = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")

    iConf.Load -1    ' CDO Source Defaults
    Set Flds = iConf.Fields
    With Flds
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = True
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
        .Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "fabrizio@altcom.net.br"
        .Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "80sao20ver"
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"

        .Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
        .Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
        .Update
    End With
    
    strbody = "Hi there" & vbNewLine & vbNewLine & _
              "This is line 1" & vbNewLine & _
              "This is line 2" & vbNewLine & _
              "This is line 3" & vbNewLine & _
              "This is line 4"

    With iMsg
        Set .Configuration = iConf
        .To = "fabrizio.salvade@gmail.com"
        .CC = ""
        .BCC = ""
        ' Note: The reply address is not working if you use this Gmail example
        ' It will use your Gmail address automatic. But you can add this line
        ' to change the reply address  .ReplyTo = "Reply@something.nl"
        .From = "fabrizio@altcom.net.br"
        .Subject = "Important message"
        .TextBody = strbody
        .Send
    End With

End Sub


Public Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2007
    Dim FSO As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "/" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         fileName:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set ts = FSO.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.ReadAll
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close SaveChanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set FSO = Nothing
    Set TempWB = Nothing
End Function


