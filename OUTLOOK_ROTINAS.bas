Attribute VB_Name = "OUTLOOK_ROTINAS"
Sub analizaRange()
    
    Dim oApp As Outlook.Application
    Dim ns As Outlook.Namespace
    Dim fld As Outlook.Folder
    Dim oCnt As Outlook.ContactItem
    Dim myRecipients As Outlook.Recipients
    
    Dim myDistList As Outlook.DistListItem

    Dim r As Range
    Dim lista As String
    Dim str As String
    
    
    Set oApp = New Outlook.Application
    Set myDistList = oApp.CreateItem(olDistributionListItem)
    
    Set myTempItem = oApp.CreateItem(olMailItem)
    Set myRecipients = myTempItem.Recipients
    
    'myDistList.DLName = "Representantes"
    
    Set r = Selection
    
    mail = Split(r.Text, ";")
    
    For i = 0 To UBound(mail)
        
        str = mail(i)
        strNome = Trim(Split(str, "<")(0))
        strMail = Replace(Split(str, "<")(1), ">", "")
        
        Set oCnt = oApp.CreateItem(olContactItem)
        oCnt.FullName = strNome
        oCnt.Email1Address = strMail
        oCnt.Department = "Matriz"
        oCnt.JobTitle = "Diretor"
        
        myRecipients.Add strMail
       
        oCnt.Save
        
    Next i
    
    myRecipients.ResolveAll
    myDistList.AddMembers myRecipients
    myDistList.Save



End Sub


Sub montaListaDistribuicaoRepresentantes()
    
    Dim oApp As Outlook.Application
    Dim ns As Outlook.Namespace
    Dim fld As Outlook.Folder
    Dim oCnt As Outlook.ContactItem
    Dim myRecipients As Outlook.Recipients
    
    Dim myDistList As Outlook.DistListItem
    
    Dim s As Worksheet
    Dim r As Range
    Dim lista As String
    Dim str As String
    
    Set s = ActiveSheet
    
    Set oApp = New Outlook.Application
    Set myDistList = oApp.CreateItem(olDistributionListItem)
    
    Set myTempItem = oApp.CreateItem(olMailItem)
    Set myRecipients = myTempItem.Recipients
    
    myDistList.DLName = "Representantes"
    
    
    
    For i = 2 To s.UsedRange.Rows.Count
        
        Set r = s.Cells(i, 1)
        
        
        
        Set oCnt = oApp.CreateItem(olContactItem)
        oCnt.FullName = r.Offset(0, 2).Text
        oCnt.CompanyName = r.Offset(0, 1).Text
        oCnt.NickName = r.Text
        oCnt.Email1Address = Split(r.Offset(0, 4).Text, "/")(0)
        If UBound(Split(r.Offset(0, 4).Text, "/")) > 0 Then oCnt.Email2Address = Split(r.Offset(0, 4).Text, "/")(1)
        oCnt.MobileTelephoneNumber = r.Offset(0, 5).Text
        oCnt.BusinessAddressState = r.Offset(0, 3).Text
        
        oCnt.Department = "Comercial"
        oCnt.JobTitle = "Representante"
        
        myRecipients.Add Replace(r.Offset(0, 4).Text, "/", ";")
       
        oCnt.Save
        
    Next i
    
    myRecipients.ResolveAll
    myDistList.AddMembers myRecipients
    myDistList.Save



End Sub





