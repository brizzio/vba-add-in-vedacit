Attribute VB_Name = "AUX_COMANDOS_DE_TECLADO"

Public Sub SHORTCUT_ACTIVATE()
 Application.OnKey "^e", "CALCULA_PONTUACAO_GERAL"
 Application.OnKey "^h", "IMPORTA_TABELA_HISTORICO"
 Application.OnKey "^t", "INFORMA_REPRESENTANTES"
 'Application.OnKey "^p", ""
 'Application.OnKey "^k", ""
End Sub

Public Sub SHORTCUT_DEACTIVATE()
 Application.OnKey "^e", ""
 Application.OnKey "^h", ""
 Application.OnKey "^t", ""
 'Application.OnKey "^p", ""
 'Application.OnKey "^k", ""
End Sub

'=======================================================
'deleta forms criados pelo sistema e que são dispensáveis
'========================================================
'if you want to do this by code, set a reference to the VBA 5.3 Extensibility
'Library (in VBA, go to the Tools menu, choose References, and check
'"Microsoft Visual Basic For Applications Extensibility Library 5.3"). Then
'use code like
Sub DeleteForm()
    Dim VBComps As VBIDE.VBComponents
    Dim VBComp As VBIDE.VBComponent
    Set VBComps = ThisWorkbook.VBProject.VBComponents
    
    For Each VBComp In VBComps
        
        If InStr(1, VBComp.Name, "UserForm") > 0 Then VBComps.Remove VBComp
    
    Next
    
    'Set VBComp = VBComps("UserForm1")
    
    '    VBComps.Remove VBComp
End Sub


Public Function FileNameFromPath(strFullPath As String) As String

      FileNameFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "\"))

End Function

 Public Function FileNameNoExtensionFromPath(strFullPath As String) As String

     Dim intStartLoc As Integer
     Dim intEndLoc As Integer
     Dim intLength As Integer

     intStartLoc = Len(strFullPath) - (Len(strFullPath) - InStrRev(strFullPath, "\") - 1)
     intEndLoc = Len(strFullPath) - (Len(strFullPath) - InStrRev(strFullPath, "."))
     intLength = intEndLoc - intStartLoc

     FileNameNoExtensionFromPath = Mid(strFullPath, intStartLoc, intLength)

 End Function

 Public Function FolderFromPath(ByRef strFullPath As String) As String

      FolderFromPath = Left(strFullPath, InStrRev(strFullPath, "\"))

  End Function

 Public Function FileExtensionFromPath(ByRef strFullPath As String) As String

      FileExtensionFromPath = Right(strFullPath, Len(strFullPath) - InStrRev(strFullPath, "."))

 End Function


Public Sub Delete_Folder(caminhoDaPasta As String)
'Delete whole folder without removing the files first like in DeleteExample4
    Dim FSO As Object
   
    Set FSO = CreateObject("scripting.filesystemobject")

    If Right(caminhoDaPasta, 1) = "\" Then
        caminhoDaPasta = Left(caminhoDaPasta, Len(caminhoDaPasta) - 1)
    End If

    If FSO.FolderExists(caminhoDaPasta) = False Then
        Debug.Print caminhoDaPasta & " não existe"
        Exit Sub
    End If

    FSO.DeleteFolder caminhoDaPasta

End Sub

Public Function cleanString(strng As String) As String
        
        Dim strPattern As String: strPattern = "[^a-zA-Z0-9]" 'The regex pattern to find special characters
        Dim strReplace As String: strReplace = "" 'The replacement for the special characters
        Set regEx = CreateObject("vbscript.regexp") 'Initialize the regex object
        
        
        ' Configure the regex object
        With regEx
            .Global = True
            .MultiLine = True
            .IgnoreCase = False
            .Pattern = strPattern
        End With
        
        ' Perform the regex replacement
        cleanString = regEx.Replace(strng, strReplace)


End Function

Public Sub DeleteFilteredOutRows(Optional asheet As Worksheet)
'
' DeleteFilteredOutRows Macro
'
Dim x As Integer, HelperC As Integer, LastRow As Integer

If asheet Is Nothing Then Set asheet = ActiveSheet

'Find LastRow
asheet.Range("A1").Select
LastRow = asheet.Cells(Rows.Count, 1).End(xlUp).Row

'Add Helper Column to identify if visible
asheet.Range("A1").Select
Selection.End(xlToRight).Select
ActiveCell.Offset(0, 1).Select
HelperC = ActiveCell.Column ' HelperC = Column number of helper column
ActiveCell.Value = "Visible?"

'If visible, add 1 to Visible column
For x = 2 To LastRow
If asheet.Rows(x).EntireRow.Hidden Then
Else
asheet.Cells(x, HelperC).Value = 1
End If
Next x

Set delRange = Nothing
'If not visible(Visible column <> 1) then delete row
asheet.Outline.ShowLevels RowLevels:=8, ColumnLevels:=8
LastRow = asheet.UsedRange.Rows.Count
For x = 2 To LastRow
If asheet.Cells(x, HelperC).Value <> 1 Then

If delRange Is Nothing Then Set delRange = asheet.Cells(x, HelperC) Else Set delRange = Union(delRange, asheet.Cells(x, HelperC))
End If
Next x

delRange.Select
delRange.EntireRow.Delete xlUp

asheet.Columns(HelperC).EntireColumn.Delete 'Delete Helper Column
asheet.Range("A1").Select ' Select cell A1



' Removes filters
On Error GoTo finaliza
asheet.ShowAllData

finaliza:

End Sub


