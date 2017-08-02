Attribute VB_Name = "RENTABILIDADE"
Public Sub Rentabilidade()

    
    Dim oConn As New ADODB.Connection
    Dim rs As ADODB.Recordset
    Dim query As String
    Dim fileName As String
    Dim cnStr As String
    
    On Error GoTo errHandler
    
    fileName = "C:\Dropbox\VEDACIT\RENTABILIDADE_METAS_E_HISTORICO.xlsm"
    '-----------------------------------------------------------------------------------
    'I ********references are set to:********
    'I * Visual Basic For Applications
    'I * Microsoft Excel 12.0 ObjectLibrary
    'I * Microsoft ADO Ext. 6.0 for DDL and Security
    'I * Microsoft ActiveX Data Objects 6.1 Library
    'I * Microsoft AcitveX Data Objects Recordset 6.0 Library
    '-----------------------------------------------------------------------------------
    'query = "SELECT * FROM [Sheet1$D1:D15]"
    
    query = "SELECT * FROM [DADOS$]"
    
    oConn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
               "Data Source=" & fileName & ";" & _
               "Extended Properties=""Excel 8.0;"""
    
    rs.Open query, cnStr, adOpenStatic, adLockReadOnly

    'Set rs = New ADODB.Recordset
    rs.Open query, cnStr, adOpenStatic, adLockReadOnly

    'Begin row processing
        Do While Not rs.EOF
            
            'Retrieve the name of the first city in the selected rows
            strCity = rst!City
        
            'Now filter the Recordset to return only the customers from that city
            rst.Filter = "City = '" & strCity & "'"
            Set rstFiltered = rst.OpenRecordset
        
            'Process the rows
            Do While Not rstFiltered.EOF
                rstFiltered.Edit
                rstFiltered!ToBeVisited = True
                rstFiltered.Update
                rstFiltered.MoveNext
            Loop
        
            'We've done what was needed. Now exit
            Exit Do
            rst.MoveNext
           
        Loop
    

    
    
    ' Copy the entire result table to the Destination
            If Not rdest Is Nothing Then
                 If insereCabecalho Then
                            For lThisField = 0 To oRs.Fields.Count - 1
                                rdest.Offset(0, lThisField) = oRs.Fields(lThisField).Name
                            Next
                            Set rdest = rdest.Offset(1, 0)
                  End If
                  rdest.CopyFromRecordset oRs
            Else
                    If Not (oRs.BOF And oRs.EOF) Then
                        SqlFilter = oRs.GetRows()
                    Else
                        SqlFilter = ""
                    End If
            End If

ExitRoutine:
                If Not oConn Is Nothing Then
                    If oConn.State <> adStateClosed Then
                        oConn.Close
                    End If
                    Set oConn = Nothing
                End If
                If Not oRs Is Nothing Then
                    If oRs.State <> adStateClosed Then
                        oRs.Close
                    End If
                    Set oRs = Nothing
                End If
                Exit Sub
errHandler:
               
                    Debug.Print "VBA Error #" & Err.Number & ": " & Err.Description
                    Err.Clear
               ' End If
                
                Stop
                Resume ExitRoutine


End Sub


