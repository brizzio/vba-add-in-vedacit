Attribute VB_Name = "SHEET_FUNCTIONS"
Public Function SEPARA_STRING(txt As Range, delimitador As String, Optional posicao As Integer = 1) As String
    
    Dim r As Range
    Dim arrayDeSubStrings
    On Error GoTo function_error
    Err.Clear
    'Set r = Application.Caller
    'aaa = r.Address
    'arrayDeSubStrings = Split(txt.Text, delimitador)
    
     '   SEPARA_STRING = arrayDeSubStrings(posicao - 1)
 Err.Clear
Exit Function
function_error:
    Err.Clear
    SEPARA_STRING = vbNullString
    
End Function
