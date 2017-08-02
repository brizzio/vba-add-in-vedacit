Attribute VB_Name = "VARIABLES"
Public addin As Workbook
Public dbConnection As ADODB.Connection

Public Const fileExt As String = ".xlsx"

Public Const conta_vedateam As String = "vedateam@outlook.com"

Public Const ARQUIVO_PARAMETROS As String = "C:\Dropbox\VEDACIT_DADOS\PARAMETROS" & fileExt
Public Const ARQUIVO_DADOS As String = "C:\Dropbox\VEDACIT\Resultados Fev_17" & fileExt
Public Const ARQUIVO_METAS As String = "C:\Dropbox\VEDACIT\ORCAMENTO2017.xlsm"



Public dePara As Worksheet
Public param As Worksheet
Public dados As Worksheet
Public metas As Worksheet


Public Enum listagemDe
    
    region = 0
    repres = 1
    cidades = 2
    clientes = 3

End Enum

Public Enum tipo

    Representante
    regional

End Enum

Public Enum colunaParaRetornarDoRepresentanteParaPontuarNaRegional

    colunaComPontosDeFaturamento
    colunaComPontosDeClientesAtivos
    colunaComPontosDeMix
    
End Enum




Public Type SETS

    set1 As String ' = "1,2,3"
    set2 As String '= "4,5,6"
    set3 As String '= "7,8,9"
    set4 As String '= "10,11,12"

End Type


Public w As Workbook
Public s As Worksheet
Public CABECALHO As Range
Public rr As Range

'Public Const ARQUIVO_DADOS As String = "C:\Dropbox\VEDACIT_DADOS\DATABASE.XLSX"
Public Const PLANILHA_DE_DADOS As String = "base 2016"

Public Const ACCESS_DATABASE = "C:\Dropbox\VEDACIT_DADOS\dbVedacit.mdb"
Public Const tabelaBienio As String = "20152016"

    Public commentFat As String
    Public commentCat As String
    Public commentMix As String
    Public commentCap As String
    Public commentRen As String
    
    Public mesExtenso As String

Public Const cabFat As String = "FATURAMENTO"
Public Const cabCat As String = "CLIENTES ATIVOS"
Public Const cabMix As String = "MIX PRODUTO"
Public Const cabCap As String = "CAPILARIDADE"
Public Const cabRen As String = "RENTABILIDADE"

Public Const cabRegFat As String = "FATURA MENTO"
Public Const cabRegCap As String = "CAPILA RIDADE"
Public Const cabRegRen As String = "RENTABI LIDADE"
Public Const cabRegMix As String = "MIX REGIONAL"
Public Const cabRegFatRep As String = "FATURA MENTO REPRESEN TANTES"
Public Const cabRegCatRep As String = "CLIENTES ATIVOS REPRESEN TANTES"
Public Const cabRegMixRep As String = "MIX REPRESEN TANTES"

Public Const PLANILHA_MASTER_REPRESENTANTES = "MASTER2"
Public Const PLANILHA_MASTER_REGIONAIS = "REGIONAIS_PONTUACAO"
Public Const meses As String = "JANEIRO,FEVEREIRO,MARCO,ABRIL,MAIO,JUNHO,JULHO,AGOSTO,SETEMBRO,OUTUBRO,NOVEMBRO,DEZEMBRO"



Public Sub CONECTA()

   ' Set dbConnection = New ADODB.connection
    
   ' dbConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ACCESS_DATABASE & ";User Id=admin;Password=;"
                

End Sub


Public Function removeDupesInArray(ByRef myArray()) As Variant
    Dim StrtTime As Double, Endtime As Double
    Dim invalidKey As String
    Dim d As Scripting.Dictionary, i As Long  '' Early Binding
    
    invalidKey = "-"
    
    Set d = New Scripting.Dictionary
    For i = LBound(myArray) To UBound(myArray): d(myArray(i)) = 1: Next i
    
    If d.Exists(invalidKey) Then d.Remove invalidKey
    
    removeDupesInArray = d.Keys()
End Function
