VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOSDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Incluir(ByVal vOS As clsOsDTO) As Boolean
    Dim srtSQL           As String
    Dim clsConexao       As clsConexao

    strSQL = "               INSERT INTO os( codigo"
    strSQL = "                             , data"
    strSQL = "                             , dataentrega"
    strSQL = "                             , codcliente"
    strSQL = "                             , codservico"
    strSQL = "                             , codinstrumento"
    strSQL = "                   VALUE)    ( " & vOS.Codigo
    strSQL = strSQL & "                    , " & vOS.Data & ""
    strSQL = strSQL & "                    , " & vOS.DataEntrega & ""
    strSQL = strSQL & "                    , " & vOS.CodCLiente & ""
    strSQL = strSQL & "                    , " & vOS.CodServico & ""
    strSQL = strSQL & "                    , " & vOS.CodInstrumento & ""
    strSQL = strSQL & "                    )"

    Set clsConexao = New clsConexao
    clsConexao.ProvedorTipo = OleDb
    clsConexao.ConexaoString = CONEXAO_ACESSO
    Call clsConexao.Abrir
    Call clsConexao.ExecuteNonQuery(strSQL)
End Function

Public Function Excluir(ByVal vCodigo As Long) As Long

    Dim strSQL           As String
    Dim clsConexao       As clsConexao

    strSQL = "DELETE FROM os WHERER codigo = " & vCodigo
    clsConexao.ProvedorTipo = OleDb
    clsConexao.ConexaoString = CONEXAO_ACESSO
    Call clsConexao.Abrir
    Excluir = clsConexao.ExecuteNonQuery(strSQL)
    Call clsConexao.Fechar
    Set clsConexao = Nothing

End Function
