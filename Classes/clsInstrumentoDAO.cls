VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInstrumentoDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Incluir(ByVal vInstrumento As clsInstrumentoDTO) As Boolean
    Dim strSQL           As String
    Dim clsConexao       As clsConexao

    strSQL = "               INSERT INTO instrumento ( codigo "
    strSQL = strSQL & "                              , nome"
    strSQL = strSQL & "                              , modelo"
    strSQL = strSQL & "                              , marca"
    strSQL = strSQL & "                  ) VALUE     (  " & vInstrumento.Codigo
    strSQL = strSQL & "                              , '" & vInstrumento.Nome & "'"
    strSQL = strSQL & "                              , '" & vInstrumento.Modelo & "'"
    strSQL = strSQL & "                              , '" & vInstrumento.Marca & "'"
    strSQL = strSQL & "                              )"

    Set clsConexao = New clsConexao
    clsConexao.ProvedorTipo = OleDb
    clsConexao.ConexaoString = CONEXAO_ACESSO
    Call clsConexao.Abrir
    Call clsConexao.ExecuteNonQuery(strSQL)
End Function

Public Function Excluir(ByVal vCodigo As Long) As Long

    Dim strSQL           As String
    Dim clsConexao       As clsConexao

    strSQL = "DELETE FROM instrumento WHERER codigo = " & vCodigo
    clsConexao.ProvedorTipo = OleDb
    clsConexao.ConexaoString = CONEXAO_ACESSO
    
    Call clsConexao.Abrir
    Excluir = clsConexao.ExecuteNonQuery(strSQL)
    
    Call clsConexao.Fechar
    Set clsConexao = Nothing

End Function
Public Function Alterar(ByVal vCodigo As Long) As Boolean

    Dim strSQL           As String
    Dim clsConexao       As clsConexao
    
    clsConexao.ProvedorTipo = OleDb
    clsConexao.ConexaoString = CONEXAO_ACESSO
    
    
    strSQL = "         UPDATE instrumento "
    strSQL = strSQL & "   SET nome = '" & vInstrumento.Nome & "'"
    strSQL = strSQL & "     , modelo = '" & vInstrumento.Modelo & "'"
    strSQL = strSQL & "     , marca = '" & vInstrumento.Marca & "'"
    strSQL = strSQL & " WHERE codigo = " & vInstrumento.Codigo

    Call clsConexao.Abrir
    Call clsConexao.ExecuteNonQuery(strSQL)
    Call clsConexao.Fechar
    
    Alterar = True

    Set clsConexao = Nothing

End Function
