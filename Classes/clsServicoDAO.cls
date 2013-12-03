VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServicoDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Incluir(ByVal vServico As clsServicoDTO) As Boolean

    Dim strSQL           As String
    Dim clsConexao       As clsConexao

    strSQL = "                                 INSERT INTO servico ( codigo "
    strSQL = strSQL & "                                            , nome"
    strSQL = strSQL & "                                            , valor"
    strSQL = strSQL & "                                ) VALUE     (  " & vServico.Codigo
    strSQL = strSQL & "                                            , '" & vServico.Nome & "'"
    strSQL = strSQL & "                                            , " & vServico.Valor & ""
    strSQL = strSQL & "                                            )"

    Set clsConexao = New clsConexao
    
    clsConexao.ProvedorTipo = OleDb
    clsConexao.ConexaoString = CONEXAO_ACESSO
    
    Call clsConexao.Abrir
    Call clsConexao.ExecuteNonQuery(strSQL)
    Call clsConexao.Fechar
    
    Incluir = True
    
    Set clsConexao = Nothing
    
End Function


Public Function Excluir(ByVal vCodigo As Long) As Long
    Dim strSQL           As String
    Dim clsConexao       As clsConexao

    strSQL = "DELETE FROM servico where codigo = " & vCodigo
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
    
    
    strSQL = "         UPDATE servico "
    strSQL = strSQL & "   SET nome = '" & vServico.Nome & "'"
    strSQL = strSQL & "     , valor = '" & vServico.Valor & "'"
    strSQL = strSQL & " WHERE codigo = " & vServico.Codigo

    Call clsConexao.Abrir
    Call clsConexao.ExecuteNonQuery(strSQL)
    Call clsConexao.Fechar
    
    Alterar = True

    Set clsConexao = Nothing

End Function
