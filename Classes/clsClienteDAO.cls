VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClienteDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Incluir(ByVal vCliente As clsClienteDTO) As Boolean

    Dim strSQL           As String
    Dim clsConexao       As clsConexao

    strSQL = "           INSERT INTO cliente ( codigo "
    strSQL = strSQL & "                      , nome "
    strSQL = strSQL & "                      , endereco"
    strSQL = strSQL & "                      , nro"
    strSQL = strSQL & "                      , cidade"
    strSQL = strSQL & "                      , bairro"
    strSQL = strSQL & "                      , uf"
    strSQL = strSQL & "                      , fonecel"
    strSQL = strSQL & "                      , fonefixo"
    strSQL = strSQL & "                      , email "
    strSQL = strSQL & "             ) VALUE  (  " & vCliente.codigo
    strSQL = strSQL & "                      , '" & vCliente.Nome & "'"
    strSQL = strSQL & "                      , '" & vCliente.Endereco & "'"
    strSQL = strSQL & "                      , '" & vCliente.Nro & "'"
    strSQL = strSQL & "                      , '" & vCliente.Cidade & "'"
    strSQL = strSQL & "                      , '" & vCliente.Bairro & "'"
    strSQL = strSQL & "                      , '" & vCliente.Uf & "'"
    strSQL = strSQL & "                      , '" & vCliente.Fonecel & "'"
    strSQL = strSQL & "                      , '" & vCliente.Fonefixo & "'"
    strSQL = strSQL & "                      , '" & vCliente.Email & "'"
    strSQL = strSQL & "                      )"

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

    strSQL = "DELETE FROM cliente WHERER codigo = " & vCodigo
    clsConexao.ProvedorTipo = OleDb
    clsConexao.ConexaoString = CONEXAO_ACESSO
    
    Call clsConexao.Abrir
    Excluir = clsConexao.ExecuteNonQuery(strSQL)
    Call clsConexao.Fechar
    
    Set clsConexao = Nothing

End Function
Public Function Alterar(ByVal vCliente As clsClienteDTO) As Boolean

    Dim strSQL           As String
    Dim clsConexao       As clsConexao
    
    clsConexao.ProvedorTipo = OleDb
    clsConexao.ConexaoString = CONEXAO_ACESSO
    
    
    strSQL = "         UPDATE cliente "
    strSQL = strSQL & "   SET nome = '" & vCliente.Nome & "'"
    strSQL = strSQL & "                      , '" & vCliente.Endereco & "'"
    strSQL = strSQL & "                      , '" & vCliente.Nro & "'"
    strSQL = strSQL & "                      , '" & vCliente.Cidade & "'"
    strSQL = strSQL & "                      , '" & vCliente.Bairro & "'"
    strSQL = strSQL & "                      , '" & vCliente.Uf & "'"
    strSQL = strSQL & "                      , '" & vCliente.Fonecel & "'"
    strSQL = strSQL & "                      , '" & vCliente.Fonefixo & "'"
    strSQL = strSQL & "                      , '" & vCliente.Email & "'"
    strSQL = strSQL & " WHERE codigo = " & vCliente.codigo

    Call clsConexao.Abrir
    Call clsConexao.ExecuteNonQuery(strSQL)
    Call clsConexao.Fechar
    
    Alterar = True

    Set clsConexao = Nothing

End Function


