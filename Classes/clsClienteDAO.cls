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
    Dim clsconexao       As clsconexao

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
    strSQL = strSQL & "             ) VALUE  (  " & vCliente.Codigo
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

    Set clsconexao = New clsconexao
    clsconexao.ProvedorTipo = OleDb
    clsconexao.ConexaoString = CONEXAO_ACESSO
    Call clsconexao.Abrir
    Call clsconexao.ExecuteNonQuery(strSQL)
    Call clsconexao.Fechar
    Incluir = True
    Set clsconexao = Nothing
End Function
