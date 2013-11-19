VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShippersDAO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Public Function Incluir(ByVal vShippers As clsShippersDTO) As Boolean

    Dim strSQL           As String
    Dim clsConexao       As clsConexao

    strSQL = "          INSERT INTO shippers ( ShipperID "
    strSQL = strSQL & "                      , CompanyName "
    strSQL = strSQL & "                      , Phone "
    strSQL = strSQL & "              ) VALUE (  " & vShippers.ShipperID
    strSQL = strSQL & "                      , '" & vShippers.CompanyName & "'"
    strSQL = strSQL & "                      , '" & vShippers.Phone & "'"
    strSQL = strSQL & "                      ) "

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

        Dim strSQL As String
        Dim clsConexao       As clsConexao
        
        strSQL = "DELETE FROM shippers WHERER ShipperID = " & vCodigo
        
        clsConexao.ProvedorTipo = OleDb
    clsConexao.ConexaoString = CONEXAO_ACESSO

    Call clsConexao.Abrir

    Excluir = clsConexao.ExecuteNonQuery(strSQL)

    Call clsConexao.Fechar

    Set clsConexao = Nothing

End Function


Public Function Alterar(ByVal vShippers As clsShippersDTO) As Boolean

    Dim strSQL           As String
    Dim clsConexao       As clsConexao
    
    clsConexao.ProvedorTipo = OleDb
    clsConexao.ConexaoString = CONEXAO_ACESSO
    
    
    strSQL = "         UPDATE shippers "
    strSQL = strSQL & "   SET CompanyName = '" & vShippers.CompanyName & "'"
    strSQL = strSQL & "     , Phone = '" & vShippers.Phone & "'"
    strSQL = strSQL & " WHERE ShipperID = " & vShippers.ShipperID

    Call clsConexao.Abrir
    Call clsConexao.ExecuteNonQuery(strSQL)
    Call clsConexao.Fechar
    
    Alterar = True

    Set clsConexao = Nothing

End Function