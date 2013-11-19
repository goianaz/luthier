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
    Dim srtSQL           As String
    Dim clsconexao       As clsconexao

    strSQL = "                                 INSERT INTO servico ( codigo "
    strSQL = strSQL & "                                            , nome"
    strSQL = strSQL & "                                            , valor"
    strSQL = strSQL & "                                ) VALUE     (  " & vServico.Codigo
    strSQL = strSQL & "                                            , '" & vServico.Nome & "'"
    strSQL = strSQL & "                                            , " & vServico.Valor & ""
    strSQL = strSQL & "                                            )"

    Set clsconexao = New clsconexao
    clsconexao.ProvedorTipo = OleDb
    clsconexao.ConexaoString = CONEXAO_ACESSO
    Call clsconexao.Abrir
    Call clsconexao.ExecuteNonQuery(strSQL)
End Function

