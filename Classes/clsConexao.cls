VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConexao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum Provedor
    ConnectionString = 1
    OleDb = 2
    Odbc = 3
End Enum

Private m_blnTransacao   As Boolean
Private m_strErro As String
Private m_strConexaoString As String
Private m_enuProvedorTipo As Provedor
Private m_objConexao     As ADODB.Connection
Private m_objComando     As ADODB.Command

Public Property Get Erro() As String
    Erro = m_strErro
End Property

Public Property Let ConexaoString(ByVal strConexaoString As String)
    m_strConexaoString = strConexaoString
End Property

Public Property Let ProvedorTipo(ByVal enuProvedorTipo As Provedor)
    m_enuProvedorTipo = enuProvedorTipo
End Property

Private Function ObterConexao(ByVal vProvedorTipo As Provedor) As ADODB.Connection

    Dim strConexao       As String

    Select Case vProvedorTipo

        Case ConnectionString:

            Set ObterConexao = New ADODB.Connection

        Case OleDb, Odbc:
            Set ObterConexao = New ADODB.Connection

        Case Else:
            Err.Raise "Provedor de acesso inválido."
    End Select

End Function

Private Function ObterComando(ByVal vProvedorTipo As Provedor) As ADODB.Command
    
    Select Case vProvedorTipo
    
    Case ConnectionString:
    
        Set ObterComando = New ADODB.Command
        
    Case OleDb, Odbc:
        
        Set ObterComando = New ADODB.Command
        
    End Select
    
End Function

Public Sub Abrir()

    Set m_objConexao = ObterConexao(m_enuProvedorTipo)

    Select Case m_enuProvedorTipo

        Case ConnectionString:

            m_objConexao.ConnectionString = m_strConexaoString

            If (m_objConexao.State <> adStateOpen) Then
                m_objConexao.Open
            End If

        Case OleDb, Odbc:

            If (m_objConexao.State <> adStateOpen) Then
                m_objConexao.Open m_strConexaoString
            End If

        Case Else

            Err.Raise "Provedor de acesso inválido."

    End Select

    Set m_objComando = ObterComando(m_enuProvedorTipo)

End Sub


Public Sub Fechar()
    
    If (m_objConexao.State <> adStateClosed) Then
        m_objConexao.Close
    End If
        
End Sub

Public Function ExecuteReader(ByVal vSql As String, Optional ByVal vComandoTipo As CommandTypeEnum = adCmdText) As ADODB.Recordset

    Call PrepararComando(m_objComando, vSql, vComandoTipo)
    
    Set ExecuteReader = m_objComando.Execute
        
End Function

Public Function ExecuteNonQuery(ByVal vSql As String, Optional ByVal vComandoTipo As CommandTypeEnum = adCmdText) As Long

    Dim lngLinhasAfetadas As Long
    
    Call PrepararComando(m_objComando, vSql, vComandoTipo)
    
    Call m_objComando.Execute(lngLinhasAfetadas)
    
    ExecuteNonQuery = lngLinhasAfetadas

End Function

Private Sub PrepararComando(ByVal vComando As ADODB.Command, ByVal vSql As String, Optional ByVal vComandoTipo As CommandTypeEnum = adCmdText)
        
    vComando.ActiveConnection = m_objConexao
    vComando.Prepared = True
    vComando.CommandText = vSql
    vComando.CommandType = vComandoTipo
    
End Sub

Public Sub DestruirObjetos()

    Set m_objConexao = Nothing
    Set m_objComando = Nothing

End Sub

Private Sub Class_Initialize()
    
    m_enuProvedorTipo = ConnectionString
    
End Sub

Private Sub Class_Terminate()
     Call DestruirObjetos
End Sub

Public Sub IniciarTransacao()

    If m_blnTransacao Then
        Exit Sub
    End If
    
    m_objConexao.BeginTrans
    

End Sub

Public Sub ConfirmarTransacao()

    If Not m_blnTransacao Then
        Exit Sub
    End If
    
    m_objConexao.CommitTrans

End Sub

Public Sub DesfazerTransacao()

    If Not m_blnTransacao Then
        Exit Sub
    End If
    
    m_objConexao.RollbackTrans

End Sub

Private Sub Erros(ByVal vErros As Errors)

    Dim i                As Integer
    Dim Erro             As Error
    Dim strMensagemErro  As String

    For Each Erro In vErros
    
        i = i + 1
        strMensagemErro = strMensagemErro & vbCrLf & "ADO Error # " & i & ":"
        strMensagemErro = strMensagemErro & vbCrLf & "   ADO Error   # " & Erro.Number
        strMensagemErro = strMensagemErro & vbCrLf & "   Description   " & Erro.Description
        strMensagemErro = strMensagemErro & vbCrLf & "   Source        " & Erro.Source

    Next

End Sub
