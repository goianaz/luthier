VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClienteDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lngCodigo As Long
Private m_strNome As String
Private m_strEndereco As String
Private m_strNro As String
Private m_strCidade As String
Private m_strBairro As String
Private m_strUf As String
Private m_strFonecel As String
Private m_strFonefixo As String
Private m_strEmail As String

Public Property Get Codigo() As Long

    Codigo = m_lngCodigo

End Property

Public Property Let Codigo(ByVal lngCodigo As Long)

    m_lngCodigo = lngCodigo

End Property

Public Property Get Nome() As String

    Nome = m_strNome

End Property

Public Property Let Nome(ByVal strNome As String)

    m_strNome = strNome

End Property

Public Property Get Endereco() As String

    Endereco = m_strEndereco

End Property

Public Property Let Endereco(ByVal strEndereco As String)

    m_strEndereco = strEndereco

End Property

Public Property Get Nro() As String

    Nro = m_strNro

End Property

Public Property Let Nro(ByVal strNro As String)

    m_strNro = strNro

End Property

Public Property Get Cidade() As String

    Cidade = m_strCidade

End Property

Public Property Let Cidade(ByVal strCidade As String)

    m_strCidade = strCidade

End Property

Public Property Get Bairro() As String

    Bairro = m_strBairro

End Property

Public Property Let Bairro(ByVal strBairro As String)

    m_strBairro = strBairro

End Property

Public Property Get Uf() As String

    Uf = m_strUf

End Property

Public Property Let Uf(ByVal strUf As String)

    m_strUf = strUf

End Property

Public Property Get Fonecel() As String

    Fonecel = m_strFonecel

End Property

Public Property Let Fonecel(ByVal strFonecel As String)

    m_strFonecel = strFonecel

End Property

Public Property Get Fonefixo() As String

    Fonefixo = m_strFonefixo

End Property

Public Property Let Fonefixo(ByVal strFonefixo As String)

    m_strFonefixo = strFonefixo

End Property

Public Property Get Email() As String

    Email = m_strEmail

End Property

Public Property Let Email(ByVal strEmail As String)

    m_strEmail = strEmail

End Property
