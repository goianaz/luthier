VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsInstrumentoDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lngCodigo As Long
Private m_strNome As String
Private m_strModelo As String

Public Property Get Codigo() As Long

    Codigo = m_lngCodigo

End Property

Public Property Get Modelo() As String

    Modelo = m_strModelo

End Property

Public Property Let Modelo(ByVal strModelo As String)

    m_strModelo = strModelo

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
