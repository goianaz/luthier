VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOsDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lngCodigo As Long
Private m_dtData As Date
Private m_dtDataEntrega As Date
Private m_lngCodCLiente As Long
Private m_lngCodServico As Long
Private m_lngCodInstrumento As Long
Private m_curValor As Currency


Public Property Get Codigo() As Long

    Codigo = m_lngCodigo

End Property

Public Property Get CodCLiente() As Integer

    CodCLiente = m_lngCodCLiente
    

End Property

Public Property Let CodCLiente(ByVal lngCodCLiente As Long)

    m_lCodCLiente = lngCodCLiente

End Property

Public Property Let Codigo(ByVal lngCodigo As Long)

    m_lngCodigo = lngCodigo

End Property


Public Property Get Data() As Date

    Data = m_dtData

End Property

Public Property Let Data(ByVal dtData As Date)

    m_dtData = dtData

End Property

Public Property Get DataEntrega() As Date

    DataEntrega = m_dtDataEntrega

End Property

Public Property Let DataEntrega(ByVal dtDataEntrega As Date)

    m_dtDataEntrega = dtDataEntrega

End Property

Public Property Get CodServico() As Long

    CodServico = m_lngCodServico

End Property

Public Property Let CodServico(ByVal lCodServico As Long)

    m_lngCodServico = lngCodServico

End Property

Public Property Get CodInstrumento() As Long

    CodInstrumento = m_lngCodInstrumento

End Property

Public Property Let CodInstrumento(ByVal lCodInstrumento As Long)

    m_lngCodInstrumento = lngCodInstrumento

End Property

Public Property Get Valor() As Currency

    Valor = m_curValor

End Property

Public Property Let Valor(ByVal curValor As Currency)

    m_curValor = curValor

End Property

