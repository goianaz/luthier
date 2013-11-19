VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsShippersDTO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private m_intShipperID As Integer
Private m_strCompanyName As String
Private m_strPhone As String

Public Property Get ShipperID() As Integer

    ShipperID = m_intShipperID

End Property

Public Property Let ShipperID(ByVal intShipperID As Integer)

    m_intShipperID = intShipperID

End Property

Public Property Get CompanyName() As String

    CompanyName = m_strCompanyName

End Property

Public Property Let CompanyName(ByVal strCompanyName As String)

    m_strCompanyName = strCompanyName

End Property

Public Property Get Phone() As String

    Phone = m_strPhone

End Property

Public Property Let Phone(ByVal strPhone As String)

    m_strPhone = strPhone

End Property
