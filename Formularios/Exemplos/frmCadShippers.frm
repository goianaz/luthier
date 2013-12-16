VERSION 5.00
Begin VB.Form frmCadShippers 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   2040
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCodigo 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   285
      Left            =   960
      TabIndex        =   4
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtNome 
      Height          =   285
      Left            =   960
      TabIndex        =   3
      Top             =   720
      Width           =   4575
   End
   Begin VB.TextBox txtTelefone 
      Height          =   285
      Left            =   960
      TabIndex        =   2
      Top             =   1080
      Width           =   2055
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   2880
      TabIndex        =   1
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   4320
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Codigo "
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   360
      Width           =   540
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   720
      Width           =   420
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Telefone"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   630
   End
End
Attribute VB_Name = "frmCadShippers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function Validar() As String

    If Trim(txtNome.Text) = vbNullString Then
        Validar = "Nome inválido."
        Exit Function
    End If
    
    If Trim(txtTelefone.Text) = vbNullString Then
        Validar = "Telefone inválido."
        Exit Function
    End If

End Function

Private Sub cmdCancelar_Click()
    Call Unload(Me)
End Sub

Private Sub cmdOk_Click()

    Dim strRet As String
    Dim clsShippersDTO As clsShippersDTO
        
    strRet = Validar
    
    If strRet <> vbNullString Then
        Call MsgBox(strRet, vbInformation, TITULO_MENSAGENS)
        GoTo Finalizar
    End If
    
    Set clsShippersDTO = New clsShippersDTO
    
    If txtCodigo.Text <> vbNullString Then
        clsShippersDTO.ShipperID = CLng(txtCodigo.Text)
    End If
    
    clsShippersDTO.CompanyName = txtNome.Text
    clsShippersDTO.Phone = txtTelefone
        
    Call Salvar(clsShippersDTO)
    
    Call MsgBox("Registro salvo com sucesso.", vbInformation, TITULO_MENSAGENS)
        
Finalizar:
    
    Set clsShippersDTO = Nothing

End Sub

Private Sub Salvar(ByVal clsShippersDTO As clsShippersDTO)

    Dim clsShippersDAO   As clsShippersDAO
    
    Set clsShippersDAO = New clsShippersDAO

    If clsShippersDTO.ShipperID = 0 Then
            
        Call clsShippersDAO.Incluir(clsShippersDTO)
            
    Else
    
        Call clsShippersDAO.Alterar(clsShippersDTO)
    
    End If
    
    
Finalizar:

    Set clsShippersDAO = Nothing

End Sub
