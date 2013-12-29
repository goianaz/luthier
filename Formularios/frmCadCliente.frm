VERSION 5.00
Begin VB.Form frmCadCliente 
   Caption         =   "Cadastro de Clientes"
   ClientHeight    =   4425
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7920
   Icon            =   "frmCadCliente.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   7920
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2640
      TabIndex        =   21
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   375
      Left            =   1320
      TabIndex        =   20
      Top             =   3720
      Width           =   1215
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   1080
      TabIndex        =   19
      Text            =   "txtEmail"
      Top             =   3120
      Width           =   2775
   End
   Begin VB.TextBox txtFonefixo 
      Height          =   285
      Left            =   3840
      TabIndex        =   18
      Text            =   "txtFonefixo"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtFonecel 
      Height          =   285
      Left            =   1080
      TabIndex        =   17
      Text            =   "txtFonecel"
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox txtUf 
      Height          =   285
      Left            =   1080
      TabIndex        =   16
      Text            =   "txtUf"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.TextBox txtCidade 
      Height          =   285
      Left            =   1080
      TabIndex        =   15
      Text            =   "txtCidade"
      Top             =   2040
      Width           =   4575
   End
   Begin VB.TextBox txtBairro 
      Height          =   285
      Left            =   1080
      TabIndex        =   14
      Text            =   "txtBairro"
      Top             =   1680
      Width           =   2175
   End
   Begin VB.TextBox txtNro 
      Height          =   285
      Left            =   1080
      TabIndex        =   13
      Text            =   "txtNro"
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox txtEndereco 
      Height          =   285
      Left            =   1080
      TabIndex        =   12
      Text            =   "txtEndereco"
      Top             =   960
      Width           =   4575
   End
   Begin VB.TextBox txtNome 
      Height          =   285
      Left            =   1080
      TabIndex        =   11
      Text            =   "txtNome"
      Top             =   600
      Width           =   4575
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Text            =   "txtCodigo"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Email"
      Height          =   195
      Index           =   9
      Left            =   240
      TabIndex        =   9
      Top             =   3120
      Width           =   375
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Fixo"
      Height          =   195
      Index           =   8
      Left            =   3360
      TabIndex        =   8
      Top             =   2760
      Width           =   285
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Celular"
      Height          =   195
      Index           =   7
      Left            =   240
      TabIndex        =   7
      Top             =   2760
      Width           =   480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "UF"
      Height          =   195
      Index           =   6
      Left            =   240
      TabIndex        =   6
      Top             =   2400
      Width           =   210
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Bairro"
      Height          =   195
      Index           =   5
      Left            =   240
      TabIndex        =   5
      Top             =   1680
      Width           =   405
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Cidade"
      Height          =   195
      Index           =   4
      Left            =   240
      TabIndex        =   4
      Top             =   2040
      Width           =   495
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Nro"
      Height          =   195
      Index           =   3
      Left            =   240
      TabIndex        =   3
      Top             =   1320
      Width           =   255
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Endereco"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   960
      Width           =   690
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   195
      Index           =   1
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   420
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Codigo "
      Height          =   195
      Index           =   0
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   540
   End
End
Attribute VB_Name = "frmCadCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
