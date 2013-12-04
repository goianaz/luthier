VERSION 5.00
Begin VB.Form frmCadServico 
   Caption         =   "Cadastro de Serviços"
   ClientHeight    =   1980
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "frmCadServico.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   1980
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   840
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.TextBox txtValor 
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Text            =   "txtValor"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtNome 
      Height          =   285
      Left            =   720
      TabIndex        =   4
      Text            =   "txtNome"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   720
      TabIndex        =   3
      Text            =   "txtCodigo"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Valor"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   360
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   420
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Código"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmCadServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Text_Change()

End Sub
