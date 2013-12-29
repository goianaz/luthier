VERSION 5.00
Begin VB.Form frmCadInstrumento 
   Caption         =   "Cadastro de Instrumentos"
   ClientHeight    =   2745
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5835
   Icon            =   "frmCadInstrumento.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2745
   ScaleWidth      =   5835
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2040
      TabIndex        =   9
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalvar 
      Caption         =   "&Salvar"
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   2160
      Width           =   1215
   End
   Begin VB.TextBox txtMarca 
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Text            =   "txtMarca"
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtModelo 
      Height          =   285
      Left            =   840
      TabIndex        =   6
      Text            =   "txtModelo"
      Top             =   840
      Width           =   3255
   End
   Begin VB.TextBox txtNome 
      Height          =   285
      Left            =   840
      TabIndex        =   5
      Text            =   "txtNome"
      Top             =   480
      Width           =   3255
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   840
      TabIndex        =   4
      Text            =   "txtCodigo"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Marca"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   450
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Modelo"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   525
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Nome"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   420
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Codigo"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmCadInstrumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
