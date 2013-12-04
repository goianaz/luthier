VERSION 5.00
Begin VB.Form frmCadOs 
   Caption         =   "Ordem De Serviço"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   Icon            =   "frmCadOs.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   495
      Left            =   2400
      TabIndex        =   13
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   495
      Left            =   960
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.TextBox txtInstrumento 
      Height          =   285
      Left            =   1200
      TabIndex        =   11
      Text            =   "txtInstrumento"
      Top             =   1920
      Width           =   2655
   End
   Begin VB.TextBox txtServico 
      Height          =   285
      Left            =   1200
      TabIndex        =   10
      Text            =   "txtServico"
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtCliente 
      Height          =   285
      Left            =   1200
      TabIndex        =   9
      Text            =   "txtCliente"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox txtDataEntrega 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Text            =   "txtDataEntrega"
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox txtData 
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Text            =   "txtData"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox txtCodigo 
      Height          =   285
      Left            =   1200
      TabIndex        =   6
      Text            =   "txtCodigo"
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Instrumento"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   825
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Serviço"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   540
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Cliente"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   480
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Data Entrega"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label 
      AutoSize        =   -1  'True
      Caption         =   "Data"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   345
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
Attribute VB_Name = "frmCadOs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
