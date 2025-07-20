VERSION 5.00
Begin VB.Form frmBuscar_por 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Buscar por Items"
   ClientHeight    =   1965
   ClientLeft      =   3915
   ClientTop       =   3780
   ClientWidth     =   3735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1965
   ScaleWidth      =   3735
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Buscar por..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Aceptar"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   480
         TabIndex        =   5
         Top             =   1320
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1860
         TabIndex        =   4
         Top             =   1320
         Width           =   1335
      End
      Begin VB.ComboBox cboBuscar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         ItemData        =   "frmBuscar_por.frx":0000
         Left            =   1200
         List            =   "frmBuscar_por.frx":0007
         TabIndex        =   2
         Text            =   "Cedula"
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtBuscar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   3495
      End
      Begin VB.Label Label1 
         Caption         =   "Campo :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmBuscar_por"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public txtCampo01 As String
Public txtCampo02 As String

Private Sub cmdCancelar_Click()
    txtCampo01 = ""
    txtCampo02 = ""
    Me.Hide
End Sub

Private Sub cmdUpdate_Click()
    txtCampo01 = txtBuscar
    If cboBuscar.ListIndex = -1 Then cboBuscar.ListIndex = 0
    txtCampo02 = cboBuscar.List(cboBuscar.ListIndex)
    Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call cmdCancelar_Click
End Sub
