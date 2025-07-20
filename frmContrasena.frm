VERSION 5.00
Begin VB.Form frmContrasena 
   BorderStyle     =   0  'None
   Caption         =   "Contraseña"
   ClientHeight    =   2190
   ClientLeft      =   3765
   ClientTop       =   3660
   ClientWidth     =   4260
   DrawStyle       =   5  'Transparent
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4215
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
         Height          =   900
         Left            =   2160
         Picture         =   "frmContrasena.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1080
         Width           =   1455
      End
      Begin VB.CommandButton cmdAceptar 
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
         Height          =   900
         Left            =   600
         Picture         =   "frmContrasena.frx":0427
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtContraseña 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   480
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   585
         Width           =   3255
      End
      Begin VB.Data Data2 
         Caption         =   "Tabla_Gerenar"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1320
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Tabla_Gerenar"
         Top             =   120
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   240
         Left            =   120
         Picture         =   "frmContrasena.frx":0828
         Top             =   150
         Width           =   240
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   480
         TabIndex        =   4
         Top             =   240
         Width           =   1560
      End
   End
End
Attribute VB_Name = "frmContrasena"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
'-- Autor : Kenner Y. Roa B. Telefono (0212) 433.55.14
'-- E-mail : kyrb2000@cantv.net / kyrb2000@hotmail.com
'-- Fecha : 22/05/2002 (Caracas / Venezuela)
'-----------------------------------------------------------
Public Respuesta As Boolean
Dim Contador As Integer

Private Sub cmdAceptar_Click()
    Contador = Contador + 1
    If Data2.Recordset.Fields("Contraseña") = txtContraseña Then
        Respuesta = True
        Me.Hide
        Exit Sub
    End If
    If Contador = 3 Then
        MsgBox "La Contraseña es Errada"
        Call cmdCancelar_Click
    Else
        MsgBox "La Contraseña es Errada"
        txtContraseña = ""
    End If
End Sub

Private Sub cmdCancelar_Click()
    Respuesta = False
    Me.Hide
End Sub

Private Sub Form_Load()
    frmContrasena.Move (Screen.Width - frmContrasena.Width) / 2, (Screen.Height - frmContrasena.Height) / 2
    Contador = 0
End Sub

Private Sub txtContraseña_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdAceptar.SetFocus
End Sub
