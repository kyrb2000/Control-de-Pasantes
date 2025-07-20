VERSION 5.00
Begin VB.Form frmContraseña 
   BorderStyle     =   0  'None
   Caption         =   "Contraseña"
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1125
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame7 
      Height          =   1095
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   3495
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Default         =   -1  'True
         Height          =   300
         Left            =   1800
         TabIndex        =   4
         Top             =   720
         Width           =   975
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   300
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Width           =   975
      End
      Begin VB.TextBox txtContraseña 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   120
         MaxLength       =   15
         PasswordChar    =   "*"
         TabIndex        =   0
         Top             =   345
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
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Tabla_Gerenar"
         Top             =   720
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Contraseña :"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   900
      End
   End
End
Attribute VB_Name = "FRMContraseña"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Respuesta As Boolean

Private Sub cmdAceptar_Click()
    If Data2.Recordset.Fields("Contraseña") = txtContraseña Then
        Respuesta = True
        Me.Hide
    Else
        MsgBox "La Contraseña es Errada"
        Call cmdCancelar_Click
    End If
End Sub

Private Sub cmdCancelar_Click()
    Respuesta = False
    Me.Hide
End Sub

Private Sub Form_Load()
        FRMContraseña.Move (Screen.Width - FRMContraseña.Width) / 2, (Screen.Height - FRMContraseña.Height) / 2
End Sub
