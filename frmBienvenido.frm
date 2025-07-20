VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBienvenido 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   1230
   ClientTop       =   1725
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   5820
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame_Base_Datos 
      Caption         =   "Frame_Base_Datos"
      Height          =   2175
      Left            =   480
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox txtPeriodo 
         DataField       =   "Periodo"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   120
         TabIndex        =   5
         Top             =   1680
         Width           =   2535
      End
      Begin VB.Data Data3 
         Caption         =   "Tabla_Gerenar"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Tabla_Gerenar"
         Top             =   240
         Visible         =   0   'False
         Width           =   2505
      End
      Begin VB.TextBox txtNombre_Empresa 
         DataField       =   "Nombre_Empresa"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtRif_Empresa 
         DataField       =   "Rif_Empresa"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   2535
      End
      Begin MSAdodcLib.Adodc Adodc1_Diario 
         Height          =   330
         Left            =   120
         Top             =   1320
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Adodc1_Diario"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   1920
      Top             =   720
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Se autoriza el uso de este producto a :"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   2010
      TabIndex        =   1
      Tag             =   "Descripción de la aplicación"
      Top             =   3720
      Width           =   4095
   End
   Begin VB.Label lblAutorizado 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   3915
      Width           =   4095
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4245
      Left            =   0
      Picture         =   "frmBienvenido.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5805
   End
End
Attribute VB_Name = "frmBienvenido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Tiempo As Integer

Private Sub Form_Activate()
    Data3.DatabaseName = Base_de_Datos
    Data3.Refresh
    WEmpresa = txtNombre_Empresa
    '"Instituto Universitario de Tecnología de Administración Industrial (IUTA)"
    WRif = txtRif_Empresa
    '"Región Capital"
    X_Periodo = txtPeriodo
    lblAutorizado.Caption = WEmpresa & Chr(10) & Chr(13) & WRif
End Sub

Private Sub Form_Load()
    DSN_Pasantias = "DSN=Pasantias"
    Adodc1_Diario.ConnectionString = DSN_Pasantias
    Adodc1_Diario.RecordSource = "Diario"
    Adodc1_Diario.Refresh
    If Not Adodc1_Diario.Recordset.EOF Then
        Buscar = "[Fecha]=" & Date
        Adodc1_Diario.Recordset.MoveFirst
        Adodc1_Diario.Recordset.Find Buscar
        msg_Tipo_Msg = "-"
        If Adodc1_Diario.Recordset.EOF Then
'            MsgBox "No Existe la Fecha Actual en el Anuario", vbCritical, "Buscar Fecha Anuario"
        Else
            msg_Fecha = Adodc1_Diario.Recordset.Fields("Fecha")
            msg_Hora = Adodc1_Diario.Recordset.Fields("Hora")
            msg_Mensaje = Adodc1_Diario.Recordset.Fields("Mensaje")
            msg_Tipo_Msg = "Urgente"
        End If
    End If
    Adodc1_Diario.Recordset.Close
    Tiempo = 0
End Sub

Private Sub Timer1_Timer()
    If Tiempo = 5 Then
        Unload Me
    Else
        Tiempo = Tiempo + 1
    End If
End Sub
