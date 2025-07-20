VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCartas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cartas de "
   ClientHeight    =   6375
   ClientLeft      =   960
   ClientTop       =   1335
   ClientWidth     =   8895
   Icon            =   "frmCartas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   8895
   Begin VB.Frame Frame_Base_Datos 
      Caption         =   "Frame_Base_Datos"
      Height          =   2895
      Left            =   1680
      TabIndex        =   33
      Top             =   1560
      Visible         =   0   'False
      Width           =   3975
      Begin VB.Data DSolvencias 
         Caption         =   "Solvencias"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Solvencias"
         Top             =   2400
         Visible         =   0   'False
         Width           =   2290
      End
      Begin VB.Data DControl 
         Caption         =   "Control"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Control"
         Top             =   2040
         Visible         =   0   'False
         Width           =   2290
      End
      Begin VB.TextBox txtCodigo_Tipo 
         DataField       =   "Codigo_Tipo"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "txtCodigo_Tipo"
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtNumero_ID 
         DataField       =   "Numero_ID"
         DataSource      =   "Data1"
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "txtNumero_ID"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox adodc2_txtNumero 
         DataField       =   "Numero"
         DataSource      =   "Adodc2"
         Height          =   285
         Left            =   2520
         TabIndex        =   35
         Text            =   "adodc2_txtNumero"
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Data Data2 
         Caption         =   "Alumnos"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Alumnos"
         Top             =   1680
         Visible         =   0   'False
         Width           =   2290
      End
      Begin VB.Data Data3 
         Caption         =   "Especialidad"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Especialidad"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2290
      End
      Begin VB.Data Data4 
         Caption         =   "Tabla_Gerenar"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Tabla_Gerenar"
         Top             =   960
         Visible         =   0   'False
         Width           =   2290
      End
      Begin VB.TextBox adodc1_txtNumero 
         DataField       =   "Numero"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   2520
         TabIndex        =   34
         Text            =   "adodc1_txtNumero"
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   120
         Top             =   240
         Visible         =   0   'False
         Width           =   2290
         _ExtentX        =   4048
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=Reporte_Mov"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "Reporte_Mov"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Cartas_Postulacion"
         Caption         =   "Cartas_Postulacion"
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
      Begin MSAdodcLib.Adodc Adodc2 
         Height          =   330
         Left            =   120
         Top             =   600
         Visible         =   0   'False
         Width           =   2290
         _ExtentX        =   4048
         _ExtentY        =   582
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   3
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   "DSN=Reporte_Mov"
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "Reporte_Mov"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Cartas_Presentacion"
         Caption         =   "Cartas_Presentacion"
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
   Begin VB.TextBox txtFecha_Solicitud 
      Alignment       =   2  'Center
      DataField       =   "Fecha_Solicitud"
      DataSource      =   "Data1"
      Enabled         =   0   'False
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
      Left            =   6120
      MaxLength       =   25
      TabIndex        =   32
      Top             =   45
      Width           =   2655
   End
   Begin VB.TextBox txtNumero 
      DataField       =   "Numero"
      DataSource      =   "Data1"
      Enabled         =   0   'False
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
      Left            =   7320
      MaxLength       =   15
      TabIndex        =   31
      Top             =   480
      Width           =   1455
   End
   Begin VB.TextBox txtCedula 
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
      Left            =   1440
      Locked          =   -1  'True
      MaxLength       =   9
      TabIndex        =   7
      Top             =   480
      Width           =   1695
   End
   Begin VB.TextBox txtApeNom 
      Enabled         =   0   'False
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
      Left            =   2760
      TabIndex        =   30
      Top             =   960
      Width           =   6015
   End
   Begin VB.TextBox txtTurno 
      Enabled         =   0   'False
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
      Left            =   6480
      TabIndex        =   27
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CheckBox chkImprimida 
      Caption         =   "Imprimida"
      DataField       =   "Imprimida"
      DataSource      =   "Data1"
      Enabled         =   0   'False
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
      Left            =   3600
      TabIndex        =   26
      Top             =   540
      Width           =   1575
   End
   Begin VB.TextBox txtEspecialidad 
      DataField       =   "Especialidad"
      DataSource      =   "Data1"
      Enabled         =   0   'False
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
      Left            =   1920
      TabIndex        =   24
      Top             =   1395
      Width           =   3495
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
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
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Cartas"
      Top             =   5970
      Width           =   8895
   End
   Begin VB.Frame Frame_Datos_de_la_Empresas 
      Caption         =   "Datos de la Empresas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   1920
      Width           =   8775
      Begin VB.TextBox txtDepartamento_Emp 
         DataField       =   "Departamento_Emp"
         DataSource      =   "Data1"
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
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   9
         Top             =   720
         Width           =   6615
      End
      Begin VB.TextBox txtNombre_Emp 
         DataField       =   "Nombre_Emp"
         DataSource      =   "Data1"
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
         Left            =   2040
         MaxLength       =   150
         TabIndex        =   8
         Top             =   270
         Width           =   6615
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
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
         Index           =   2
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Departamento :"
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
         Index           =   3
         Left            =   120
         TabIndex        =   17
         Top             =   795
         Width           =   1890
      End
   End
   Begin VB.Frame Frame_Destinatario 
      Caption         =   "Destinatario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   19
      Top             =   3120
      Width           =   8775
      Begin VB.ComboBox CobDestinatario_Cargo 
         DataField       =   "Destinatario_Prof"
         DataSource      =   "Data1"
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
         ItemData        =   "frmCartas.frx":0442
         Left            =   2760
         List            =   "frmCartas.frx":046D
         TabIndex        =   10
         Text            =   "-"
         Top             =   240
         Width           =   975
      End
      Begin VB.TextBox txtDestinatario_Cargo 
         DataField       =   "Destinatario_Cargo"
         DataSource      =   "Data1"
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   13
         Top             =   1185
         Width           =   6615
      End
      Begin VB.TextBox txtDestinatario_Profesion 
         DataField       =   "Destinatario_Profesion"
         DataSource      =   "Data1"
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
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   12
         Top             =   735
         Width           =   6615
      End
      Begin VB.TextBox txtDestinatario_Nombre 
         DataField       =   "Destinatario_Nombre"
         DataSource      =   "Data1"
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
         Left            =   3840
         MaxLength       =   50
         TabIndex        =   11
         Top             =   285
         Width           =   4815
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Apellidos y Nombres :"
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
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   2595
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Profesión :"
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
         Index           =   5
         Left            =   120
         TabIndex        =   21
         Top             =   810
         Width           =   1305
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Cargo :"
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
         Index           =   6
         Left            =   120
         TabIndex        =   20
         Top             =   1245
         Width           =   885
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   720
      TabIndex        =   6
      Top             =   4800
      Width           =   8175
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cerrar"
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
         Left            =   6720
         Picture         =   "frmCartas.frx":04C6
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   150
         Width           =   1335
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
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
         Left            =   5400
         Picture         =   "frmCartas.frx":08ED
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Modificar"
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
         Left            =   4080
         Picture         =   "frmCartas.frx":0D08
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         Width           =   1335
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
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
         Left            =   2760
         Picture         =   "frmCartas.frx":1130
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   150
         Width           =   1335
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Eliminar"
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
         Left            =   1440
         Picture         =   "frmCartas.frx":156D
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   150
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Agregar"
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
         Left            =   120
         Picture         =   "frmCartas.frx":19D6
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   150
         Width           =   1335
      End
   End
   Begin VB.Label lblTipo_Carta 
      BorderStyle     =   1  'Fixed Single
      DataField       =   "Codigo_Tipo"
      DataSource      =   "Data1"
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
      TabIndex        =   38
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Apellidos y Nombres :"
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
      Index           =   8
      Left            =   120
      TabIndex        =   29
      Top             =   960
      Width           =   2595
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Turno :"
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
      Index           =   10
      Left            =   5520
      TabIndex        =   28
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Especialidad :"
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
      Index           =   9
      Left            =   120
      TabIndex        =   25
      Top             =   1395
      Width           =   1695
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Numero :"
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
      Index           =   7
      Left            =   6120
      TabIndex        =   23
      Top             =   540
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Fecha de Solicitud :"
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
      Index           =   1
      Left            =   3600
      TabIndex        =   15
      Top             =   60
      Width           =   2400
   End
   Begin VB.Label lblLabels 
      AutoSize        =   -1  'True
      Caption         =   "Cedula :"
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
      Index           =   0
      Left            =   120
      TabIndex        =   14
      Top             =   540
      Width           =   1005
   End
End
Attribute VB_Name = "frmCartas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
'-- Autor : Kenner Y. Roa B. Telefono (0212) 433.55.14
'-- E-mail : kyrb2000@cantv.net / kyrb2000@hotmail.com
'-- Fecha : 22/05/2002 (Caracas / Venezuela)
'-----------------------------------------------------------
Public Tipo_Carta As String
Public Mensaje_Act As Boolean

Private Sub cmdAdd_Click()
    Dim Numero As Integer
    txtCedula = ""
    Data1.Recordset.AddNew
    txtFecha_Solicitud = Format(Date, "dd/mm/yyyy")
    txtTipo_Carta = Tipo_Carta
    lblTipo_Carta.Caption = Tipo_Carta
    txtDepartamento_Emp = "Dpto. Recursos Humanos"
    txtNumero = "."
    CobDestinatario_Cargo.Text = "-"
    Frame_Datos_de_la_Empresas.Enabled = True
    Frame_Destinatario.Enabled = True
    '--------- Botones ------------
    cmdUpdate.Caption = "&Actualizar"
    cmdUpdate.Enabled = True
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdBuscar.Enabled = False
    cmdImprimir.Enabled = False
    cmdClose.Caption = "&Cancelar"
    Data1.Enabled = False
    '------------------------------'
    txtCedula.Locked = False
    txtCedula.SetFocus
End Sub

Private Sub cmdDelete_Click()
    Dim Respuesta
    If Not Data1.Recordset.EOF Then
        'esto puede producir un error si elimina el último
        'registro o el único registro del recordset
        Respuesta = MsgBox("Desea Eliminar el Registro", vbYesNo + vbCritical, "Eliminación de Datos")
        If Respuesta = vbYes Then
            Data1.Recordset.Delete
            Data1.Recordset.MoveFirst
        End If
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim Buscar, Campo_Sel, WNumero_ID As String
    If (Data1.Recordset.AbsolutePosition + 1) = 0 Then
        Exit Sub
    End If
    '************************************************
    frmBuscar_por.cboBuscar.AddItem "Numero"
    frmBuscar_por.cboBuscar.AddItem "Numero_ID"
    frmBuscar_por.cboBuscar.AddItem "Nombres"
    frmBuscar_por.cboBuscar.AddItem "Apellidos"
    frmBuscar_por.Show vbModal
    Buscar = frmBuscar_por.txtCampo01
    Campo_Sel = frmBuscar_por.txtCampo02
    Unload frmBuscar_por
    '************************************************
'    Buscar = InputBox("Introduzca La Cedula :", "Busqueda de Datos")
    If Buscar <> "" Then
        Select Case Campo_Sel
        Case "Numero_ID":
            Buscar = "[" & Campo_Sel & "]" & "=" & Buscar
        Case "Numero":
            Buscar = "[" & Campo_Sel & "]" & "=" & Buscar
        Case Else
            Buscar = "[" & Campo_Sel & "]" & "=" & "'" & Buscar & "'"
'        Buscar = "[Cedula]" & "=" & "'" & Buscar & "'"
        End Select
        If Campo_Sel <> "Numero" Then
            Data2.Recordset.FindFirst Buscar
            WNumero_ID = Trim(Data2.Recordset.Fields("Numero_ID"))
            Buscar = "[Numero_ID]" & "=" & WNumero_ID
        End If
        Data1.Recordset.MoveFirst
        Data1.Recordset.FindFirst Buscar
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim Buscar As String
    Dim Seccion, Prof As String
    Mensaje_Act = False
    On Error Resume Next
    If Tipo_Carta = "POS" Then
        If Not Adodc1.Recordset.EOF Then
            Adodc1.Recordset.MoveFirst
            Do Until Adodc1.Recordset.EOF
                If Not Adodc1.Recordset.EOF Then
                    Adodc1.Recordset.Delete
                End If
            Adodc1.Recordset.MoveFirst
            Loop
        End If
    Else
        If Not Adodc2.Recordset.EOF Then
            Adodc2.Recordset.MoveFirst
            Do Until Adodc2.Recordset.EOF
                If Not Adodc2.Recordset.EOF Then
                    Adodc2.Recordset.Delete
                End If
            Adodc2.Recordset.MoveFirst
            Loop
        End If
    End If
    Data1.Recordset.MoveFirst
    Do Until Data1.Recordset.EOF
        If Tipo_Carta = "POS" Then
            If Not Data1.Recordset.EOF Then
                If Not Data1.Recordset.Fields("Imprimida") Then
                    Data1.Recordset.Edit
                    Data1.Recordset.Fields("Imprimida") = True
                    Prof = Data1.Recordset.Fields("Destinatario_Prof")
                    Data1.Recordset.Update
                    Buscar = "[Numero_ID]" & "=" & txtNumero_ID
                    Data2.Recordset.FindFirst Buscar
                    If Data2.Recordset.NoMatch Then
                        Buscar = "[Cedula]" & "=" & "'" & Data1.Recordset.Fields("Cedula") & "'"
                        Data2.Recordset.FindFirst Buscar
                        If Data2.Recordset.NoMatch Then
                            MsgBox "Cedula No Encontrada, Favor Chequear", , "Busqueda de Registro"
                        End If
                    Else
                        Seccion = Data2.Recordset.Fields("Seccion")
                    End If
                    Buscar = "[Codigo]" & "=" & "'" & Mid(Seccion, 1, 2) & "'"
                    Data3.Recordset.FindFirst Buscar
                    txtEspecialidad = Data3.Recordset.Fields("Descripcion")
                    Select Case Mid(Seccion, 5, 1)
                    Case "1"
                        txtTurno = "Mañana"
                    Case "3"
                        txtTurno = "Noche"
                    End Select
                    Adodc1.Recordset.AddNew
                    Adodc1.Recordset.Fields("Numero") = txtNumero
                    Adodc1.Recordset.Fields("Cedula") = Str(txtCedula)
                    Adodc1.Recordset.Fields("Nombre_Alumno") = Trim(txtApeNom)
                    Adodc1.Recordset.Fields("Especialidad") = Trim(txtEspecialidad)
                    Adodc1.Recordset.Fields("Fecha_Solicitud") = Mid(Format(txtFecha_Solicitud, "Long Date"), 2 + Len(StripFile(Format(txtFecha_Solicitud, "Long Date"), ",")), Len(Format(txtFecha_Solicitud, "Long Date")))
                    Adodc1.Recordset.Fields("Nombre_Emp") = Trim(txtNombre_Emp)
                    Adodc1.Recordset.Fields("Departamento_Emp") = Trim(txtDepartamento_Emp)
                    Adodc1.Recordset.Fields("Destinatario_Nombre") = Trim(txtDestinatario_Nombre)
                    Adodc1.Recordset.Fields("Destinatario_Profesion") = Trim(txtDestinatario_Profesion)
                    Adodc1.Recordset.Fields("Destinatario_Cargo") = Trim(txtDestinatario_Cargo)
                    Adodc1.Recordset.Fields("Destinatario_Prof") = Trim(Prof)
                    Adodc1.Recordset.Update
                End If
            End If
        Else
            If Not Data1.Recordset.EOF Then
                If Not Data1.Recordset.Fields("Imprimida") Then
                    Data1.Recordset.Edit
                    Data1.Recordset.Fields("Imprimida") = True
                    Prof = Data1.Recordset.Fields("Destinatario_Prof")
                    Data1.Recordset.Update
                    Buscar = "[Numero_ID]" & "=" & txtNumero_ID
                    Data2.Recordset.FindFirst Buscar
                    If Data2.Recordset.NoMatch Then
                        Buscar = "[Cedula]" & "=" & "'" & Data1.Recordset.Fields("Cedula") & "'"
                        Data2.Recordset.FindFirst Buscar
                        If Data2.Recordset.NoMatch Then
                            MsgBox "Cedula No Encontrada, Favor Chequear", , "Busqueda de Registro"
                        End If
                    Else
                        Seccion = Data2.Recordset.Fields("Seccion")
                    End If
                    Buscar = "[Codigo]" & "=" & "'" & Mid(Seccion, 1, 2) & "'"
                    Data3.Recordset.FindFirst Buscar
                    txtEspecialidad = Data3.Recordset.Fields("Descripcion")
                    Select Case Mid(Seccion, 5, 1)
                    Case "1"
                        txtTurno = "Mañana"
                    Case "3"
                        txtTurno = "Noche"
                    End Select
                    Adodc2.Recordset.AddNew
                    Adodc2.Recordset.Fields("Numero") = txtNumero
                    Adodc2.Recordset.Fields("Cedula") = Str(txtCedula)
                    Adodc2.Recordset.Fields("Nombre_Alumno") = Trim(txtApeNom)
                    Adodc2.Recordset.Fields("Especialidad") = Trim(txtEspecialidad)
                    Adodc2.Recordset.Fields("Fecha_Solicitud") = Mid(Format(txtFecha_Solicitud, "Long Date"), 2 + Len(StripFile(Format(txtFecha_Solicitud, "Long Date"), ",")), Len(Format(txtFecha_Solicitud, "Long Date")))
                    Adodc2.Recordset.Fields("Nombre_Emp") = Trim(txtNombre_Emp)
                    Adodc2.Recordset.Fields("Departamento_Emp") = Trim(txtDepartamento_Emp)
                    Adodc2.Recordset.Fields("Destinatario_Nombre") = Trim(txtDestinatario_Nombre)
                    Adodc2.Recordset.Fields("Destinatario_Profesion") = Trim(txtDestinatario_Profesion)
                    Adodc2.Recordset.Fields("Destinatario_Cargo") = Trim(txtDestinatario_Cargo)
                    Adodc2.Recordset.Fields("Destinatario_Prof") = Trim(Prof)
                    Adodc2.Recordset.Update
                End If
            End If
        End If
        Data1.Recordset.MoveNext
    Loop
    Data1.Recordset.MoveFirst
    Mensaje_Act = True
    If Tipo_Carta = "POS" Then
        Shell "C:\Archivos de programa\Microsoft Office\Office\winword.exe C:\Archiv~1\Sistem~2\Reportes\Postulación.doc"
    Else
        Shell "C:\Archivos de programa\Microsoft Office\Office\winword.exe C:\Archiv~1\Sistem~2\Reportes\Presentación.doc"
    End If
End Sub

Private Sub cmdUpdate_Click()
    If cmdUpdate.Caption = "&Actualizar" Then
        ' Si el Control Text esta en blanco
        ' Reeplazalo con un "-"
        For i = 0 To Me.Controls.Count - 1
            If TypeOf Me.Controls(i) Is TextBox Then
                If Me.Controls(i).Text = "" Then
                    Me.Controls(i).Text = "-"
                End If
            End If
        Next i
        ' ----------------------------------
        If txtNumero = "." Then
            If Tipo_Carta = "PRE" Then
                txtNumero = Data4.Recordset.Fields("NumeroCPres_ID")
            Else
                txtNumero = Data4.Recordset.Fields("NumeroCPost_ID")
            End If
            Numero_ID = Val(txtNumero) + 1
            Data4.Recordset.Edit
            If Tipo_Carta = "PRE" Then
                Data4.Recordset.Fields("NumeroCPres_ID") = Data4.Recordset.Fields("NumeroCPres_ID") + 1
            Else
                Data4.Recordset.Fields("NumeroCPost_ID") = Data4.Recordset.Fields("NumeroCPost_ID") + 1
            End If
            Data4.Recordset.Update
        End If
        lblTipo_Carta.Caption = Tipo_Carta
        Data1.UpdateRecord
        'Data1.Recordset.Bookmark = Data1.Recordset.LastModified
        Frame_Datos_de_la_Empresas.Enabled = False
        Frame_Destinatario.Enabled = False
        chkImprimida.Enabled = False
        txtCedula.Enabled = True
        '--------- Botones ------------
        cmdUpdate.Caption = "&Modificar"
        cmdAdd.Enabled = True
        cmdDelete.Enabled = True
        cmdBuscar.Enabled = True
        cmdImprimir.Enabled = True
        cmdClose.Caption = "&Cerrar"
        Data1.Enabled = True
        '------------------------------'
    Else
        If Not Data1.Recordset.EOF Then
            Data1.Recordset.Edit
            txtFecha_Solicitud = Format(Date, "dd/mm/yyyy")
            Frame_Datos_de_la_Empresas.Enabled = True
            Frame_Destinatario.Enabled = True
            chkImprimida.Enabled = True
            '--------- Botones ------------
            cmdUpdate.Caption = "&Actualizar"
            cmdAdd.Enabled = False
            cmdDelete.Enabled = False
            cmdBuscar.Enabled = False
            cmdImprimir.Enabled = False
            cmdClose.Caption = "&Cancelar"
            Data1.Enabled = False
            '------------------------------'
            txtCedula.Locked = True
            txtNombre_Emp.SetFocus
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    If cmdClose.Caption = "&Cerrar" Then
        Unload Me
    Else
        On Error GoTo Error_Al_Cancelar
        '--------- Botones ------------
        cmdUpdate.Caption = "&Modificar"
        cmdAdd.Enabled = True
        cmdDelete.Enabled = True
        cmdBuscar.Enabled = True
        cmdImprimir.Enabled = True
        cmdClose.Caption = "&Cerrar"
        Data1.Enabled = True
        '------------------------------'
        Data1.Recordset.CancelUpdate
    End If
Error_Al_Cancelar:
End Sub

Private Sub CobDestinatario_Cargo_Click()
    Select Case CobDestinatario_Cargo.List(CobDestinatario_Cargo.ListIndex)
    Case "Ing.":
        If Len(txtDestinatario_Profesion) < 11 Then
            txtDestinatario_Profesion = "Ingeniero"
        End If
    Case "Dr.", "Dra.":
        If Len(txtDestinatario_Profesion) < 11 Then
            txtDestinatario_Profesion = IIf(CobDestinatario_Cargo.List(CobDestinatario_Cargo.ListIndex) = "Dr.", "Doctor", "Doctora")
        End If
    Case "Lic.":
        If Len(txtDestinatario_Profesion) < 11 Then
            txtDestinatario_Profesion = "Licenciado"
        End If
    Case "T.S.U.":
        If Len(txtDestinatario_Profesion) < 11 Then
            txtDestinatario_Profesion = "T.S.U."
        End If
    Case "Prof.":
        If Len(txtDestinatario_Profesion) < 11 Then
            txtDestinatario_Profesion = "Profesor"
        End If
    Case "Econ.":
        If Len(txtDestinatario_Profesion) < 11 Then
            txtDestinatario_Profesion = "Economista"
        End If
    Case "GN.":
        If Len(txtDestinatario_Profesion) < 11 Then
            txtDestinatario_Profesion = "Guardia Naconal"
        End If
    Case "Coronel (EJ)":
        If Len(txtDestinatario_Profesion) < 11 Then
            txtDestinatario_Profesion = "Coronel del Ejercito"
        End If
    End Select
    txtDestinatario_Nombre.SetFocus
End Sub

Private Sub CobDestinatario_Cargo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13:
        txtDestinatario_Nombre.SetFocus
    End Select
End Sub

Private Sub Data1_Error(DataErr As Integer, Response As Integer)
  'Aquí es donde se coloca el código de control de errores
  'Si quiere ignorar los errores, marque como comentario la línea siguiente
  'Si desea detectarlos, agregue código aquí para controlarlos
  MsgBox "El error de datos alcanzó err:" & Error$(DataErr)
  Response = 0  'ignorar el error
End Sub

Private Sub Data1_Reposition()
  On Error Resume Next
  'Esto mostrará la posición del registro actual
  'para dynasets y snapshots
  Data1.Caption = "Record: " & (Data1.Recordset.AbsolutePosition + 1)
End Sub

Private Sub Data1_Validate(Action As Integer, Save As Integer)
  'Aquí es donde se coloca el código de validación
  'Se llama a este evento cuando se produce la siguiente acción
  Select Case Action
    Case vbDataActionMoveFirst
        If (Data1.Recordset.AbsolutePosition + 1) <> 0 Then
            Data1.Recordset.MoveFirst
            Call Buscar_Datos
            Data1.Recordset.MovePrevious
        End If
    Case vbDataActionMovePrevious
        If (Data1.Recordset.AbsolutePosition + 1) <> 0 Then
            Data1.Recordset.MovePrevious
            Call Buscar_Datos
            Data1.Recordset.MoveNext
        End If
    Case vbDataActionMoveNext
        If (Data1.Recordset.AbsolutePosition + 1) <> 0 Then
            Data1.Recordset.MoveNext
            Call Buscar_Datos
            Data1.Recordset.MovePrevious
        End If
    Case vbDataActionMoveLast
        If (Data1.Recordset.AbsolutePosition + 1) <> 0 Then
            Data1.Recordset.MoveLast
            Call Buscar_Datos
        End If
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select
End Sub

Private Sub Form_Activate()
    Adodc1.ConnectionString = DSN_Reporte_Mov
    Adodc1.RecordSource = "Cartas_Postulacion"
    Adodc1.Refresh
    Adodc2.ConnectionString = DSN_Reporte_Mov
    Adodc2.RecordSource = "Cartas_Presentacion"
    Adodc2.Refresh
    Data4.DatabaseName = Base_de_Datos
    Data4.Refresh
    Data3.DatabaseName = Base_de_Datos
    Data3.Refresh
    Data2.DatabaseName = Base_de_Datos
    Data2.Refresh
    DControl.DatabaseName = Base_de_Datos
    DControl.Refresh
    DSolvencias.DatabaseName = Base_de_Datos
    DSolvencias.Refresh
    Data1.DatabaseName = Base_de_Datos
    Data1.Refresh
    Call Data1_Validate(4, 0) 'Buscar_Datos
End Sub

Private Sub Form_Load()
    Frame_Datos_de_la_Empresas.Enabled = False
    Frame_Destinatario.Enabled = False
    txtCedula.Locked = True
    Mensaje_Act = True
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13:
        Call Buscar_Datos
        txtNombre_Emp.SetFocus
    End Select
End Sub

Sub Buscar_Datos()
    Dim Buscar As String
    Dim Seccion As String
    If txtNumero_ID = "" Then
        txtNumero_ID = 0
    End If
    Buscar = "[Numero_ID]" & "=" & txtNumero_ID
    Data2.Recordset.FindFirst Buscar
    If Data2.Recordset.NoMatch Then
        Buscar = "[Cedula]" & "=" & "'" & txtCedula & "'"
        Data2.Recordset.FindFirst Buscar
        If Data2.Recordset.NoMatch Then
            MsgBox "Cedula No Encontrada, Favor Chequear", , "Busqueda de Registro"
            Exit Sub
        End If
    End If
    txtCedula = Data2.Recordset.Fields("Cedula")
    txtNumero_ID = Data2.Recordset.Fields("Numero_ID")
    Seccion = Data2.Recordset.Fields("Seccion")
    txtApeNom = Data2.Recordset.Fields("Apellidos") & ", " & Data2.Recordset.Fields("Nombres")
    Buscar = "[Codigo]" & "=" & "'" & Mid(Seccion, 1, 2) & "'"
    Data3.Recordset.FindFirst Buscar
    txtEspecialidad = Data3.Recordset.Fields("Descripcion")
    Select Case Mid(Seccion, 5, 1)
    Case "1"
        txtTurno = "Mañana"
    Case "3"
        txtTurno = "Noche"
    End Select
    '------------------------------------------------------------------------
    If Mensaje_Act Then
        Contador = 0
        Do
            Select Case Contador
            Case 0:
                Titulos = "Taller de Pasantías"
                Tipo = "TA"
            Case 1:
                Titulos = "Control de Estudio"
                Tipo = "CO"
            Case 2:
                Titulos = "Departamento de Grado"
                Tipo = "GR"
            Case 3:
                Titulos = "Casos Especiales"
                Tipo = "CE"
            Case 4:
                Titulos = "Pasantías"
                Tipo = "PA"
            End Select
            Buscar = "[Cedula]" & "=" & "'" & txtCedula & Tipo & "'"
            DControl.Recordset.FindFirst Buscar
            If Not DControl.Recordset.NoMatch Then
                If DControl.Recordset.Fields("Cedula") = Trim(txtCedula & Tipo) Then
                    XObservacion = DControl.Recordset.Fields("Observacion")
                    If XObservacion <> "-" Then
                        MsgBox XObservacion, vbCritical, Titulos
                    End If
                End If
            End If
            Contador = Contador + 1
        Loop While Contador <= 3
        Buscar = "[Cedula]" & "=" & "'" & txtCedula & "'"
        DSolvencias.Recordset.FindFirst Buscar
        If Not DSolvencias.Recordset.NoMatch Then
           If DSolvencias.Recordset.Fields("Administrativo_Caja") = False Then
                Titulos = "Caja Solvencia Negada"
                XObservacion = "Alumno Con Caja Solvencia Negada"
                MsgBox XObservacion, vbCritical, Titulos
           End If
        End If
    End If
End Sub

Private Sub txtDepartamento_Emp_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13:
        CobDestinatario_Cargo.SetFocus
    End Select
End Sub

Private Sub txtDestinatario_Cargo_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13:
        cmdClose.SetFocus
    End Select
End Sub

Private Sub txtDestinatario_Nombre_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13:
        txtDestinatario_Profesion.SetFocus
    End Select
End Sub

Private Sub txtDestinatario_Profesion_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13:
        txtDestinatario_Cargo.SetFocus
    End Select
End Sub

Private Sub txtNombre_Emp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case KeyAscii
    Case 13:
        txtDepartamento_Emp.SetFocus
    End Select
End Sub
