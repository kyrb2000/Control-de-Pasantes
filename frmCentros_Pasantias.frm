VERSION 5.00
Begin VB.Form frmCentros_Pasantias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Centros_Pasantias"
   ClientHeight    =   7200
   ClientLeft      =   720
   ClientTop       =   975
   ClientWidth     =   9120
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCentros_Pasantias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7200
   ScaleWidth      =   9120
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   120
      TabIndex        =   35
      Top             =   5640
      Width           =   8895
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
         Left            =   7320
         Picture         =   "frmCentros_Pasantias.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   150
         Width           =   1455
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
         Left            =   5880
         Picture         =   "frmCentros_Pasantias.frx":0869
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   150
         Width           =   1455
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
         Left            =   4440
         Picture         =   "frmCentros_Pasantias.frx":0C84
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   150
         Width           =   1455
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
         Left            =   3000
         Picture         =   "frmCentros_Pasantias.frx":10AC
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   150
         Width           =   1455
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
         Left            =   1560
         Picture         =   "frmCentros_Pasantias.frx":14E9
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   150
         Width           =   1455
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
         Picture         =   "frmCentros_Pasantias.frx":1952
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   9015
      Begin VB.TextBox txtEspecialidad 
         DataField       =   "Descripcion"
         DataSource      =   "Data3"
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
         Left            =   2040
         TabIndex        =   29
         Top             =   1035
         Width           =   3975
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
         Left            =   7200
         TabIndex        =   28
         Top             =   1080
         Width           =   1335
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
         TabIndex        =   27
         Top             =   600
         Width           =   5775
      End
      Begin VB.TextBox txtCedula 
         DataField       =   "Cedula"
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
         Left            =   1200
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   26
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtSeccion 
         Alignment       =   2  'Center
         DataField       =   "Seccion"
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
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   120
         Width           =   1455
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
         Index           =   12
         Left            =   6120
         TabIndex        =   34
         Top             =   1080
         Width           =   855
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
         Index           =   11
         Left            =   120
         TabIndex        =   33
         Top             =   180
         Width           =   1005
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
         Index           =   10
         Left            =   120
         TabIndex        =   32
         Top             =   1035
         Width           =   1695
      End
      Begin VB.Label lblLabels 
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
         Height          =   375
         Index           =   9
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Sección :"
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
         Left            =   4560
         TabIndex        =   30
         Top             =   120
         Width           =   1125
      End
   End
   Begin VB.Frame Frame_Base_Datos 
      Caption         =   "Frame_Base_Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   3000
      TabIndex        =   19
      Top             =   2040
      Visible         =   0   'False
      Width           =   2535
      Begin VB.Data DTabla_Gerenar 
         Caption         =   "Tabla_Gerenar"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Tabla_Gerenar"
         Top             =   1680
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Data Data8 
         Caption         =   "Cartas"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Cartas"
         Top             =   960
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.TextBox txtNumero_ID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "txtNumero_ID"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Data Data3 
         Caption         =   "Especialidad"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Especialidad"
         Top             =   240
         Visible         =   0   'False
         Width           =   2340
      End
      Begin VB.Data Data2 
         Caption         =   "Alumnos"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Alumnos"
         Top             =   600
         Visible         =   0   'False
         Width           =   2340
      End
   End
   Begin VB.Frame Frame9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   0
      TabIndex        =   21
      Top             =   1560
      Visible         =   0   'False
      Width           =   6375
      Begin VB.ListBox List_Buscar 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3060
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   6135
      End
      Begin VB.CommandButton cmdCerrar4 
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
         Height          =   555
         Left            =   2520
         TabIndex        =   22
         Top             =   3390
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   4095
      Left            =   0
      TabIndex        =   3
      Top             =   1560
      Width           =   9015
      Begin VB.CommandButton cmdBoton10 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   3360
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   270
         Width           =   375
      End
      Begin VB.ComboBox cobHorario 
         DataField       =   "Horario"
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
         ItemData        =   "frmCentros_Pasantias.frx":1D53
         Left            =   5400
         List            =   "frmCentros_Pasantias.frx":1D5D
         TabIndex        =   2
         Text            =   "Tiempo Completo"
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtHorario 
         DataField       =   "Horario"
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
         Left            =   5400
         TabIndex        =   18
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtFecha_Culminacion 
         Alignment       =   2  'Center
         DataField       =   "Fecha_Culminacion"
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
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   17
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox txtFecha_Inicio 
         Alignment       =   2  'Center
         DataField       =   "Fecha_Inicio"
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
         Left            =   3000
         MaxLength       =   15
         TabIndex        =   16
         Top             =   3120
         Width           =   2895
      End
      Begin VB.TextBox txtDireccion_Emp 
         DataField       =   "Direccion_Emp"
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
         Height          =   885
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   2160
         Width           =   5895
      End
      Begin VB.TextBox txtTelefono_Emp 
         DataField       =   "Telefono_Emp"
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
         Left            =   2520
         TabIndex        =   14
         Top             =   1680
         Width           =   4815
      End
      Begin VB.TextBox txtTutor_Emp 
         DataField       =   "Tutor_Emp"
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
         Left            =   2520
         TabIndex        =   13
         Top             =   1200
         Width           =   4815
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
         Left            =   1440
         TabIndex        =   12
         Top             =   720
         Width           =   7455
      End
      Begin VB.TextBox txtNumero_Oficio 
         DataField       =   "Numero_Oficio"
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
         Left            =   1800
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Caption         =   "N° de Oficio :"
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
         Index           =   1
         Left            =   120
         TabIndex        =   11
         Top             =   195
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Caption         =   "Empresa :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   645
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tutor Empresarial :"
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
         Index           =   3
         Left            =   120
         TabIndex        =   9
         Top             =   1200
         Width           =   2295
      End
      Begin VB.Label lblLabels 
         Caption         =   "Telefono :"
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
         Index           =   4
         Left            =   120
         TabIndex        =   8
         Top             =   1635
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Caption         =   "Dirección :"
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
         Index           =   5
         Left            =   120
         TabIndex        =   7
         Top             =   2085
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Caption         =   "Horario :"
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
         Index           =   6
         Left            =   4200
         TabIndex        =   6
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fecha de Inicio :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   5
         Top             =   3195
         Width           =   2295
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fecha de Culminacion :"
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
         Index           =   8
         Left            =   120
         TabIndex        =   4
         Top             =   3645
         Width           =   3255
      End
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
      RecordSource    =   "Centros_Pasantias"
      Top             =   6795
      Width           =   9120
   End
End
Attribute VB_Name = "frmCentros_Pasantias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
'-- Autor : Kenner Y. Roa B. Telefono (0212) 433.55.14
'-- E-mail : kyrb2000@cantv.net / kyrb2000@hotmail.com
'-- Fecha : 22/05/2002 (Caracas / Venezuela)
'-----------------------------------------------------------

Private Sub cmdAdd_Click()
    Frame1.Enabled = True
    Data1.Recordset.AddNew
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
    cobHorario.Text = "Tiempo Completo"
End Sub

Private Sub cmdBoton10_Click(Index As Integer)
    Frame9.Visible = True
    Frame3.Enabled = False
    Frame1.Enabled = False
    List_Buscar.Clear
    If Not Data8.Recordset.EOF Then
        Data8.Recordset.MoveFirst
        Do Until Data8.Recordset.EOF
            If Not Data8.Recordset.EOF Then
                If Data8.Recordset.Fields("Numero_ID") = Val(txtNumero_ID) And Data8.Recordset.Fields("Codigo_Tipo") = "POS" Then
'                    MsgBox Data8.Recordset.Fields("Codigo_Tipo")
                    List_Buscar.AddItem Data8.Recordset.Fields("Numero") & ", " & Data8.Recordset.Fields("Nombre_Emp")
                    List_Buscar.ItemData(List_Buscar.NewIndex) = Data8.Recordset.Fields("Numero")
                End If
            End If
            Data8.Recordset.MoveNext
        Loop
        Data8.Recordset.MoveFirst
        Do Until Data8.Recordset.EOF
            If Not Data8.Recordset.EOF Then
                If Data8.Recordset.Fields("Numero_ID") = Val(txtNumero_ID) And Data8.Recordset.Fields("Codigo_Tipo") = "PRE" Then
                    List_Buscar.AddItem Data8.Recordset.Fields("Numero") & ", " & Data8.Recordset.Fields("Nombre_Emp") & ", " & Data8.Recordset.Fields("Codigo_Tipo")
                    List_Buscar.ItemData(List_Buscar.NewIndex) = Data8.Recordset.Fields("Numero")
                End If
            End If
            Data8.Recordset.MoveNext
        Loop
        Data8.Recordset.MoveFirst
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim Buscar, Campo_Sel As String
    If (Data1.Recordset.AbsolutePosition + 1) = 0 Then
        Exit Sub
    End If
    '************************************************
    frmBuscar_por.cboBuscar.AddItem "Numero_Oficio"
    frmBuscar_por.cboBuscar.AddItem "Nombre_Emp"
'    frmBuscar_por.cboBuscar.AddItem "Apellidos"
    frmBuscar_por.Show vbModal
    Buscar = frmBuscar_por.txtCampo01
    Campo_Sel = frmBuscar_por.txtCampo02
    Unload frmBuscar_por
    '************************************************
'    Buscar = InputBox("Introduzca La Cedula :", "Busqueda de Datos")
    If Buscar <> "" Then
        If Campo_Sel = "Numero_Oficio" Then
            Buscar = "[" & Campo_Sel & "]" & "=" & Buscar
        Else
            Buscar = "[" & Campo_Sel & "]" & "=" & "'" & Buscar & "'"
        End If
        Data1.Recordset.MoveFirst
        Data1.Recordset.FindFirst Buscar
    End If
End Sub

Private Sub cmdCerrar4_Click()
    Frame9.Visible = False
    Frame3.Enabled = True
    Frame1.Enabled = True
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

Private Sub cmdImprimir_Click()
    Dim fReportes As New frmReportes
    fReportes.Caption = "Reportes Registro del Pasante"
    fReportes.Tipo_Reporte = "REG"
    Call centrarform(fReportes)
End Sub

Private Sub cmdUpdate_Click()
    If cmdUpdate.Caption = "&Actualizar" Then
        ' Si el Control Text esta en blanco
        ' Reeplazalo con un "-"
        Horar = cobHorario.Text
        For i = 0 To Me.Controls.Count - 1
            If TypeOf Me.Controls(i) Is TextBox Then
                If Me.Controls(i).Text = "" Then
                    Me.Controls(i).Text = "-"
                End If
            End If
        Next i
        cobHorario.Text = Horar
        ' ----------------------------------
        Data1.Recordset.Fields("Cedula") = txtCedula
        If txtNumero_Oficio = "-" Then
            txtNumero_Oficio = 0
        End If
        Data1.UpdateRecord
        Data1.Recordset.Bookmark = Data1.Recordset.LastModified
        txtCedula.Locked = True
        '--------- Botones ------------
        cmdUpdate.Caption = "&Modificar"
        cmdAdd.Enabled = True
        cmdDelete.Enabled = True
        cmdBuscar.Enabled = True
        cmdImprimir.Enabled = True
        cmdClose.Caption = "&Cerrar"
        Data1.Enabled = True
        Frame1.Enabled = False
        '------------------------------'
    Else
        If Not Data1.Recordset.EOF Then
            ' Si el Control Text esta en blanco
            ' Reeplazalo con un "-"
            Horar = cobHorario.Text
            For i = 0 To Me.Controls.Count - 1
                If TypeOf Me.Controls(i) Is TextBox Then
                    If Me.Controls(i).Text = "" Then
                        Me.Controls(i).Text = "-"
                    End If
                End If
            Next i
            cobHorario.Text = Horar
            ' ----------------------------------
            Frame1.Enabled = True
            Data1.Recordset.Edit
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
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    If cmdClose.Caption = "&Cerrar" Then
        Unload Me
    Else
        '--------- Botones ------------
        cmdUpdate.Caption = "&Modificar"
        cmdAdd.Enabled = True
        cmdDelete.Enabled = True
        cmdBuscar.Enabled = True
        cmdImprimir.Enabled = True
        cmdClose.Caption = "&Cerrar"
        Data1.Enabled = True
        '------------------------------'
        txtCedula.Locked = True
        Frame1.Enabled = False
        Data1.Recordset.CancelUpdate
    End If
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
  'para el objeto tabla debe establecer la propiedad index cuando
  'se crea el recordset; use la línea siguiente
  'Data1.Caption = "Record: " & (Data1.Recordset.RecordCount * (Data1.Recordset.PercentPosition * 0.01)) + 1
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
    Data1.DatabaseName = Base_de_Datos
    Data1.Refresh
    Data3.DatabaseName = Base_de_Datos
    Data3.Refresh
    Data2.DatabaseName = Base_de_Datos
    Data2.Refresh
    Data8.DatabaseName = Base_de_Datos
    Data8.Refresh
    DTabla_Gerenar.DatabaseName = Base_de_Datos
    DTabla_Gerenar.Refresh
End Sub

Private Sub Form_Load()
    If Me.WindowState <> 2 Then
        Me.Move (Screen.Width - Me.Width) / 2, 0 '(Screen.Height - f.Height) / 2
    End If
End Sub

Private Sub List_Buscar_DblClick()
    txtNumero_Oficio.Text = List_Buscar.ItemData(List_Buscar.ListIndex)
'    Buscar = "[Numero]" & "=" & txtNumero_Oficio
'    Data8.Recordset.FindFirst Buscar
'    If Data8.Recordset.NoMatch Then
'        Exit Sub
'    End If
    txtNombre_Emp = List_Buscar.List(List_Buscar.ListIndex) 'Data8.Recordset.Fields("Nombre_Emp")
    Call cmdCerrar4_Click
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13:
        Call Buscar_Datos
        If Frame1.Enabled = True Then txtNumero_Oficio.SetFocus
    End Select
End Sub

Sub Buscar_Datos()
    Dim Buscar As String
    Dim Seccion As String
    Buscar = "[Cedula]" & "=" & "'" & txtCedula & "'"
    Data2.Recordset.FindFirst Buscar
    If Data2.Recordset.NoMatch Then
        MsgBox "Cedula No Encontrada, Favor Chequear", , "Busqueda de Registro"
        Exit Sub
    End If
    txtCedula = Data2.Recordset.Fields("Cedula")
    txtNumero_ID = Data2.Recordset.Fields("Numero_ID")
    Seccion = Data2.Recordset.Fields("Seccion")
    txtApeNom = Data2.Recordset.Fields("Apellidos") & ", " & Data2.Recordset.Fields("Nombres")
    txtSeccion = Seccion
    Buscar = "[Codigo]" & "=" & "'" & Mid(Seccion, 1, 2) & "'"
    Data3.Recordset.FindFirst Buscar
    txtEspecialidad = Data3.Recordset.Fields("Descripcion")
    Select Case Mid(Seccion, 5, 1)
    Case "1"
        txtTurno = "Mañana"
    Case "3"
        txtTurno = "Noche"
    End Select
End Sub

Private Sub txtDireccion_Emp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtFecha_Inicio.SetFocus
End Sub

Private Sub txtFecha_Culminacion_Click()
    If txtFecha_Culminacion = "" Or txtFecha_Culminacion = "-" Then
        Fecha_Act.XFecha_Actual = Format(Date, Tipo_Fecha)
    Else
        Fecha_Act.XFecha_Actual = txtFecha_Culminacion
    End If
    Fecha_Act.XDia_Min = -2000
    Fecha_Act.XDia_Max = 120
    Fecha_Act.Show vbModal
    txtFecha_Culminacion = Fecha_Act.txtFecha
    Unload Fecha_Act
End Sub

Private Sub txtFecha_Culminacion_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    Call txtFecha_Culminacion_Click
End Sub

Private Sub txtFecha_Inicio_Click()
    If txtFecha_Inicio = "" Or txtFecha_Inicio = "-" Then
        Fecha_Act.XFecha_Actual = Format(Date, Tipo_Fecha)
    Else
        Fecha_Act.XFecha_Actual = txtFecha_Inicio
    End If
    Fecha_Act.XDia_Min = -2000
    Fecha_Act.XDia_Max = 120
    Fecha_Act.Show vbModal
    txtFecha_Inicio = Fecha_Act.txtFecha
    Unload Fecha_Act
End Sub

Private Sub txtFecha_Inicio_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    Call txtFecha_Inicio_Click
End Sub

Private Sub txtNombre_Emp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtTutor_Emp.SetFocus
End Sub

Private Sub txtNumero_Oficio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNombre_Emp.SetFocus
End Sub

Private Sub txtTelefono_Emp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDireccion_Emp.SetFocus
End Sub

Private Sub txtTutor_Emp_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTelefono_Emp.SetFocus
End Sub
