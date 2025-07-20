VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmFicha_Pasantia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ficha de Pasantía"
   ClientHeight    =   7005
   ClientLeft      =   345
   ClientTop       =   750
   ClientWidth     =   11205
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFicha_Pasantia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   11205
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
      Height          =   2175
      Left            =   600
      TabIndex        =   98
      Top             =   1680
      Visible         =   0   'False
      Width           =   5175
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
         Left            =   2640
         Locked          =   -1  'True
         TabIndex        =   99
         Top             =   1320
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.Data DSolvencias 
         Caption         =   "Solvencias"
         Connect         =   "Access"
         DatabaseName    =   ""
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
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Solvencias"
         Top             =   600
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Data DEspecialidad 
         Caption         =   "Especialidad"
         Connect         =   "Access"
         DatabaseName    =   ""
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
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Especialidad"
         Top             =   960
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Data DNotas 
         Caption         =   "Notas"
         Connect         =   "Access"
         DatabaseName    =   ""
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
         Left            =   2640
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Notas"
         Top             =   240
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Data DDocentes 
         Caption         =   "Docentes"
         Connect         =   "Access"
         DatabaseName    =   ""
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
         RecordSource    =   "Docentes"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Data DControl 
         Caption         =   "Control"
         Connect         =   "Access"
         DatabaseName    =   ""
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
         RecordSource    =   "Control"
         Top             =   1680
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Data DCentros_Pasantias 
         Caption         =   "Centros_Pasantias"
         Connect         =   "Access"
         DatabaseName    =   ""
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
         RecordSource    =   "Centros_Pasantias"
         Top             =   960
         Visible         =   0   'False
         Width           =   2460
      End
      Begin VB.Data DAlumnos 
         Caption         =   "Alumnos"
         Connect         =   "Access"
         DatabaseName    =   ""
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
         Width           =   2460
      End
      Begin VB.Data DCartas 
         Caption         =   "Cartas"
         Connect         =   "Access"
         DatabaseName    =   ""
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
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Cartas"
         Top             =   240
         Visible         =   0   'False
         Width           =   2460
      End
   End
   Begin VB.Frame Frame9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   120
      TabIndex        =   93
      Top             =   1680
      Visible         =   0   'False
      Width           =   4695
      Begin VB.CommandButton cmdCerrar4 
         Caption         =   "&Cerrar"
         Height          =   1035
         Left            =   1560
         Picture         =   "frmFicha_Pasantia.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   3030
         Width           =   1335
      End
      Begin VB.ListBox List_Buscar 
         Height          =   2760
         Left            =   120
         TabIndex        =   94
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4695
      Left            =   4560
      TabIndex        =   70
      Top             =   1080
      Visible         =   0   'False
      Width           =   6495
      Begin VB.CommandButton cmdCerrar3 
         Caption         =   "&Cerrar"
         Height          =   1035
         Left            =   4920
         Picture         =   "frmFicha_Pasantia.frx":0869
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton cmdBoton10 
         Height          =   300
         Index           =   0
         Left            =   2775
         MaskColor       =   &H8000000F&
         Picture         =   "frmFicha_Pasantia.frx":0C90
         TabIndex        =   92
         Top             =   165
         Width           =   375
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   1035
         Left            =   3600
         Picture         =   "frmFicha_Pasantia.frx":0FD2
         Style           =   1  'Graphical
         TabIndex        =   91
         Top             =   3600
         Width           =   1335
      End
      Begin VB.ComboBox CmbNumero_Visitas 
         DataField       =   "Numero_Visitas"
         DataSource      =   "DNotas"
         Height          =   420
         ItemData        =   "frmFicha_Pasantia.frx":13FA
         Left            =   2640
         List            =   "frmFicha_Pasantia.frx":1407
         Locked          =   -1  'True
         TabIndex        =   89
         Text            =   "0"
         Top             =   3750
         Width           =   615
      End
      Begin VB.TextBox txtCedula_Tutor 
         DataField       =   "Cedula_Tutor"
         DataSource      =   "DNotas"
         Height          =   420
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   85
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtNombre2 
         Enabled         =   0   'False
         Height          =   420
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   645
         Width           =   4815
      End
      Begin VB.TextBox txtCargo2 
         Enabled         =   0   'False
         Height          =   420
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   83
         Top             =   120
         Width           =   2175
      End
      Begin VB.Frame Frame8 
         Caption         =   "Fechas"
         Enabled         =   0   'False
         Height          =   2415
         Left            =   240
         TabIndex        =   72
         Top             =   1080
         Width           =   6135
         Begin VB.TextBox txtFecha_Final_Proceso 
            Alignment       =   2  'Center
            DataField       =   "Fecha_Final_Proceso"
            DataSource      =   "DNotas"
            Height          =   420
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   82
            Top             =   1845
            Width           =   2895
         End
         Begin VB.TextBox txtFecha_Aceptacion_Informe 
            Alignment       =   2  'Center
            DataField       =   "Fecha_Aceptacion_Informe"
            DataSource      =   "DNotas"
            Height          =   420
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   81
            Top             =   1440
            Width           =   2895
         End
         Begin VB.TextBox txtFecha_Entrega_Informe 
            Alignment       =   2  'Center
            DataField       =   "Fecha_Entrega_Informe"
            DataSource      =   "DNotas"
            Height          =   420
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   80
            Top             =   1050
            Width           =   2895
         End
         Begin VB.TextBox txtFecha_Aceptacion_Plan 
            Alignment       =   2  'Center
            DataField       =   "Fecha_Aceptacion_Plan"
            DataSource      =   "DNotas"
            Height          =   420
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   79
            Top             =   645
            Width           =   2895
         End
         Begin VB.TextBox txtFecha_Entrega_Plan 
            Alignment       =   2  'Center
            DataField       =   "Fecha_Entrega_Plan"
            DataSource      =   "DNotas"
            Height          =   420
            Left            =   3120
            Locked          =   -1  'True
            TabIndex        =   78
            Top             =   240
            Width           =   2895
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Final del Proceso :"
            Height          =   300
            Index           =   0
            Left            =   120
            TabIndex        =   77
            Top             =   1845
            Width           =   2235
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Aceptacion del Informe :"
            Height          =   300
            Index           =   4
            Left            =   120
            TabIndex        =   76
            Top             =   1440
            Width           =   2955
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Entrega del Informe :"
            Height          =   300
            Index           =   5
            Left            =   120
            TabIndex        =   75
            Top             =   1050
            Width           =   2565
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Aceptacion del Plan :"
            Height          =   300
            Index           =   10
            Left            =   120
            TabIndex        =   74
            Top             =   645
            Width           =   2550
         End
         Begin VB.Label lblLabels 
            AutoSize        =   -1  'True
            Caption         =   "Entrega del Plan :"
            Height          =   300
            Index           =   11
            Left            =   120
            TabIndex        =   73
            Top             =   240
            Width           =   2160
         End
      End
      Begin VB.Label Label16 
         Caption         =   "Cedula :"
         Height          =   375
         Left            =   120
         TabIndex        =   88
         Top             =   120
         Width           =   1095
      End
      Begin VB.Label Label15 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   120
         TabIndex        =   87
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label Label14 
         Caption         =   "Cargo :"
         Height          =   375
         Left            =   3240
         TabIndex        =   86
         Top             =   120
         Width           =   975
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Numero de Visitas :"
         Height          =   300
         Index           =   9
         Left            =   240
         TabIndex        =   71
         Top             =   3750
         Width           =   2355
      End
   End
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   4920
      TabIndex        =   44
      Top             =   2160
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdCerrar2 
         Caption         =   "&Cerrar"
         Height          =   915
         Left            =   4560
         Picture         =   "frmFicha_Pasantia.frx":1414
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   2450
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd2 
         Caption         =   "&Agregar"
         Height          =   915
         Left            =   4560
         Picture         =   "frmFicha_Pasantia.frx":183B
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   1540
         Width           =   1335
      End
      Begin VB.CommandButton cmdBoton10 
         Height          =   300
         Index           =   2
         Left            =   2895
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   97
         Top             =   285
         Width           =   255
      End
      Begin VB.TextBox txtFecha_Culminacion 
         DataField       =   "Fecha_Culminacion"
         Height          =   420
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   69
         Top             =   2280
         Width           =   2175
      End
      Begin VB.TextBox txtFecha_Inicio 
         DataField       =   "Fecha_Inicio"
         Height          =   420
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   68
         Top             =   1620
         Width           =   2175
      End
      Begin VB.TextBox txtTutor_Emp 
         DataField       =   "Tutor_Emp"
         Height          =   420
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   67
         Top             =   1080
         Width           =   4215
      End
      Begin VB.TextBox txtNombre_Emp 
         DataField       =   "Nombre_Emp"
         Height          =   420
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   66
         Top             =   660
         Width           =   4215
      End
      Begin VB.TextBox txtHorario 
         DataField       =   "Horario"
         Height          =   420
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox txtNumero_Oficio 
         DataField       =   "Numero_Oficio"
         Height          =   420
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   64
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fecha de Culminacion :"
         Height          =   615
         Index           =   8
         Left            =   120
         TabIndex        =   63
         Top             =   2160
         Width           =   1935
      End
      Begin VB.Label lblLabels 
         Caption         =   "Fecha de Inicio :"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   62
         Top             =   1740
         Width           =   3015
      End
      Begin VB.Label lblLabels 
         Caption         =   "Horario :"
         Height          =   375
         Index           =   6
         Left            =   3240
         TabIndex        =   61
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Tutor Emp. :"
         Height          =   375
         Index           =   3
         Left            =   120
         TabIndex        =   60
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Empresa :"
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   59
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "N° de Oficio :"
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   58
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   4920
      TabIndex        =   45
      Top             =   2160
      Visible         =   0   'False
      Width           =   6015
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   915
         Left            =   4560
         Picture         =   "frmFicha_Pasantia.frx":1C3C
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   2445
         Width           =   1335
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Eliminar"
         Height          =   915
         Left            =   3240
         Picture         =   "frmFicha_Pasantia.frx":2063
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   2445
         Width           =   1335
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Agregar"
         Height          =   915
         Left            =   1920
         Picture         =   "frmFicha_Pasantia.frx":24CC
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   2445
         Width           =   1335
      End
      Begin VB.CommandButton cmdBoton10 
         Height          =   300
         Index           =   1
         Left            =   2535
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   96
         Top             =   270
         Width           =   375
      End
      Begin VB.TextBox txtObservacionX 
         Height          =   1095
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   49
         Top             =   1200
         Width           =   5775
      End
      Begin VB.TextBox txtCargoX 
         Enabled         =   0   'False
         Height          =   420
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   47
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox txtNombreX 
         Enabled         =   0   'False
         Height          =   420
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   48
         Top             =   720
         Width           =   4695
      End
      Begin VB.TextBox txtCedulaX 
         Height          =   420
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   46
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label13 
         Caption         =   "Cargo :"
         Height          =   375
         Left            =   3000
         TabIndex        =   54
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label12 
         Caption         =   "Nombre :"
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label11 
         Caption         =   "Cedula :"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   1095
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   4215
      Left            =   4680
      TabIndex        =   43
      Top             =   1560
      Visible         =   0   'False
      Width           =   6345
      _ExtentX        =   11192
      _ExtentY        =   7435
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Taller"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   23
      Top             =   5760
      Width           =   11055
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cerrar"
         Height          =   1020
         Left            =   9600
         Picture         =   "frmFicha_Pasantia.frx":28CD
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   150
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Actualizar"
         Enabled         =   0   'False
         Height          =   1020
         Left            =   8280
         Picture         =   "frmFicha_Pasantia.frx":2CF4
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   150
         Width           =   1335
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   1020
         Left            =   6960
         Picture         =   "frmFicha_Pasantia.frx":311C
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   150
         Width           =   1335
      End
      Begin VB.TextBox txtCasos 
         BackColor       =   &H00C0FFFF&
         Height          =   975
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   100
         Top             =   160
         Visible         =   0   'False
         Width           =   6735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   5895
      Left            =   0
      TabIndex        =   2
      Top             =   -120
      Width           =   11175
      Begin VB.CommandButton cmdBoton10 
         Height          =   300
         Index           =   3
         Left            =   2550
         MaskColor       =   &H8000000F&
         Picture         =   "frmFicha_Pasantia.frx":3537
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   280
         Width           =   375
      End
      Begin VB.TextBox txtTelefono 
         Height          =   420
         Left            =   5640
         TabIndex        =   36
         Top             =   1200
         Visible         =   0   'False
         Width           =   3735
      End
      Begin VB.TextBox txtPeriodo 
         Enabled         =   0   'False
         Height          =   420
         Left            =   4200
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   210
         Width           =   975
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   420
         Left            =   6240
         Locked          =   -1  'True
         TabIndex        =   33
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtTurno 
         Enabled         =   0   'False
         Height          =   420
         Left            =   9120
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox txtEspecialidad 
         Enabled         =   0   'False
         Height          =   420
         Left            =   1800
         Locked          =   -1  'True
         TabIndex        =   31
         Top             =   1200
         Width           =   3735
      End
      Begin VB.TextBox txtSeccion 
         Alignment       =   2  'Center
         DataField       =   "Seccion"
         Enabled         =   0   'False
         Height          =   420
         Left            =   7920
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtApeNom 
         Enabled         =   0   'False
         Height          =   420
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   720
         Width           =   5175
      End
      Begin VB.TextBox txtCedula 
         DataSource      =   "DAlumnos"
         Height          =   420
         Left            =   1200
         MaxLength       =   9
         TabIndex        =   0
         Top             =   240
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Caption         =   "Solvencias"
         Enabled         =   0   'False
         Height          =   4095
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   10935
         Begin Crystal.CrystalReport CrystalReport2 
            Left            =   2640
            Top             =   240
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.CommandButton cmdBoton04 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   1800
            Width           =   255
         End
         Begin Crystal.CrystalReport CrystalReport1 
            Left            =   2160
            Top             =   240
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            PrintFileLinesPerPage=   60
         End
         Begin VB.CheckBox Check8 
            Caption         =   "Realizando Pasantía"
            Height          =   300
            Left            =   480
            TabIndex        =   18
            Top             =   2880
            Width           =   4215
         End
         Begin VB.CheckBox Check7 
            Caption         =   "Entrego Carta de Aceptación"
            Height          =   300
            Left            =   480
            TabIndex        =   17
            Top             =   2520
            Width           =   4455
         End
         Begin VB.CheckBox Check6 
            Caption         =   "Casos Especiales"
            Height          =   300
            Left            =   480
            TabIndex        =   16
            Top             =   2160
            Width           =   4095
         End
         Begin VB.CheckBox Check4 
            Caption         =   "Academica (Departamento de Grado)"
            Height          =   300
            Left            =   480
            TabIndex        =   14
            Top             =   1440
            Width           =   5055
         End
         Begin VB.CheckBox Check3 
            Caption         =   "Academica (Control de Estudio)"
            Height          =   300
            Left            =   480
            TabIndex        =   13
            Top             =   1080
            Width           =   4815
         End
         Begin VB.CheckBox Check2 
            Caption         =   "Administrativo (Caja)"
            Height          =   300
            Left            =   480
            TabIndex        =   12
            Top             =   720
            Width           =   3855
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Taller"
            Height          =   300
            Left            =   480
            TabIndex        =   11
            Top             =   360
            Width           =   3015
         End
         Begin VB.CommandButton cmdBoton06 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   2880
            Width           =   255
         End
         Begin VB.CommandButton cmdBoton05 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   2160
            Width           =   255
         End
         Begin VB.CommandButton cmdBoton03 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   1440
            Width           =   255
         End
         Begin VB.CommandButton cmdBoton02 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   1080
            Width           =   255
         End
         Begin VB.CommandButton cmdBoton01 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   360
            Width           =   255
         End
         Begin VB.Frame Frame4 
            Caption         =   "Notas"
            Height          =   1575
            Left            =   6360
            TabIndex        =   24
            Top             =   240
            Width           =   3135
            Begin VB.TextBox txtTotal 
               Enabled         =   0   'False
               Height          =   420
               Left            =   1920
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   960
               Width           =   1095
            End
            Begin VB.TextBox txtAcum40 
               Height          =   420
               Left            =   1080
               TabIndex        =   20
               Top             =   960
               Width           =   735
            End
            Begin VB.TextBox txtAcum60 
               Height          =   420
               Left            =   120
               TabIndex        =   19
               Top             =   960
               Width           =   735
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               Caption         =   "Total Acum."
               Height          =   675
               Left            =   1920
               TabIndex        =   27
               Top             =   240
               Width           =   1185
            End
            Begin VB.Label Label9 
               Alignment       =   2  'Center
               Caption         =   "Acum. 40%"
               Height          =   555
               Left            =   1080
               TabIndex        =   26
               Top             =   240
               Width           =   945
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               Caption         =   "Acum. 60%"
               Height          =   675
               Left            =   135
               TabIndex        =   25
               Top             =   240
               Width           =   825
            End
         End
         Begin VB.CheckBox Check5 
            Caption         =   "Notas Entregadas Tutor Academico"
            Height          =   300
            Left            =   480
            TabIndex        =   15
            Top             =   1800
            Width           =   4935
         End
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Periodo :"
         Height          =   300
         Left            =   3120
         TabIndex        =   30
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Fecha :"
         Height          =   300
         Left            =   5280
         TabIndex        =   10
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Turno :"
         Height          =   300
         Left            =   8160
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Especialidad :"
         Height          =   300
         Left            =   120
         TabIndex        =   7
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Sección :"
         Height          =   300
         Left            =   6720
         TabIndex        =   5
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombres :"
         Height          =   300
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1230
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cedula :"
         Height          =   300
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1005
      End
   End
End
Attribute VB_Name = "frmFicha_Pasantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
'-- Autor : Kenner Y. Roa B. Telefono (0212) 433.55.14
'-- E-mail : kyrb2000@cantv.net / kyrb2000@hotmail.com
'-- Fecha : 22/05/2002 (Caracas / Venezuela)
'-----------------------------------------------------------
Dim Buscar_Bot As Integer

Private Sub cmdAdd_Click()
    Dim Buscar, Tipo As String
    '-----------------------------------------------
    Select Case TabStrip1.SelectedItem
    Case "Taller":
        Tipo = "TA"
    Case "Control de Estudio":
        Tipo = "CO"
    Case "Dpto. de Grado":
        Tipo = "GR"
    Case "Casos Especiales":
        Tipo = "CE"
    Case "Pasantías":
        Tipo = "PA"
    End Select
    Buscar = "[Cedula]" & "=" & "'" & txtCedula & Tipo & "'"
    DControl.Recordset.FindFirst Buscar
'--------------------------------------------
            txtCedulaX.Locked = False
            'txtCargoX.Locked = False
            'txtNombreX.Locked = False
            txtObservacionX.Locked = False
'--------------------------------------------
    Select Case cmdAdd.Caption
    Case "&Agregar":
        '-------------------------------
        cmdAdd.Caption = "&Aceptar"
        cmdDelete.Enabled = True
        cmdDelete.Caption = "&Cancelar"
        cmdCerrar.Enabled = False
        cmdAdd.Picture = LoadPicture(App.Path & "\GIF\Guardar.gif")
        cmdDelete.Picture = LoadPicture(App.Path & "\GIF\Eliminar.gif")
        '-------------------------------
        If Not DControl.Recordset.EOF Then
            If DControl.Recordset.Fields("Cedula") = txtCedula & Tipo Then
                txtCedulaX = DControl.Recordset.Fields("Cedula_Doc")
                txtObservacionX = DControl.Recordset.Fields("Observacion")
                Buscar = "[Cedula]" & "=" & "'" & txtCedulaX & "'"
                DDocentes.Recordset.FindFirst Buscar
                If DDocentes.Recordset.NoMatch Then
                    txtNombreX = ""
                    txtCargoX = ""
                Else
                    txtCargoX = DDocentes.Recordset.Fields("Cargo")
                    txtNombreX = Trim(DDocentes.Recordset.Fields("Apellidos")) & ", " & Trim(DDocentes.Recordset.Fields("Nombres"))
                End If
            Else
                txtCedulaX = ""
                txtCargoX = ""
                txtNombreX = ""
                txtObservacionX = ""
            End If
        Else
            txtCedulaX = ""
            txtCargoX = ""
            txtNombreX = ""
            txtObservacionX = ""
        End If
    Case "&Modificar":
        '-------------------------------
        cmdAdd.Caption = "&Aceptar"
        cmdDelete.Enabled = True
        cmdDelete.Caption = "&Cancelar"
        cmdCerrar.Enabled = False
        cmdAdd.Picture = LoadPicture(App.Path & "\GIF\Guardar.gif")
        cmdDelete.Picture = LoadPicture(App.Path & "\GIF\Eliminar.gif")
        '-------------------------------
        If DControl.Recordset.Fields("Cedula") = txtCedula & Tipo Then
            txtCedulaX = DControl.Recordset.Fields("Cedula_Doc")
            txtObservacionX = DControl.Recordset.Fields("Observacion")
            Buscar = "[Cedula]" & "=" & "'" & txtCedulaX & "'"
            DDocentes.Recordset.FindFirst Buscar
            If DDocentes.Recordset.NoMatch Then
                txtNombreX = ""
                txtCargoX = ""
            Else
                txtCargoX = DDocentes.Recordset.Fields("Cargo")
                txtNombreX = Trim(DDocentes.Recordset.Fields("Apellidos")) & ", " & Trim(DDocentes.Recordset.Fields("Nombres"))
            End If
        End If
    Case "&Aceptar"
        If txtCedulaX = "" Then
            Call TabStrip1_Click
            Exit Sub
        Else
            Buscar = "[Cedula]" & "=" & "'" & txtCedulaX & "'"
            DDocentes.Recordset.FindFirst Buscar
            If DDocentes.Recordset.NoMatch Then
                txtNombreX = ""
                txtCargoX = ""
                Call TabStrip1_Click
                Exit Sub
            Else
                txtCargoX = DDocentes.Recordset.Fields("Cargo")
                txtNombreX = Trim(DDocentes.Recordset.Fields("Apellidos")) & ", " & Trim(DDocentes.Recordset.Fields("Nombres"))
            End If
        End If
        If txtObservacionX = "" Then
            txtObservacionX = "-"
        End If
        '-------------------------------
        cmdAdd.Caption = "&Modificar"
        cmdDelete.Enabled = True
        cmdDelete.Caption = "&Eliminar"
        cmdCerrar.Enabled = True
        cmdAdd.Picture = LoadPicture(App.Path & "\GIF\Modificar.gif")
        cmdDelete.Picture = LoadPicture(App.Path & "\GIF\Eliminar.gif")
        '-------------------------------
        If Not DControl.Recordset.EOF Then
            If DControl.Recordset.Fields("Cedula") = txtCedula & Tipo Then
                DControl.Recordset.Edit
            Else
                DControl.Recordset.AddNew
            End If
        Else
            DControl.Recordset.AddNew
        End If
        DControl.Recordset.Fields("Cedula") = txtCedula & Tipo
        DControl.Recordset.Fields("Cedula_Doc") = txtCedulaX
        DControl.Recordset.Fields("Observacion") = txtObservacionX
        DControl.Recordset.Update
'--------------------------------------------
            txtCedulaX.Locked = True
            txtObservacionX.Locked = True
'--------------------------------------------
    End Select
End Sub

Private Sub cmdAdd2_Click()
    Dim Buscar, Tipo As String
    '-----------------------------------------------
    Buscar = "[Cedula]" & "=" & "'" & txtCedula & "'"
    DDocentes.Recordset.FindFirst Buscar
    Select Case cmdAdd2.Caption
    Case "&Agregar":
        cmdAdd2.Caption = "&Consultar"
        cmdAdd2.Picture = LoadPicture(App.Path & "\GIF\Buscar.gif")
        If Not DControl.Recordset.EOF Then
            frmCentros_Pasantias.cmdAdd.Value = True
            frmCentros_Pasantias.txtCedula = txtCedula
            frmCentros_Pasantias.Show
        Else
        End If
    Case "&Consultar":
        cmdAdd2.Caption = "&Consultar"
        cmdAdd2.Picture = LoadPicture(App.Path & "\GIF\Buscar.gif")
        Buscar = txtCedula
        Buscar = "[Cedula]" & "=" & "'" & Buscar & "'"
        frmCentros_Pasantias.Data1.Recordset.FindFirst Buscar
        frmCentros_Pasantias.Show
    End Select
End Sub

Private Sub cmdBoton01_Click()
    TabStrip1.SelectedItem = "Taller":
    Call Bloque_Boton("TA")
End Sub

Private Sub cmdBoton02_Click()
    TabStrip1.SelectedItem = "Control de Estudio":
    Call Bloque_Boton("CO")
End Sub

Private Sub cmdBoton03_Click()
    TabStrip1.SelectedItem = "Dpto. de Grado":
    Call Bloque_Boton("GR")
End Sub

Private Sub cmdBoton04_Click()
    '---------------------------
    TabStrip1.Visible = True
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame7.Visible = True
    Frame3.Enabled = False
    '---------------------------
    If txtCedula_Tutor <> "" Then
        Buscar = "[Cedula]" & "=" & "'" & txtCedula_Tutor & "'"
        DDocentes.Recordset.FindFirst Buscar
        If DDocentes.Recordset.NoMatch Then
            'MsgBox "Cedula No Encontrada, Favor Chequear", , "Busqueda de Registro"
            txtNombre2 = ""
            txtCargo2 = ""
        Else
            txtCargo2 = DDocentes.Recordset.Fields("Cargo")
            txtNombre2 = Trim(DDocentes.Recordset.Fields("Apellidos")) & ", " & Trim(DDocentes.Recordset.Fields("Nombres"))
        End If
    End If
End Sub

Private Sub cmdBoton05_Click()
    TabStrip1.SelectedItem = "Casos Especiales":
    Call Bloque_Boton("CE")
End Sub

Private Sub cmdBoton06_Click()
    Dim Buscar As String
    '-----------------------------------------------
    DCentros_Pasantias.Refresh
    Buscar = "[Cedula]" & "=" & "'" & txtCedula & "'"
    DCentros_Pasantias.Recordset.FindFirst Buscar
    If Not DCentros_Pasantias.Recordset.EOF Then
        If DCentros_Pasantias.Recordset.Fields("Cedula") = txtCedula Then
            txtNumero_Oficio = DCentros_Pasantias.Recordset.Fields("Numero_Oficio")
            txtHorario = DCentros_Pasantias.Recordset.Fields("Horario")
            txtNombre_Emp = DCentros_Pasantias.Recordset.Fields("Nombre_Emp")
            txtTutor_Emp = DCentros_Pasantias.Recordset.Fields("Tutor_Emp")
            txtFecha_Inicio = DCentros_Pasantias.Recordset.Fields("Fecha_Inicio")
            txtFecha_Culminacion = DCentros_Pasantias.Recordset.Fields("Fecha_Culminacion")
            '-------------------------------
            cmdAdd2.Caption = "&Consultar"
            cmdAdd2.Picture = LoadPicture(App.Path & "\GIF\Buscar.gif")
            '-------------------------------
        Else
            '-------------------------------
            cmdAdd2.Caption = "&Agregar"
            cmdAdd2.Picture = LoadPicture(App.Path & "\GIF\Agregar.gif")
            '-------------------------------
            txtNumero_Oficio = ""
            txtHorario = ""
            txtNombre_Emp = ""
            txtTutor_Emp = ""
            txtFecha_Inicio = ""
            txtFecha_Culminacion = ""
        End If
    Else
        '-------------------------------
        cmdAdd2.Caption = "&Agregar"
        cmdAdd2.Picture = LoadPicture(App.Path & "\GIF\Agregar.bmp")
        '-------------------------------
        txtNumero_Oficio = ""
        txtHorario = ""
        txtNombre_Emp = ""
        txtTutor_Emp = ""
        txtFecha_Inicio = ""
        txtFecha_Culminacion = ""
    End If
    TabStrip1.SelectedItem = "Pasantías":
    '---------------------------
    TabStrip1.Visible = True
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame5.Visible = False
    Frame6.Visible = True
    Frame3.Enabled = False
    '---------------------------
End Sub

Sub Bloque_Boton(Tipo As String)
    Dim Buscar As String
    '-----------------------------------------------
    Buscar = "[Cedula]" & "=" & "'" & txtCedula & Tipo & "'"
    DControl.Recordset.FindFirst Buscar
    If Not DControl.Recordset.EOF Then
        If DControl.Recordset.Fields("Cedula") = txtCedula & Tipo Then
            txtCedulaX = DControl.Recordset.Fields("Cedula_Doc")
            txtObservacionX = DControl.Recordset.Fields("Observacion")
            Buscar = "[Cedula]" & "=" & "'" & txtCedulaX & "'"
            DDocentes.Recordset.FindFirst Buscar
            If DDocentes.Recordset.NoMatch Then
                txtNombreX = ""
                txtCargoX = ""
            Else
                txtCargoX = DDocentes.Recordset.Fields("Cargo")
                txtNombreX = Trim(DDocentes.Recordset.Fields("Apellidos")) & ", " & Trim(DDocentes.Recordset.Fields("Nombres"))
            End If
            '-------------------------------
            cmdAdd.Caption = "&Modificar"
            cmdDelete.Enabled = True
            cmdDelete.Caption = "&Eliminar"
            cmdCerrar.Enabled = True
            cmdAdd.Picture = LoadPicture(App.Path & "\GIF\Modificar.gif")
            cmdDelete.Picture = LoadPicture(App.Path & "\GIF\Eliminar.gif")
            '-------------------------------
        Else
            '-------------------------------
            cmdAdd.Caption = "&Agregar"
            cmdDelete.Enabled = False
            cmdDelete.Caption = "&Eliminar"
            cmdCerrar.Enabled = True
            cmdAdd.Picture = LoadPicture(App.Path & "\GIF\Agregar.gif")
            cmdDelete.Picture = LoadPicture(App.Path & "\GIF\Eliminar.gif")
            '-------------------------------
            txtCedulaX = ""
            txtCargoX = ""
            txtNombreX = ""
            txtObservacionX = ""
        End If
    Else
        '-------------------------------
        cmdAdd.Caption = "&Agregar"
        cmdDelete.Enabled = False
        cmdDelete.Caption = "&Eliminar"
        cmdCerrar.Enabled = True
        cmdAdd.Picture = LoadPicture(App.Path & "\GIF\Agregar.gif")
        cmdDelete.Picture = LoadPicture(App.Path & "\GIF\Eliminar.gif")
        '-------------------------------
        txtCedulaX = ""
        txtCargoX = ""
        txtNombreX = ""
        txtObservacionX = ""
    End If
    '---------------------------
    TabStrip1.Visible = True
    Frame1.Enabled = False
    Frame2.Enabled = False
    Frame5.Visible = True
    Frame6.Visible = False
    Frame3.Enabled = False
    '---------------------------
    txtCedulaX.Locked = True
    txtCargoX.Locked = True
    txtNombreX.Locked = True
    txtObservacionX.Locked = True
    '---------------------------
End Sub

Private Sub cmdBoton10_Click(Index As Integer)
    Buscar_Bot = Index
    Select Case Buscar_Bot
    Case 0:
        Frame9.Visible = True
        Frame7.Enabled = False
        List_Buscar.Clear
        If Not DDocentes.Recordset.EOF Then
            DDocentes.Recordset.MoveFirst
            Do Until DDocentes.Recordset.EOF
                If Not DDocentes.Recordset.EOF Then
                    List_Buscar.AddItem DDocentes.Recordset.Fields("Cedula") & ", " & DDocentes.Recordset.Fields("Apellidos") & ", " & DDocentes.Recordset.Fields("Nombres")
                    List_Buscar.ItemData(List_Buscar.NewIndex) = DDocentes.Recordset.Fields("Cedula")
                End If
                DDocentes.Recordset.MoveNext
            Loop
            DDocentes.Recordset.MoveFirst
        End If
    Case 1:
        Frame9.Visible = True
        Frame5.Enabled = False
        List_Buscar.Clear
        If Not DDocentes.Recordset.EOF Then
            DDocentes.Recordset.MoveFirst
            Do Until DDocentes.Recordset.EOF
                If Not DDocentes.Recordset.EOF Then
                    List_Buscar.AddItem DDocentes.Recordset.Fields("Cedula") & ", " & DDocentes.Recordset.Fields("Apellidos") & ", " & DDocentes.Recordset.Fields("Nombres")
                    List_Buscar.ItemData(List_Buscar.NewIndex) = DDocentes.Recordset.Fields("Cedula")
                End If
                DDocentes.Recordset.MoveNext
            Loop
            DDocentes.Recordset.MoveFirst
        End If
    Case 2:
        Frame9.Visible = True
        Frame6.Enabled = False
        List_Buscar.Clear
        If Not DCartas.Recordset.EOF Then
            DCartas.Recordset.MoveFirst
            Do Until DCartas.Recordset.EOF
                If Not DCartas.Recordset.EOF Then
                    If DCartas.Recordset.Fields("Numero_ID") = Val(txtNumero_ID) Then
                        List_Buscar.AddItem DCartas.Recordset.Fields("Numero") & ", " & DCartas.Recordset.Fields("Nombre_Emp")
                        List_Buscar.ItemData(List_Buscar.NewIndex) = DCartas.Recordset.Fields("Numero")
                    End If
                End If
                DCartas.Recordset.MoveNext
            Loop
            DCartas.Recordset.MoveFirst
        End If
    Case 3:
        Buscar = txtCedula
        Buscar = "[Cedula]" & "=" & "'" & Buscar & "'"
        frmAlumnos.mdbAlumnos.Recordset.Find Buscar
        frmAlumnos.Show
        'Call centrarform(frmAlumnos)
    End Select
End Sub

Private Sub cmdCerrar_Click()
    '---------------------------
    TabStrip1.Visible = False
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame5.Visible = False
    Frame6.Visible = False
    Frame3.Enabled = True
    cmdClose.Picture = LoadPicture(App.Path & "\GIF\Cerrar.gif")
    '---------------------------
End Sub

Private Sub cmdCerrar2_Click()
    '---------------------------
    TabStrip1.Visible = False
    Frame1.Enabled = True
    Frame2.Enabled = True
    Frame5.Visible = False
    Frame6.Visible = False
    Frame3.Enabled = True
    '---------------------------
End Sub

Private Sub cmdCerrar3_Click()
    If cmdCerrar3.Caption = "&Cerrar" Then
        TabStrip1.Visible = False
        Frame1.Enabled = True
        Frame2.Enabled = True
        Frame7.Visible = False
        Frame3.Enabled = True
        '---------------------------
    Else
        Frame8.Enabled = False
        txtCedula_Tutor.Locked = True
        txtFecha_Entrega_Plan.Locked = True
        txtFecha_Aceptacion_Plan.Locked = True
        txtFecha_Entrega_Informe.Locked = True
        txtFecha_Aceptacion_Informe.Locked = True
        txtFecha_Final_Proceso.Locked = True
        CmbNumero_Visitas.Locked = True
        '-----------------------------
        DNotas.Recordset.CancelUpdate
    End If
    '--------- Botones ------------
    cmdModificar.Caption = "&Modificar"
    cmdCerrar3.Caption = "&Cerrar"
    cmdModificar.Picture = LoadPicture(App.Path & "\GIF\Modificar.gif")
    cmdCerrar3.Picture = LoadPicture(App.Path & "\GIF\Cerrar.gif")
    '-----------------------------
End Sub

Private Sub cmdCerrar4_Click()
    Frame9.Visible = False
    Select Case Buscar_Bot
    Case 0:
        Frame7.Enabled = True
    Case 1:
        Frame5.Enabled = True
    Case 2:
        Frame6.Enabled = True
    End Select
End Sub

Private Sub cmdClose_Click()
    cmdClose.Picture = LoadPicture(App.Path & "\GIF\Cerrar.gif")
    If cmdClose.Caption = "&Cerrar" Then
        Unload Me
    Else
        '--------- Botones ------------'
        cmdClose.Caption = "&Cerrar"
        cmdClose.SetFocus
        cmdClose.Picture = LoadPicture(App.Path & "\GIF\Cerrar.gif")
        '------------------------------'
        txtCedula = ""
        txtNumero_ID = ""
        txtApeNom = ""
        txtPeriodo = ""
        txtSeccion = ""
        txtEspecialidad = ""
        txtTurno = ""
        Check1.Value = 0
        Check2.Value = 0
        Check3.Value = 0
        Check4.Value = 0
        Check5.Value = 0
        Check6.Value = 0
        Check7.Value = 0
        Check8.Value = 0
        cmdBoton01.Picture = LoadPicture("")
        cmdBoton02.Picture = LoadPicture("")
        cmdBoton03.Picture = LoadPicture("")
        cmdBoton04.Picture = LoadPicture("")
        cmdBoton05.Picture = LoadPicture("")
        cmdBoton06.Picture = LoadPicture("")
        txtAcum60 = 0
        txtAcum40 = 0
        txtTotal = 0
        cmdUpdate.Enabled = False
        Frame2.Enabled = False
        txtCasos = ""
        txtCasos.Visible = False
    End If
End Sub

Private Sub cmdDelete_Click()
    Dim Respuesta
        Select Case TabStrip1.SelectedItem
    Case "Taller":
        Tipo = "TA"
    Case "Control de Estudio":
        Tipo = "CO"
    Case "Dpto. de Grado":
        Tipo = "GR"
    Case "Casos Especiales":
        Tipo = "CE"
    Case "Pasantías":
        Tipo = "PA"
    End Select
    Buscar = "[Cedula]" & "=" & "'" & txtCedula & Tipo & "'"
    DControl.Recordset.FindFirst Buscar
    Select Case cmdDelete.Caption
    Case "&Eliminar":
        If Not DControl.Recordset.EOF Then
            'esto puede producir un error si elimina el último
            'registro o el único registro del recordset
            Respuesta = MsgBox("Desea Eliminar el Registro", vbYesNo + vbCritical, "Eliminación de Datos")
            If Respuesta = vbYes Then
                DControl.Recordset.Delete
                DControl.Recordset.MoveFirst
                Call TabStrip1_Click
            End If
        End If
    Case "&Cancelar":
        '-------------------------------
        cmdAdd.Caption = "&Modificar"
        cmdDelete.Caption = "&Eliminar"
        cmdCerrar.Enabled = True
        cmdAdd.Picture = LoadPicture(App.Path & "\GIF\Modificar.gif")
        cmdDelete.Picture = LoadPicture(App.Path & "\GIF\Eliminar.gif")
        '-------------------------------
        Select Case TabStrip1.SelectedItem
        Case "Taller":
            Call cmdBoton01_Click
        Case "Control de Estudio":
            Call cmdBoton02_Click
        Case "Dpto. de Grado":
            Call cmdBoton03_Click
        Case "Casos Especiales":
            Call cmdBoton05_Click
        Case "Pasantías":
            Call cmdBoton06_Click
        End Select
    End Select
End Sub

Private Sub cmdImprimir_Click()
    Dim fReportes As New frmReportes
    fReportes.Caption = "Reportes Fichas de Pasantía"
    fReportes.Tipo_Reporte = "FIC"
    Call centrarform(fReportes)
End Sub

Private Sub cmdModificar_Click()
    Select Case cmdModificar.Caption
    Case "&Modificar"
        cmdModificar.Caption = "&Actualizar"
        cmdCerrar3.Caption = "&Cancelar"
        cmdModificar.Picture = LoadPicture(App.Path & "\GIF\Guardar.gif")
        cmdCerrar3.Picture = LoadPicture(App.Path & "\GIF\Eliminar.gif")
        '-----------------------------
        Frame8.Enabled = True
        txtCedula_Tutor.Locked = False
        txtFecha_Entrega_Plan.Locked = False
        txtFecha_Aceptacion_Plan.Locked = False
        txtFecha_Entrega_Informe.Locked = False
        txtFecha_Aceptacion_Informe.Locked = False
        txtFecha_Final_Proceso.Locked = False
        CmbNumero_Visitas.Locked = False
        '-----------------------------
        DNotas.Recordset.Edit
    Case "&Actualizar"
        cmdModificar.Caption = "&Modificar"
        cmdCerrar3.Caption = "&Cerrar"
        cmdModificar.Picture = LoadPicture(App.Path & "\GIF\Modificar.gif")
        cmdCerrar3.Picture = LoadPicture(App.Path & "\GIF\Cerrar.gif")
        '-----------------------------
        Frame8.Enabled = False
        txtCedula_Tutor.Locked = True
        txtFecha_Entrega_Plan.Locked = True
        txtFecha_Aceptacion_Plan.Locked = True
        txtFecha_Entrega_Informe.Locked = True
        txtFecha_Aceptacion_Informe.Locked = True
        txtFecha_Final_Proceso.Locked = True
        CmbNumero_Visitas.Locked = True
        '-----------------------------
        If txtFecha_Entrega_Plan = "" Then
            txtFecha_Entrega_Plan = "-"
        End If
        If txtFecha_Aceptacion_Plan = "" Then
            txtFecha_Aceptacion_Plan = "-"
        End If
        If txtFecha_Entrega_Informe = "" Then
            txtFecha_Entrega_Informe = "-"
        End If
        If txtFecha_Aceptacion_Informe = "" Then
            txtFecha_Aceptacion_Informe = "-"
        End If
        If txtFecha_Final_Proceso = "" Then
            txtFecha_Final_Proceso = "-"
        End If
        '-----------------------------
        DNotas.UpdateRecord
        DNotas.Recordset.Bookmark = DNotas.Recordset.LastModified
    End Select
End Sub

Private Sub cmdUpdate_Click()
    DSolvencias.Recordset.Edit
    DSolvencias.Recordset.Fields("Cedula") = txtCedula
    DSolvencias.Recordset.Fields("Taller") = Check1.Value
    DSolvencias.Recordset.Fields("Administrativo_Caja") = Check2.Value
    DSolvencias.Recordset.Fields("Academica_Control_de_Estudio") = Check3.Value
    DSolvencias.Recordset.Fields("Academica_Departamento_de_Grado") = Check4.Value
    DSolvencias.Recordset.Fields("Notas_Entregadas") = Check5.Value
    DSolvencias.Recordset.Fields("Casos_Especiales") = Check6.Value
    DSolvencias.Recordset.Fields("Entrego_Carta_Aceptacion") = Check7.Value
    DSolvencias.Recordset.Fields("Realizando_Pasantia") = Check8.Value
    DSolvencias.UpdateRecord
    DNotas.Recordset.Edit
    DNotas.Recordset.Fields("Cedula") = txtCedula
    DNotas.Recordset.Fields("Acum_60") = txtAcum60
    DNotas.Recordset.Fields("Acum_40") = txtAcum40
    DNotas.UpdateRecord
End Sub

Private Sub Form_Activate()
    DCartas.DatabaseName = Base_de_Datos
    DCartas.Refresh
    DAlumnos.DatabaseName = Base_de_Datos
    DAlumnos.Refresh
    DCentros_Pasantias.DatabaseName = Base_de_Datos
    DCentros_Pasantias.Refresh
    DDocentes.DatabaseName = Base_de_Datos
    DDocentes.Refresh
    DControl.DatabaseName = Base_de_Datos
    DControl.Refresh
    DNotas.DatabaseName = Base_de_Datos
    DNotas.Refresh
    DSolvencias.DatabaseName = Base_de_Datos
    DSolvencias.Refresh
    DEspecialidad.DatabaseName = Base_de_Datos
    DEspecialidad.Refresh
End Sub

Private Sub Form_Load()
    txtFecha = Format(Date, "dd/mm/yyyy")
End Sub

Private Sub List_Buscar_DblClick()
    Select Case Buscar_Bot
    Case 0:
        txtCedula_Tutor.Text = List_Buscar.ItemData(List_Buscar.ListIndex)
    Case 1:
        txtCedulaX.Text = List_Buscar.ItemData(List_Buscar.ListIndex)
    Case 2:
        txtNumero_Oficio.Text = List_Buscar.ItemData(List_Buscar.ListIndex)
    End Select
    Call cmdCerrar4_Click
End Sub

Private Sub TabStrip1_Click()
    Select Case TabStrip1.SelectedItem
    Case "Taller":
        Call cmdBoton01_Click
    Case "Control de Estudio":
        Call cmdBoton02_Click
    Case "Dpto. de Grado":
        Call cmdBoton03_Click
    Case "Casos Especiales":
        Call cmdBoton05_Click
    Case "Pasantías":
        Call cmdBoton06_Click
    End Select
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case 13:
        txtCedula_LostFocus
        Call Buscar_Casos
    End Select
End Sub

Private Sub txtCedula_LostFocus()
    Dim Buscar As String
    cmdClose.Picture = LoadPicture(App.Path & "\GIF\Cerrar.gif")
    If (DAlumnos.Recordset.AbsolutePosition + 1) = 0 Or txtCedula = "" Then
        Exit Sub
    End If
    DAlumnos.Refresh
    Buscar = "[Cedula]" & "=" & "'" & txtCedula & "'"
    DAlumnos.Recordset.FindFirst Buscar
    If DAlumnos.Recordset.NoMatch Then
        txtApeNom = ""
        txtNumero_ID = ""
        txtPeriodo = ""
        txtSeccion = ""
        txtEspecialidad = ""
        txtTurno = ""
        txtTelefono = ""
        Check1.Value = 0
        Check2.Value = 0
        Check3.Value = 0
        Check4.Value = 0
        Check5.Value = 0
        Check6.Value = 0
        Check7.Value = 0
        Check8.Value = 0
        cmdBoton01.Picture = LoadPicture("")
        cmdBoton02.Picture = LoadPicture("")
        cmdBoton03.Picture = LoadPicture("")
        cmdBoton04.Picture = LoadPicture("")
        cmdBoton05.Picture = LoadPicture("")
        cmdBoton06.Picture = LoadPicture("")
        txtAcum60 = 0
        txtAcum40 = 0
        txtTotal = 0
        
        '---------- Boton -------------
        cmdUpdate.Enabled = False
        'cmdImprimir.Enabled = False
        '------------------------------
        Frame2.Enabled = False
    Else
        txtNumero_ID = DAlumnos.Recordset.Fields("Numero_ID")
        txtApeNom = DAlumnos.Recordset.Fields("Apellidos") & ", " & DAlumnos.Recordset.Fields("Nombres")
        txtTelefono = IIf(IsNull(DAlumnos.Recordset.Fields("Telefono")), "-", DAlumnos.Recordset.Fields("Telefono"))
        txtPeriodo = DAlumnos.Recordset.Fields("Periodo")
        txtSeccion = DAlumnos.Recordset.Fields("Seccion")
        Buscar = "[Codigo]" & "=" & "'" & Mid(txtSeccion, 1, 2) & "'"
        DEspecialidad.Recordset.FindFirst Buscar
        txtEspecialidad = DEspecialidad.Recordset.Fields("Descripcion")
        Select Case Mid(txtSeccion, 5, 1)
        Case "1"
            txtTurno = "Mañana"
        Case "3"
            txtTurno = "Noche"
        End Select
        Buscar = "[Cedula]" & "=" & "'" & txtCedula & "'"
        DSolvencias.Recordset.FindFirst Buscar
        If DSolvencias.Recordset.NoMatch Then
            'MsgBox "No encontrado y es Añadido", , "Busqueda de Registro"
            Check1.Value = 0
            Check2.Value = 0
            Check3.Value = 0
            Check4.Value = 0
            Check5.Value = 0
            Check6.Value = 0
            Check7.Value = 0
            Check8.Value = 0
            cmdBoton01.Picture = LoadPicture("")
            cmdBoton02.Picture = LoadPicture("")
            cmdBoton03.Picture = LoadPicture("")
            cmdBoton04.Picture = LoadPicture("")
            cmdBoton05.Picture = LoadPicture("")
            cmdBoton06.Picture = LoadPicture("")
            txtAcum60 = 0
            txtAcum40 = 0
            txtTotal = 0
            DSolvencias.Recordset.AddNew
            DSolvencias.Recordset.Fields("Cedula") = txtCedula
            DSolvencias.Recordset.Fields("Taller") = Check1.Value
            DSolvencias.Recordset.Fields("Administrativo_Caja") = Check2.Value
            DSolvencias.Recordset.Fields("Academica_Control_de_Estudio") = Check3.Value
            DSolvencias.Recordset.Fields("Academica_Departamento_de_Grado") = Check4.Value
            DSolvencias.Recordset.Fields("Notas_Entregadas") = Check5.Value
            DSolvencias.Recordset.Fields("Casos_Especiales") = Check6.Value
            DSolvencias.Recordset.Fields("Entrego_Carta_Aceptacion") = Check7.Value
            DSolvencias.Recordset.Fields("Realizando_Pasantia") = Check8.Value
            DSolvencias.UpdateRecord
            DNotas.Recordset.AddNew
            DNotas.Recordset.Fields("Cedula") = txtCedula
            DNotas.Recordset.Fields("Acum_60") = txtAcum60
            DNotas.Recordset.Fields("Acum_40") = txtAcum40
            DNotas.UpdateRecord
        Else
            Check1.Value = IIf(DSolvencias.Recordset.Fields("Taller") = True, 1, 0)
            Check2.Value = IIf(DSolvencias.Recordset.Fields("Administrativo_Caja") = True, 1, 0)
            Check3.Value = IIf(DSolvencias.Recordset.Fields("Academica_Control_de_Estudio") = True, 1, 0)
            Check4.Value = IIf(DSolvencias.Recordset.Fields("Academica_Departamento_de_Grado") = True, 1, 0)
            Check5.Value = IIf(DSolvencias.Recordset.Fields("Notas_Entregadas") = True, 1, 0)
            Check6.Value = IIf(DSolvencias.Recordset.Fields("Casos_Especiales") = True, 1, 0)
            Check7.Value = IIf(DSolvencias.Recordset.Fields("Entrego_Carta_Aceptacion") = True, 1, 0)
            Check8.Value = IIf(DSolvencias.Recordset.Fields("Realizando_Pasantia") = True, 1, 0)
            Buscar = "[Cedula]" & "=" & "'" & txtCedula & "'"
            DNotas.Recordset.FindFirst Buscar
            If DNotas.Recordset.NoMatch Then
                'MsgBox "No encontrado y es Añadido", , "Busqueda de Registro"
                DNotas.Recordset.AddNew
                DNotas.Recordset.Fields("Cedula") = txtCedula
                DNotas.Recordset.Fields("Acum_60") = 0
                DNotas.Recordset.Fields("Acum_40") = 0
                DNotas.UpdateRecord
                txtAcum60 = 0
                txtAcum40 = 0
            Else
                txtAcum60 = DNotas.Recordset.Fields("Acum_60")
                txtAcum40 = DNotas.Recordset.Fields("Acum_40")
                If IIf(txtCedula_Tutor <> "", "1", txtCedula_Tutor) <> "1" Then
                    Buscar = "[Cedula]" & "=" & "'" & txtCedula_Tutor & "'"
                    DDocentes.Recordset.FindFirst Buscar
                    If DDocentes.Recordset.NoMatch Then
                        'MsgBox "Cedula No Encontrada, Favor Chequear", , "Busqueda de Registro"
                        txtNombre2 = ""
                        txtCargo2 = ""
                    Else
                        txtCargo2 = DDocentes.Recordset.Fields("Cargo")
                        txtNombre2 = Trim(DDocentes.Recordset.Fields("Apellidos")) & ", " & Trim(DDocentes.Recordset.Fields("Nombres"))
                    End If
                End If
            End If
            txtTotal = Val(txtAcum60) + Val(txtAcum40)
            '---------- Boton -------------
            cmdUpdate.Enabled = True
            'cmdImprimir.Enabled = True
            '------------------------------
        End If
        Frame2.Enabled = True
        cmdClose.Caption = "&Cancelar"
        cmdClose.Picture = LoadPicture(App.Path & "\GIF\Eliminar.gif")
    End If
End Sub

Private Sub txtCedula_Tutor_KeyPress(KeyAscii As Integer)
    Dim Buscar As String
    If KeyAscii = 13 Then
        Buscar = "[Cedula]" & "=" & "'" & txtCedula_Tutor & "'"
        DDocentes.Recordset.FindFirst Buscar
        If DDocentes.Recordset.NoMatch Then
            MsgBox "Cedula No Encontrada, Favor Chequear", , "Busqueda de Registro"
            txtNombre2 = ""
            txtCargo2 = ""
        Else
            txtCargo2 = DDocentes.Recordset.Fields("Cargo")
            txtNombre2 = Trim(DDocentes.Recordset.Fields("Apellidos")) & ", " & Trim(DDocentes.Recordset.Fields("Nombres"))
        End If
    End If
End Sub

Private Sub txtCedula_Tutor_LostFocus()
    Dim Buscar As String
    Buscar = "[Cedula]" & "=" & "'" & txtCedula_Tutor & "'"
    DDocentes.Recordset.FindFirst Buscar
    If DDocentes.Recordset.NoMatch Then
        MsgBox "Cedula No Encontrada, Favor Chequear", , "Busqueda de Registro"
        txtNombre2 = ""
        txtCargo2 = ""
    Else
        txtCargo2 = DDocentes.Recordset.Fields("Cargo")
        txtNombre2 = Trim(DDocentes.Recordset.Fields("Apellidos")) & ", " & Trim(DDocentes.Recordset.Fields("Nombres"))
    End If
End Sub

Private Sub txtCedulaX_KeyPress(KeyAscii As Integer)
    Dim Buscar As String
    If KeyAscii = 13 Then
        Buscar = "[Cedula]" & "=" & "'" & txtCedulaX & "'"
        DDocentes.Recordset.FindFirst Buscar
        If DDocentes.Recordset.NoMatch Then
            MsgBox "Cedula No Encontrada, Favor Chequear", , "Busqueda de Registro"
            txtNombreX = ""
            txtCargoX = ""
        Else
            txtCargoX = DDocentes.Recordset.Fields("Cargo")
            txtNombreX = Trim(DDocentes.Recordset.Fields("Apellidos")) & ", " & Trim(DDocentes.Recordset.Fields("Nombres"))
        End If
    End If
End Sub

Private Sub txtCedulaX_LostFocus()
    Dim Buscar As String
    Buscar = "[Cedula]" & "=" & "'" & txtCedulaX & "'"
    DDocentes.Recordset.FindFirst Buscar
    If DDocentes.Recordset.NoMatch Then
        MsgBox "Cedula No Encontrada, Favor Chequear", , "Busqueda de Registro"
        txtNombreX = ""
        txtCargoX = ""
    Else
        txtCargoX = DDocentes.Recordset.Fields("Cargo")
        txtNombreX = Trim(DDocentes.Recordset.Fields("Apellidos")) & ", " & Trim(DDocentes.Recordset.Fields("Nombres"))
    End If
End Sub

Private Sub txtFecha_Aceptacion_Informe_Click()
    If txtFecha_Aceptacion_Informe = "" Or txtFecha_Aceptacion_Informe = "-" Then
        Fecha_Act.XFecha_Actual = Format(Date, Tipo_Fecha)
    Else
        Fecha_Act.XFecha_Actual = txtFecha_Aceptacion_Informe
    End If
    Fecha_Act.XDia_Min = -2000
    Fecha_Act.XDia_Max = 120
    Fecha_Act.Show vbModal
    txtFecha_Aceptacion_Informe = Fecha_Act.txtFecha
    Unload Fecha_Act
End Sub

Private Sub txtFecha_Aceptacion_Informe_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    Call txtFecha_Aceptacion_Informe_Click
End Sub

Private Sub txtFecha_Aceptacion_Plan_Click()
    If txtFecha_Aceptacion_Plan = "" Or txtFecha_Aceptacion_Plan = "-" Then
        Fecha_Act.XFecha_Actual = Format(Date, Tipo_Fecha)
    Else
        Fecha_Act.XFecha_Actual = txtFecha_Aceptacion_Plan
    End If
    Fecha_Act.XDia_Min = -2000
    Fecha_Act.XDia_Max = 120
    Fecha_Act.Show vbModal
    txtFecha_Aceptacion_Plan = Fecha_Act.txtFecha
    Unload Fecha_Act
End Sub

Private Sub txtFecha_Aceptacion_Plan_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    Call txtFecha_Aceptacion_Plan_Click
End Sub

Private Sub txtFecha_Entrega_Informe_Click()
    If txtFecha_Entrega_Informe = "" Or txtFecha_Entrega_Informe = "-" Then
        Fecha_Act.XFecha_Actual = Format(Date, Tipo_Fecha)
    Else
        Fecha_Act.XFecha_Actual = txtFecha_Entrega_Informe
    End If
    Fecha_Act.XDia_Min = -2000
    Fecha_Act.XDia_Max = 120
    Fecha_Act.Show vbModal
    txtFecha_Entrega_Informe = Fecha_Act.txtFecha
    Unload Fecha_Act
    txtFecha_Aceptacion_Informe = txtFecha_Entrega_Informe
End Sub

Private Sub txtFecha_Entrega_Informe_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    Call txtFecha_Entrega_Informe_Click
End Sub

Private Sub txtFecha_Entrega_Plan_Click()
    If txtFecha_Entrega_Plan = "" Or txtFecha_Entrega_Plan = "-" Then
        Fecha_Act.XFecha_Actual = Format(Date, Tipo_Fecha)
    Else
        Fecha_Act.XFecha_Actual = txtFecha_Entrega_Plan
    End If
    Fecha_Act.XDia_Min = -2000
    Fecha_Act.XDia_Max = 120
    Fecha_Act.Show vbModal
    txtFecha_Entrega_Plan = Fecha_Act.txtFecha
    Unload Fecha_Act
    txtFecha_Aceptacion_Plan = txtFecha_Entrega_Plan
End Sub

Private Sub txtFecha_Entrega_Plan_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    Call txtFecha_Entrega_Plan_Click
End Sub

Private Sub txtFecha_Final_Proceso_Click()
    If txtFecha_Final_Proceso = "" Or txtFecha_Final_Proceso = "-" Then
        Fecha_Act.XFecha_Actual = Format(Date, Tipo_Fecha)
    Else
        Fecha_Act.XFecha_Actual = txtFecha_Final_Proceso
    End If
    Fecha_Act.XDia_Min = -2000
    Fecha_Act.XDia_Max = 120
    Fecha_Act.Show vbModal
    txtFecha_Final_Proceso = Fecha_Act.txtFecha
    Unload Fecha_Act
End Sub

Private Sub txtFecha_Final_Proceso_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    Call txtFecha_Final_Proceso_Click
End Sub

Private Sub txtObservacionX_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Sub Buscar_Casos()
    txtCasos = ""
    Contador = 0
    cmdBoton01.Picture = LoadPicture("")
    cmdBoton02.Picture = LoadPicture("")
    cmdBoton03.Picture = LoadPicture("")
    cmdBoton04.Picture = LoadPicture("")
    cmdBoton05.Picture = LoadPicture("")
    cmdBoton06.Picture = LoadPicture("")
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
                    If txtCasos <> "" Then txtCasos = txtCasos & ", "
                    txtCasos = txtCasos & XObservacion
                    Select Case Contador
                    Case 0:
                        cmdBoton01.Picture = LoadPicture(App.Path & "\Gif\PinOut.bmp")
                    Case 1:
                        cmdBoton02.Picture = LoadPicture(App.Path & "\Gif\PinOut.bmp")
                    Case 2:
                        cmdBoton03.Picture = LoadPicture(App.Path & "\Gif\PinOut.bmp")
                    Case 3:
                        cmdBoton05.Picture = LoadPicture(App.Path & "\Gif\PinOut.bmp")
                    Case 4:
                        cmdBoton04.Picture = LoadPicture(App.Path & "\Gif\PinOut.bmp")
                        cmdBoton06.Picture = LoadPicture(App.Path & "\Gif\PinOut.bmp")
                    End Select
                End If
            End If
        End If
        Contador = Contador + 1
    Loop While Contador <= 3
    If txtCasos <> "" Then
        txtCasos.Visible = True
    Else
        txtCasos.Visible = False
    End If
'    Buscar = "[Cedula]" & "=" & "'" & txtCedula & "'"
'    DSolvencias.Recordset.FindFirst Buscar
'    If Not DSolvencias.Recordset.NoMatch Then
'       If DSolvencias.Recordset.Fields("Administrativo_Caja") = False Then
'            Titulos = "Caja Solvencia Negada"
'            XObservacion = "Alumno Con Caja Solvencia Negada"
'            MsgBox XObservacion, vbCritical, Titulos
'       End If
'    End If
End Sub
