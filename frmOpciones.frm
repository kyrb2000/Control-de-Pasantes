VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpciones 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Opciones de Mantenimiento"
   ClientHeight    =   3795
   ClientLeft      =   2370
   ClientTop       =   2115
   ClientWidth     =   4665
   Icon            =   "frmOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   3960
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picButtons 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   0
      ScaleHeight     =   540
      ScaleWidth      =   4665
      TabIndex        =   1
      Top             =   3255
      Width           =   4665
      Begin VB.CommandButton cmdCancel 
         Caption         =   "&Salir"
         Default         =   -1  'True
         Height          =   540
         Left            =   3480
         Picture         =   "frmOpciones.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   0
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   240
      TabIndex        =   13
      Top             =   480
      Visible         =   0   'False
      Width           =   4215
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Modificar"
         Height          =   540
         Left            =   3120
         Picture         =   "frmOpciones.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdPeriodo 
         Caption         =   "P&eriodo"
         Height          =   540
         Left            =   2040
         Picture         =   "frmOpciones.frx":098E
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton cmdPassword 
         Caption         =   "&Password"
         Height          =   540
         Left            =   960
         Picture         =   "frmOpciones.frx":0CD0
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   1920
         Width           =   975
      End
      Begin VB.Frame Frame7 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   0
         TabIndex        =   24
         Top             =   0
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox txtContraseña 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   28
            Top             =   345
            Width           =   3975
         End
         Begin VB.TextBox txtConfirmar_Contraseña 
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   120
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   25
            Top             =   915
            Width           =   3975
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Contraseña :"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   120
            Width           =   900
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Confirmar Contraseña :"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   675
            Width           =   1605
         End
      End
      Begin VB.Frame Frame8 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Visible         =   0   'False
         Width           =   4215
         Begin VB.TextBox txtNumeroCPres_ID 
            Alignment       =   2  'Center
            DataField       =   "NumeroCPres_ID"
            DataSource      =   "Data2"
            Height          =   285
            Left            =   3360
            Locked          =   -1  'True
            TabIndex        =   43
            Top             =   840
            Width           =   735
         End
         Begin VB.TextBox txtNumeroCPost_ID 
            Alignment       =   2  'Center
            DataField       =   "NumeroCPost_ID"
            DataSource      =   "Data2"
            Height          =   285
            Left            =   2520
            Locked          =   -1  'True
            TabIndex        =   41
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton cmdCerrar_Periodo 
            Caption         =   "&Cerrar Periodo"
            Height          =   540
            Left            =   2760
            Picture         =   "frmOpciones.frx":1012
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   1200
            Width           =   1335
         End
         Begin VB.TextBox txtFecha_Cierre 
            Alignment       =   2  'Center
            DataField       =   "Fecha_Cierre"
            DataSource      =   "Data2"
            Height          =   285
            Left            =   1200
            TabIndex        =   32
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtFecha_Inicio 
            Alignment       =   2  'Center
            DataField       =   "Fecha_Inicio"
            DataSource      =   "Data2"
            Height          =   285
            Left            =   120
            TabIndex        =   31
            Top             =   840
            Width           =   1095
         End
         Begin VB.TextBox txtPeriodo 
            DataField       =   "Periodo"
            DataSource      =   "Data2"
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
            Left            =   1440
            TabIndex        =   30
            Top             =   120
            Width           =   2655
         End
         Begin MSComctlLib.ProgressBar ProgressBar4 
            Height          =   375
            Left            =   120
            TabIndex        =   39
            Top             =   1320
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   1
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Cartas de POS / PRE"
            Height          =   195
            Left            =   2520
            TabIndex        =   42
            Top             =   600
            Width           =   1545
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Cierre"
            Height          =   195
            Left            =   1200
            TabIndex        =   35
            Top             =   600
            Width           =   1125
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Inicio"
            Height          =   195
            Left            =   120
            TabIndex        =   34
            Top             =   600
            Width           =   870
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Periodo :"
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
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   1080
         End
      End
      Begin VB.Frame Frame2 
         Enabled         =   0   'False
         Height          =   1890
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   4215
         Begin VB.TextBox txtNumeroD_ID 
            DataField       =   "NumeroD_ID"
            DataSource      =   "Data2"
            Height          =   285
            Left            =   2040
            TabIndex        =   18
            Top             =   1560
            Width           =   2055
         End
         Begin VB.TextBox txtNumero_ID 
            DataField       =   "Numero_ID"
            DataSource      =   "Data2"
            Height          =   285
            Left            =   2040
            TabIndex        =   17
            Top             =   1245
            Width           =   2055
         End
         Begin VB.TextBox txtRif_Empresa 
            DataField       =   "Rif_Empresa"
            DataSource      =   "Data2"
            Height          =   285
            Left            =   120
            MaxLength       =   25
            TabIndex        =   16
            Top             =   915
            Width           =   3975
         End
         Begin VB.TextBox txtNombre_Empresa 
            DataField       =   "Nombre_Empresa"
            DataSource      =   "Data2"
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Top             =   345
            Width           =   3975
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Numero ID Docente :"
            Height          =   195
            Left            =   120
            TabIndex        =   22
            Top             =   1560
            Width           =   1515
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Numero ID Alumnos :"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   1245
            Width           =   1500
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "R.I.F. de la Empresa"
            Height          =   195
            Left            =   120
            TabIndex        =   20
            Top             =   675
            Width           =   1440
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Nombre de la Empresa"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   120
            Width           =   1605
         End
      End
   End
   Begin VB.Frame Frame4 
      Height          =   2535
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   4215
      Begin VB.Frame Frame5 
         Caption         =   "[Respaldar]"
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   3975
         Begin VB.CommandButton cmdRespaldar 
            Caption         =   "&Respaldar"
            Height          =   375
            Left            =   120
            TabIndex        =   9
            Top             =   240
            Width           =   1095
         End
         Begin MSComctlLib.ProgressBar ProgressBar2 
            Height          =   375
            Left            =   1320
            TabIndex        =   5
            Top             =   240
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   1
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "[Reparar]"
         Height          =   735
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   3975
         Begin VB.CommandButton cmdReparar 
            Caption         =   "&Reparar"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
         Begin MSComctlLib.ProgressBar ProgressBar3 
            Height          =   375
            Left            =   1320
            TabIndex        =   12
            Top             =   240
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   1
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "[Compactar]"
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   3975
         Begin VB.CommandButton cmdCompactar 
            Caption         =   "&Compactar"
            Height          =   375
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   375
            Left            =   1320
            TabIndex        =   8
            Top             =   240
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   661
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   1
         End
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   5318
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Respaldo de Datos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Mantenimiento"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame_Base_Datos 
      Caption         =   "Frame_Base_Datos"
      Height          =   1815
      Left            =   120
      TabIndex        =   40
      Top             =   0
      Visible         =   0   'False
      Width           =   4455
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
         Top             =   1320
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Data Data2 
         Caption         =   "Tabla_Gerenar"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Tabla_Gerenar"
         Top             =   960
         Width           =   2100
      End
      Begin VB.Data DCartas 
         Caption         =   "Cartas"
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
         RecordSource    =   "Cartas"
         Top             =   240
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Data DAlumnos 
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
         Top             =   600
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Data DCentros_Pasantias 
         Caption         =   "Centros_Pasantias"
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
         RecordSource    =   "Centros_Pasantias"
         Top             =   960
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Data DNotas 
         Caption         =   "Notas"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Notas"
         Top             =   240
         Visible         =   0   'False
         Width           =   2100
      End
      Begin VB.Data DSolvencias 
         Caption         =   "Solvencias"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   2280
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Solvencias"
         Top             =   600
         Visible         =   0   'False
         Width           =   2100
      End
   End
End
Attribute VB_Name = "frmOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
'-- Autor : Kenner Y. Roa B. Telefono (0212) 433.55.14
'-- E-mail : kyrb2000@cantv.net / kyrb2000@hotmail.com
'-- Fecha : 22/05/2002 (Caracas / Venezuela)
'-----------------------------------------------------------
Dim WDirectorio As String
Dim WDirecPrint As String

Sub CompactDatabaseX2()
   Dim dbsNeptuno As Database
   Dim prpBucle As Property
   Dim NombreViejo, NombreNuevo
    On Error GoTo Err_CompactDatabaseX2
   WBaseD = WDirectorio + "Datos_Alumnos.mdb"
   WBaseDC = WDirectorio + "Datos_AlumnosC.mdb"
   WBaseDT = WDirectorio + "Datos_AlumnosT.mdb"
   ProgressBar1.Min = 0
   ProgressBar1.Max = 10
   ProgressBar4.Min = 0
   ProgressBar4.Max = 10
   Set dbsNeptuno = OpenDatabase(WBaseD)
   ProgressBar1.Value = 1
   ProgressBar4.Value = 1
   ' Muestra las propiedades de la base de datos original.
   With dbsNeptuno
      ProgressBar1.Value = 2
      ProgressBar4.Value = 2
      Debug.Print .Name & ", versión " & .Version
      Debug.Print "  Secuencia de ordenación = " & .CollatingOrder
      .Close
   End With
   ProgressBar1.Value = 3
   ProgressBar4.Value = 3
   ' Asegúrese de que no existe un archivo con el
   ' nombre de la base de datos compactada.
   If Dir(WBaseDC) <> "" Then _
      Kill WBaseDC
   ProgressBar1.Value = 4
   ProgressBar4.Value = 4
   ' Este instrucción crea una base de datos
   ' Microsoft Jet versión 2.0 compactada y encriptada
   ' de la base de datos Microsoft Jet versión 1.1.
'------------------
    DBEngine.CompactDatabase WBaseD, _
        WBaseDC, dbLangKorean
'------------------
'   DBEngine.CompactDatabase WBaseD, _
'      WBaseDC, , dbEncrypt + dbVersion20
   ProgressBar1.Value = 5
   ProgressBar4.Value = 5
   Set dbsNeptuno = OpenDatabase(WBaseDC)
   ProgressBar1.Value = 6
   ProgressBar4.Value = 6
   ' Muestra las propiedades de la base de datos compactada.
      ProgressBar1.Value = 7
      ProgressBar4.Value = 7
         ProgressBar1.Value = 8
         ProgressBar4.Value = 8
   dbsNeptuno.Close
   ProgressBar1.Value = 9
   ProgressBar4.Value = 9
   If Dir(WBaseDT) <> "" Then _
        Kill WBaseDT
   ProgressBar1.Value = 10
   ProgressBar4.Value = 10
   NombreViejo = WBaseD: NombreNuevo = WBaseDT   ' Se definen nombres de archivo.
   Name NombreViejo As NombreNuevo   ' Se cambia el nombre del archivo.
   NombreViejo = WBaseDC: NombreNuevo = WBaseD   ' Se definen nombres de archivo.
   Name NombreViejo As NombreNuevo   ' Se cambia el nombre del archivo.
   ProgressBar1.Value = 0
   ProgressBar4.Value = 0
   BaseD = WDirectorio + "Datos_Alumnos.mdb"
Exit Sub

Err_CompactDatabaseX2:
   For Each errBucle In DBEngine.Errors
      MsgBox "¡Falló Repair!" & vbCr & _
         "Número de error: " & errBucle.Number & _
         vbCr & errBucle.Description
   Next errBucle
   ProgressBar1.Value = 0
End Sub

Sub CompactDatabaseX3()
   Dim dbsNeptuno As Database
   Dim prpBucle As Property
   Dim NombreViejo, NombreNuevo
    On Error GoTo Err_CompactDatabaseX3
   WBaseD = WDirecPrint + "Movimiento.Mdb"
   WBaseDC = WDirecPrint + "MovimientoC.Mdb"
   WBaseDT = WDirecPrint + "MovimientoT.Mdb"
   ProgressBar1.Min = 0
   ProgressBar1.Max = 10
   'Set dbsNeptuno = OpenDatabase(WBaseD)
   ProgressBar1.Value = 1
   ' Muestra las propiedades de la base de datos original.
   'With dbsNeptuno
   '   ProgressBar1.Value = 2
   '   Debug.Print .Name & ", versión " & .Version
   '   Debug.Print "  Secuencia de ordenación = " & .CollatingOrder
   '   .Close
   'End With
   ProgressBar1.Value = 3
   ' Asegúrese de que no existe un archivo con el
   ' nombre de la base de datos compactada.
   If Dir(WBaseDC) <> "" Then _
      Kill WBaseDC
   ProgressBar1.Value = 4
   ' Este instrucción crea una base de datos
   ' Microsoft Jet versión 2.0 compactada y encriptada
   ' de la base de datos Microsoft Jet versión 1.1.
'------------------
    DBEngine.CompactDatabase WBaseD, _
        WBaseDC, dbLangKorean
'------------------
'   DBEngine.CompactDatabase WBaseD, _
'      WBaseDC, , dbEncrypt + dbVersion30
   ProgressBar1.Value = 5
   Set dbsNeptuno = OpenDatabase(WBaseDC)
   ProgressBar1.Value = 6
   ' Muestra las propiedades de la base de datos compactada.
      ProgressBar1.Value = 7
         ProgressBar1.Value = 8
   dbsNeptuno.Close
   ProgressBar1.Value = 9
   If Dir(WBaseDT) <> "" Then _
        Kill WBaseDT
   ProgressBar1.Value = 10
   NombreViejo = WBaseD: NombreNuevo = WBaseDT   ' Se definen nombres de archivo.
   Name NombreViejo As NombreNuevo   ' Se cambia el nombre del archivo.
   NombreViejo = WBaseDC: NombreNuevo = WBaseD   ' Se definen nombres de archivo.
   Name NombreViejo As NombreNuevo   ' Se cambia el nombre del archivo.
   ProgressBar1.Value = 0
   BaseD = WDirectorio + "Datos_Alumnos.mdb"
Exit Sub

Err_CompactDatabaseX3:
   For Each errBucle In DBEngine.Errors
      MsgBox "¡Falló Repair!" & vbCr & _
         "Número de error: " & errBucle.Number & _
         vbCr & errBucle.Description
   Next errBucle
   ProgressBar1.Value = 0
End Sub

Sub RespaldandoX5()
    Dim dbsNeptuno As Database
    Dim prpBucle As Property
    Dim NombreViejo, NombreNuevo
    WDirectpath = WDirectorio + "Datos_Alumnos.mdb"
    Dim sFile As String
    With dlgCommonDialog
        dlgCommonDialog.FileName = WDirectpath
        dlgCommonDialog.DialogTitle = "Respaldar"
        dlgCommonDialog.CancelError = False
        'Pendiente: establecer los indicadores y atributos del control common dialog
        dlgCommonDialog.Filter = "Bases de Datos (*.mdb)|*.mdb"
        dlgCommonDialog.ShowSave
        If Len(dlgCommonDialog.FileName) = 0 Or dlgCommonDialog.CancelError Then
            Exit Sub
        End If
        sFile = dlgCommonDialog.FileName
    End With
    WBaseD = WDirectorio + "Datos_Alumnos.mdb"
    WBaseDC = WDirectorio + "Datos_AlumnosC.mdb"
    WBaseDT = sFile
    If UCase(WBaseDT) = UCase(WBaseD) Then
        MsgBox "La Ruta del respaldo No es Validad", , "Respaldar"
        Exit Sub
    End If
    ProgressBar2.Min = 0
    ProgressBar2.Max = 10
'    Set dbsNeptuno = OpenDatabase(WBaseD)
    ProgressBar2.Value = 1
    ' Muestra las propiedades de la base de datos original.
'    With dbsNeptuno
'        ProgressBar2.Value = 2
'        Debug.Print .Name & ", versión " & .Version
'        Debug.Print "  Secuencia de ordenación = " & .CollatingOrder
'        .Close
'    End With
    ProgressBar2.Value = 3
    ' Asegúrese de que no existe un archivo con el
    ' nombre de la base de datos compactada.
    If Dir(WBaseDC) <> "" Then _
       Kill WBaseDC
    ProgressBar2.Value = 4
    ' Este instrucción crea una base de datos
    ' Microsoft Jet versión 2.0 compactada y encriptada
    ' de la base de datos Microsoft Jet versión 1.1.
'------------------
    DBEngine.CompactDatabase WBaseD, _
        WBaseDC, dbLangKorean
'------------------
'    DBEngine.CompactDatabase WBaseD, _
'      WBaseDC, , dbEncrypt + dbVersion20
    ProgressBar2.Value = 5
'    Set dbsNeptuno = OpenDatabase(WBaseDC)
    ProgressBar2.Value = 6
    ' Muestra las propiedades de la base de datos compactada.
      ProgressBar2.Value = 7
'    dbsNeptuno.Close
    ProgressBar2.Value = 8
    If Dir(WBaseDT) <> "" Then _
        Kill WBaseDT
    ProgressBar2.Value = 9
    NombreViejo = WBaseD: NombreNuevo = WBaseDT   ' Se definen nombres de archivo.
    Name NombreViejo As NombreNuevo   ' Se cambia el nombre del archivo.
    ProgressBar2.Value = 10
    NombreViejo = WBaseDC: NombreNuevo = WBaseD   ' Se definen nombres de archivo.
    Name NombreViejo As NombreNuevo   ' Se cambia el nombre del archivo.
    ProgressBar2.Value = 0
End Sub

Sub RepairDatabaseX1()

    Dim errBucle As Error
    WBaseD = WDirectorio + "Datos_Alumnos.mdb"
    ProgressBar3.Min = 0
    ProgressBar3.Max = 5
    ProgressBar3.Value = 1
    ProgressBar4.Min = 0
    ProgressBar4.Max = 5
    ProgressBar4.Value = 1
    On Error GoTo Err_Reparar
    ProgressBar3.Value = 2
    ProgressBar3.Value = 3
    ProgressBar4.Value = 2
    ProgressBar4.Value = 3
    DBEngine.RepairDatabase WBaseD
    ProgressBar3.Value = 4
    ProgressBar4.Value = 4
    On Error GoTo 0
    ProgressBar3.Value = 5
    ProgressBar3.Value = 0
    ProgressBar4.Value = 5
    ProgressBar4.Value = 0
   Exit Sub

Err_Reparar:

   For Each errBucle In DBEngine.Errors
      MsgBox "¡Falló Repair!" & vbCr & _
         "Número de error: " & errBucle.Number & _
         vbCr & errBucle.Description
   Next errBucle
    ProgressBar3.Value = 0
End Sub

Sub RepairDatabaseX2()

    Dim errBucle As Error
    WBaseD = WDirecPrint + "Movimiento.Mdb"
    ProgressBar3.Min = 0
    ProgressBar3.Max = 5
    ProgressBar3.Value = 1
    On Error GoTo Err_Reparar
    ProgressBar3.Value = 2
    ProgressBar3.Value = 3
    DBEngine.RepairDatabase WBaseD
    ProgressBar3.Value = 4
    On Error GoTo 0
    ProgressBar3.Value = 5
    ProgressBar3.Value = 0
   Exit Sub

Err_Reparar:

   For Each errBucle In DBEngine.Errors
      MsgBox "¡Falló Repair!" & vbCr & _
         "Número de error: " & errBucle.Number & _
         vbCr & errBucle.Description
   Next errBucle
    ProgressBar3.Value = 0
End Sub

Private Sub cmdCancel_Click()
    If cmdCancel.Caption = "&Salir" Then
        Unload Me
    Else
        txtContraseña = ""
        txtConfirmar_Contraseña = ""
        Data2.Recordset.CancelUpdate
        Frame2.Enabled = False
        Frame7.Enabled = False
        Frame7.Visible = False
        Frame8.Enabled = False
        Frame8.Visible = False
        '--------- Botones ------------
        cmdUpdate.Caption = "&Modificar"
        cmdCancel.Caption = "&Salir"
        cmdPassword.Enabled = True
        cmdPeriodo.Enabled = True
        '------------------------------'
    End If
End Sub

Sub Desactivar_Bases_de_Datos()
    Data2.DatabaseName = ""
    Data2.Refresh
    DCartas.DatabaseName = ""
    DCartas.Refresh
    DAlumnos.DatabaseName = ""
    DAlumnos.Refresh
    DCentros_Pasantias.DatabaseName = ""
    DCentros_Pasantias.Refresh
    DNotas.DatabaseName = ""
    DNotas.Refresh
    DSolvencias.DatabaseName = ""
    DSolvencias.Refresh
    DControl.DatabaseName = ""
    DControl.Refresh
    DataEnvironment1.Connection1.Close
End Sub

Sub Activar_Bases_de_Datos()
    WBaseD = WDirectorio + "Datos_Alumnos.mdb"
    Data2.DatabaseName = WBaseD
    Data2.Refresh
    Data2.Recordset.MoveFirst
    DCartas.DatabaseName = WBaseD
    DCartas.Refresh
    DAlumnos.DatabaseName = WBaseD
    DAlumnos.Refresh
    DCentros_Pasantias.DatabaseName = WBaseD
    DCentros_Pasantias.Refresh
    DNotas.DatabaseName = WBaseD
    DNotas.Refresh
    DSolvencias.DatabaseName = WBaseD
    DSolvencias.Refresh
    DControl.DatabaseName = WBaseD
    DControl.Refresh
    DataEnvironment1.Connection1.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\Datos_Alumnos.mdb"
End Sub
Private Sub cmdCerrar_Periodo_Click()
    If MsgBox("Es Recomendable Respaldar Antes de Realizar el Proceso" & Chr(10) & Chr(13) & Chr(10) & Chr(13) & "¿Desea Cerrar el Periodo?", _
         vbYesNo) = vbYes Then
        If MsgBox("Este Proceso Eliminara" & Chr(10) & Chr(13) & "Los Datos Existentes", _
            vbYesNo + vbCritical) = vbYes Then
            Desactivar_Bases_de_Datos
            CompactDatabaseX2
            RepairDatabaseX1
            Activar_Bases_de_Datos
            '------------------------------------------------
            If Not DCartas.Recordset.EOF Then
                DCartas.Recordset.MoveFirst
                Do Until DCartas.Recordset.EOF
                    If Not DCartas.Recordset.EOF Then
                        DCartas.Recordset.Delete
                    End If
                    DCartas.Recordset.MoveNext
                Loop
            End If
            '------------------------------------------------
            If Not DAlumnos.Recordset.EOF Then
                DAlumnos.Recordset.MoveFirst
                Do Until DAlumnos.Recordset.EOF
                    If Not DAlumnos.Recordset.EOF Then
                        DAlumnos.Recordset.Delete
                    End If
                    DAlumnos.Recordset.MoveNext
                Loop
            End If
            '------------------------------------------------
            If Not DCentros_Pasantias.Recordset.EOF Then
                DCentros_Pasantias.Recordset.MoveFirst
                Do Until DCentros_Pasantias.Recordset.EOF
                    If Not DCentros_Pasantias.Recordset.EOF Then
                        DCentros_Pasantias.Recordset.Delete
                    End If
                    DCentros_Pasantias.Recordset.MoveNext
                Loop
            End If
            '------------------------------------------------
            If Not DNotas.Recordset.EOF Then
                DNotas.Recordset.MoveFirst
                Do Until DNotas.Recordset.EOF
                    If Not DNotas.Recordset.EOF Then
                        DNotas.Recordset.Delete
                    End If
                    DNotas.Recordset.MoveNext
                Loop
            End If
            '------------------------------------------------
            If Not DSolvencias.Recordset.EOF Then
                DSolvencias.Recordset.MoveFirst
                Do Until DSolvencias.Recordset.EOF
                    If Not DSolvencias.Recordset.EOF Then
                        DSolvencias.Recordset.Delete
                    End If
                    DSolvencias.Recordset.MoveNext
                Loop
            End If
            '------------------------------------------------
            If Not DControl.Recordset.EOF Then
                DControl.Recordset.MoveFirst
                Do Until DControl.Recordset.EOF
                    If Not DControl.Recordset.EOF Then
                        DControl.Recordset.Delete
                    End If
                    DControl.Recordset.MoveNext
                Loop
            End If
            '------------------------------------------------
            Data2.Recordset.MoveFirst
            Data2.Recordset.Edit
            Data2.Recordset.Fields("Numero_ID") = 1
            'Data2.Recordset.Fields("NumeroD_ID") = 1
            Data2.Recordset.Fields("NumeroCPost_ID") = 1
            Data2.Recordset.Fields("NumeroCPres_ID") = 1
        End If
    End If
    cmdUpdate_Click
End Sub

Private Sub cmdCompactar_Click()
    Frame3.Caption = "[Compactando]"
    If MsgBox("¿Desea Compactar la Base de Datos?", _
         vbYesNo) = vbYes Then
        CompactDatabaseX2
        'CompactDatabaseX3
    End If
    Frame3.Caption = "[Compactar]"
End Sub

Private Sub cmdPassword_Click()
    Frame7.Visible = True
    '--------- Botones ------------
    cmdPassword.Enabled = False
    cmdPeriodo.Enabled = False
    '------------------------------'
End Sub

Private Sub cmdPeriodo_Click()
    Frame8.Visible = True
    '--------- Botones ------------
    cmdPassword.Enabled = False
    cmdPeriodo.Enabled = False
    '------------------------------'
End Sub

Private Sub cmdRespaldar_Click()
    Frame5.Caption = "[Respaldando]"
    RespaldandoX5
    Frame5.Caption = "[Respaldar]"
End Sub

Private Sub cmdReparar_Click()
    Frame6.Caption = "[Reparando]"
    If MsgBox("¿Desea Reparar la Base de Datos?", _
         vbYesNo) = vbYes Then
        RepairDatabaseX1
        'RepairDatabaseX2
    End If
    Frame6.Caption = "[Reparar]"
End Sub

Private Sub cmdUpdate_Click()
    If cmdUpdate.Caption = "&Actualizar" Then
        If Frame7.Visible Then
            If txtContraseña <> txtConfirmar_Contraseña Or txtConfirmar_Contraseña = "" Then
                MsgBox "La Contraseña es Errada"
                Call cmdCancel_Click
                Exit Sub
            End If
            Data2.Recordset.Fields("Contraseña") = txtContraseña
        End If
        If txtFecha_Inicio = "" Then
            txtFecha_Inicio = "-"
        End If
        If txtFecha_Cierre = "" Then
            txtFecha_Cierre = "-"
        End If
        Frame2.Enabled = False
        Frame7.Enabled = False
        Frame7.Visible = False
        Frame8.Enabled = False
        Frame8.Visible = False
        Data2.UpdateRecord
        Data2.Recordset.Bookmark = Data2.Recordset.LastModified
        '--------- Botones ------------
        cmdUpdate.Caption = "&Modificar"
        cmdCancel.Caption = "&Salir"
        cmdPassword.Enabled = True
        cmdPeriodo.Enabled = True
        '------------------------------'
    Else
        Frame2.Enabled = True
        Frame7.Enabled = True
        Frame8.Enabled = True
        Data2.Recordset.Edit
        '--------- Botones ------------
        cmdUpdate.Caption = "&Actualizar"
        cmdCancel.Caption = "&Cancelar"
        cmdPassword.Enabled = False
        cmdPeriodo.Enabled = False
        '------------------------------'
    End If
End Sub

Private Sub Form_Activate()
    Desactivar_Bases_de_Datos
End Sub

Private Sub Form_Load()
    WDirectorio = App.Path & "\"
    WDirecPrint = App.Path & "\Reportes\"
End Sub

Private Sub TabStrip1_Click()
    Select Case TabStrip1.SelectedItem
    Case "Respaldo de Datos":
        Desactivar_Bases_de_Datos
        Frame1.Visible = False
        Frame4.Visible = True
    Case "Mantenimiento":
        Activar_Bases_de_Datos
        Frame1.Visible = True
        Frame4.Visible = False
    End Select
End Sub

Private Sub txtConfirmar_Contraseña_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdUpdate.SetFocus
End Sub

Private Sub txtContraseña_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtConfirmar_Contraseña.SetFocus
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

Private Sub txtFecha_Cierre_Click()
    If txtFecha_Cierre = "" Or txtFecha_Cierre = "-" Then
        Fecha_Act.XFecha_Actual = Format(Date, Tipo_Fecha)
    Else
        Fecha_Act.XFecha_Actual = txtFecha_Cierre
    End If
    Fecha_Act.XDia_Min = -2000
    Fecha_Act.XDia_Max = 120
    Fecha_Act.Show vbModal
    txtFecha_Cierre = Fecha_Act.txtFecha
    Unload Fecha_Act
End Sub

Private Sub txtFecha_Cierre_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    Call txtFecha_Cierre_Click
End Sub

Private Sub txtNombre_Empresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtRif_Empresa.SetFocus
End Sub

Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
    If Len(txtPeriodo) <> 0 Then
        cmdUpdate.Enabled = True
    Else
        cmdUpdate.Enabled = False
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        cmdUpdate.SetFocus
    End If
End Sub

Private Sub txtRif_Empresa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNumero_ID.SetFocus
End Sub

Private Sub txtNumero_ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNumeroD_ID.SetFocus
End Sub

Private Sub txtNumeroD_ID_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdUpdate.SetFocus
End Sub

