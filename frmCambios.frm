VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmCambios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cambios Generales"
   ClientHeight    =   5625
   ClientLeft      =   1770
   ClientTop       =   1860
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   6750
   Begin VB.Frame Frame_Base_Datos 
      Caption         =   "Frame_Base_Datos"
      Height          =   2775
      Left            =   4320
      TabIndex        =   18
      Top             =   1680
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Data DNotas 
         Caption         =   "Notas"
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
         RecordSource    =   "Notas"
         Top             =   2400
         Visible         =   0   'False
         Width           =   2100
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
         Width           =   2100
      End
      Begin VB.Data DDocentes 
         Caption         =   "Docentes"
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
         RecordSource    =   "Docentes"
         Top             =   960
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
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Solvencias"
         Top             =   1320
         Visible         =   0   'False
         Width           =   2100
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
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Data Data1 
         Caption         =   "Alumnos"
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
         RecordSource    =   "Alumnos"
         Top             =   600
         Visible         =   0   'False
         Width           =   2055
      End
      Begin MSAdodcLib.Adodc mdbAlumnos 
         Height          =   330
         Left            =   120
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
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
         Connect         =   $"frmCambios.frx":0000
         OLEDBString     =   $"frmCambios.frx":008C
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Select * from Alumnos order by Seccion"
         Caption         =   "Alumnos"
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
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   6735
      Begin VB.Frame Frame9 
         Height          =   2775
         Left            =   120
         TabIndex        =   20
         Top             =   1800
         Visible         =   0   'False
         Width           =   3375
         Begin VB.ListBox List_Buscar 
            Height          =   1815
            Left            =   120
            TabIndex        =   22
            Top             =   240
            Width           =   3135
         End
         Begin VB.CommandButton cmdCerrar4 
            Caption         =   "&Cerrar"
            Height          =   555
            Left            =   1200
            Picture         =   "frmCambios.frx":0118
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   2190
            Width           =   975
         End
      End
      Begin VB.CommandButton cmdAceptarC 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   5520
         TabIndex        =   4
         Top             =   1200
         Width           =   1095
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "frmCambios.frx":045A
         Height          =   1335
         Left            =   120
         TabIndex        =   19
         Top             =   1680
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   2355
         _Version        =   393216
         AllowUpdate     =   0   'False
         AllowArrows     =   -1  'True
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   8202
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtPeriodo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1320
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin VB.Frame Frame3 
         Caption         =   "Campos a Modificar"
         Enabled         =   0   'False
         Height          =   1575
         Left            =   120
         TabIndex        =   15
         Top             =   3000
         Width           =   6495
         Begin VB.CommandButton cmdBoton10 
            Height          =   345
            Index           =   0
            Left            =   3840
            MaskColor       =   &H8000000F&
            Picture         =   "frmCambios.frx":0473
            TabIndex        =   23
            Top             =   1080
            Width           =   255
         End
         Begin VB.TextBox txtDocente 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   7
            Top             =   1080
            Width           =   1935
         End
         Begin VB.TextBox txtSección 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   1800
            MaxLength       =   10
            TabIndex        =   6
            Top             =   600
            Width           =   1935
         End
         Begin VB.CheckBox chkSolvente 
            Alignment       =   1  'Right Justify
            Caption         =   "Solvente"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   3615
         End
         Begin VB.Label lblDocente 
            AutoSize        =   -1  'True
            Caption         =   "Docente"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   1080
            Width           =   840
         End
         Begin VB.Label lblSección 
            AutoSize        =   -1  'True
            Caption         =   "Sección"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   765
         End
      End
      Begin VB.ComboBox cobSolvencia 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frmCambios.frx":07B5
         Left            =   1320
         List            =   "frmCambios.frx":07D4
         TabIndex        =   3
         Text            =   "-------------------------------------------"
         Top             =   1200
         Width           =   4095
      End
      Begin VB.TextBox txtCampo 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3360
         MaxLength       =   10
         TabIndex        =   2
         Top             =   720
         Width           =   1935
      End
      Begin VB.ComboBox cobCampo 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         ItemData        =   "frmCambios.frx":0901
         Left            =   1320
         List            =   "frmCambios.frx":0911
         TabIndex        =   1
         Text            =   "-------------------"
         Top             =   720
         Width           =   1935
      End
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   1200
         TabIndex        =   11
         Top             =   4680
         Width           =   4215
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cerrar"
            Height          =   615
            Left            =   2160
            Picture         =   "frmCambios.frx":094E
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   170
            Width           =   1695
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            Enabled         =   0   'False
            Height          =   615
            Left            =   360
            Picture         =   "frmCambios.frx":0C90
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   170
            Width           =   1695
         End
      End
      Begin VB.Label lblSolvencia 
         AutoSize        =   -1  'True
         Caption         =   "Solvencia :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label lblPeriodo 
         AutoSize        =   -1  'True
         Caption         =   "Periodo :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   900
      End
      Begin VB.Label lblCampo 
         AutoSize        =   -1  'True
         Caption         =   "Campo :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   840
      End
   End
End
Attribute VB_Name = "frmCambios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Campo0, Campo1, Campo2 As String

Private Sub cmdAceptar_Click()
    If mdbAlumnos.Recordset.EOF Then
        Call cmdCancelar_Click
        Exit Sub
    End If
    mdbAlumnos.Recordset.MoveFirst
    Do
        If txtSección <> "" Then
            mdbAlumnos.Recordset.Fields("Seccion") = txtSección
        End If
        txtCedula = mdbAlumnos.Recordset.Fields("Cedula")
        If cobSolvencia.List(cobSolvencia.ListIndex) <> "-------------------------------------------" Then
            Buscar = "[Cedula]" & "=" & "'" & txtCedula & "'"
            DSolvencias.Recordset.FindFirst Buscar
            If DSolvencias.Recordset.NoMatch Then
                'MsgBox "No encontrado y es Añadido", , "Busqueda de Registro"
                txtAcum60 = 0
                txtAcum40 = 0
                txtTotal = 0
                DSolvencias.Recordset.AddNew
                DSolvencias.Recordset.Fields("Cedula") = txtCedula
                DSolvencias.Recordset.Fields("Taller") = False
                DSolvencias.Recordset.Fields("Administrativo_Caja") = False
                DSolvencias.Recordset.Fields("Academica_Control_de_Estudio") = False
                DSolvencias.Recordset.Fields("Academica_Departamento_de_Grado") = False
                DSolvencias.Recordset.Fields("Notas_Entregadas") = False
                DSolvencias.Recordset.Fields("Casos_Especiales") = False
                DSolvencias.Recordset.Fields("Entrego_Carta_Aceptacion") = False
                DSolvencias.Recordset.Fields("Realizando_Pasantia") = False
                Select Case cobSolvencia.List(cobSolvencia.ListIndex)
                Case "Solvencia de Realizar el Taller":
                    DSolvencias.Recordset.Fields("Taller") = chkSolvente.Value
                Case "Solvencia Administrativo en Caja":
                    DSolvencias.Recordset.Fields("Administrativo_Caja") = chkSolvente.Value
                Case "Solvencia Academica en Control de Estudio":
                    DSolvencias.Recordset.Fields("Academica_Control_de_Estudio") = chkSolvente.Value
                Case "Solvencia Academica en Departamento de Grado":
                    DSolvencias.Recordset.Fields("Academica_Departamento_de_Grado") = chkSolvente.Value
                Case "Solvencia Notas Entregadas":
                    DSolvencias.Recordset.Fields("Notas_Entregadas") = chkSolvente.Value
                Case "Casos Especiales":
                    DSolvencias.Recordset.Fields("Casos_Especiales") = chkSolvente.Value
                Case "Entrego Carta de Aceptación":
                    DSolvencias.Recordset.Fields("Entrego_Carta_Aceptacion") = chkSolvente.Value
                Case "Realizando Pasantia":
                    DSolvencias.Recordset.Fields("Realizando_Pasantia") = chkSolvente.Value
                End Select
                DSolvencias.UpdateRecord
                DNotas.Recordset.AddNew
                DNotas.Recordset.Fields("Cedula") = txtCedula
                DNotas.Recordset.Fields("Acum_60") = txtAcum60
                DNotas.Recordset.Fields("Acum_40") = txtAcum40
                DNotas.UpdateRecord
            Else
                DSolvencias.Recordset.Edit
                Select Case cobSolvencia.List(cobSolvencia.ListIndex)
                Case "Solvencia de Realizar el Taller":
                    DSolvencias.Recordset.Fields("Taller") = chkSolvente.Value
                Case "Solvencia Administrativo en Caja":
                    DSolvencias.Recordset.Fields("Administrativo_Caja") = chkSolvente.Value
                Case "Solvencia Academica en Control de Estudio":
                    DSolvencias.Recordset.Fields("Academica_Control_de_Estudio") = chkSolvente.Value
                Case "Solvencia Academica en Departamento de Grado":
                    DSolvencias.Recordset.Fields("Academica_Departamento_de_Grado") = chkSolvente.Value
                Case "Solvencia Notas Entregadas":
                    DSolvencias.Recordset.Fields("Notas_Entregadas") = chkSolvente.Value
                Case "Casos Especiales":
                    DSolvencias.Recordset.Fields("Casos_Especiales") = chkSolvente.Value
                Case "Entrego Carta de Aceptación":
                    DSolvencias.Recordset.Fields("Entrego_Carta_Aceptacion") = chkSolvente.Value
                Case "Realizando Pasantia":
                    DSolvencias.Recordset.Fields("Realizando_Pasantia") = chkSolvente.Value
                End Select
                DSolvencias.UpdateRecord
            End If
        Else
            If txtDocente <> "" Then
                Buscar = "[Cedula]" & "=" & "'" & txtCedula & "'"
                DNotas.Recordset.FindFirst Buscar
                If Not DNotas.Recordset.NoMatch Then
                    DNotas.Recordset.Edit
                    DNotas.Recordset.Fields("Cedula_Tutor") = txtDocente.Text
                    DNotas.UpdateRecord
                End If
            End If
        End If
        mdbAlumnos.Recordset.Update
        mdbAlumnos.Recordset.MoveNext
    Loop While Not mdbAlumnos.Recordset.EOF
    
End Sub

Private Sub cmdAceptarC_Click()
    Dim SQL As String
    If txtpediodos = "" Then
        Campo0 = ""
        SQL = ""
    Else
        Campo0 = txtPeriodo
        SQL = "Where Periodo=" & "'" & txtPeriodo & "'"
    End If
    Select Case cobCampo.List(cobCampo.ListIndex)
    Case "-------------------":
        Campo1 = ""
        txtSección.Enabled = False
        If Len(SQL) < 1 Then
            SQL = ""
        End If
    Case "Sección":
        If Len(SQL) > 1 Then
            SQL = SQL & " and "
        Else
            SQL = "Where "
        End If
        Campo1 = txtCampo
        SQL = SQL & "Seccion=" & "'" & txtCampo & "'"
    Case "Cedula":
        txtSección.Enabled = False
        If Len(SQL) > 1 Then
            SQL = SQL & " and "
        Else
            SQL = "Where "
        End If
        Campo1 = txtCampo
        SQL = SQL & "Cedula=" & "'" & txtCampo & "'"
    Case "N° Identificación":
        txtSección.Enabled = False
        If Len(SQL) > 1 Then
            SQL = SQL & " and "
        Else
            SQL = "Where "
        End If
        Campo1 = txtCampo
        SQL = SQL & "Numero_ID=" & txtCampo
    End Select
    Select Case cobSolvencia.List(cobSolvencia.ListIndex)
    Case "-------------------------------------------":
        Campo2 = ""
        chkSolvente.Enabled = False
    Case Else
        Campo2 = cobSolvencia.List(cobSolvencia.ListIndex)
    End Select
    SQL = "Select * from Alumnos " & SQL & " order by Cedula"
    mdbAlumnos.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
    mdbAlumnos.RecordSource = SQL
    mdbAlumnos.Refresh
    DataGrid1.Enabled = True
    DataGrid1.Refresh
    Frame3.Enabled = True
    cmdAceptar.Enabled = True
    cmdCancelar.Caption = "&Cancelar"
End Sub

Private Sub cmdBoton10_Click(Index As Integer)
    Select Case Index
    Case 0:
        Frame9.Visible = True
        Frame3.Enabled = False
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
    End Select
End Sub

Private Sub cmdCancelar_Click()
    If cmdCancelar.Caption = "&Cerrar" Then
        Unload Me
    Else
        cobSolvencia.ListIndex = 0
        cobCampo.ListIndex = 0
        txtCampo = ""
        txtSección = ""
        txtDocente = ""
        chkSolvente.Value = 0
        txtSección.Enabled = True
        chkSolvente.Enabled = True
        Frame3.Enabled = False
        cmdAceptar.Enabled = False
        cmdCancelar.Caption = "&Cerrar"
        txtPeriodo = Data2.Recordset.Fields("Periodo")
        txtPeriodo.SetFocus
    End If
End Sub

Private Sub cmdCerrar4_Click()
    Frame9.Visible = False
    Frame3.Enabled = True
End Sub

Private Sub Form_Activate()
    txtPeriodo = Data2.Recordset.Fields("Periodo")
End Sub

Private Sub List_Buscar_DblClick()
'    Select Case Buscar_Bot
'    Case 0:
        txtDocente.Text = List_Buscar.ItemData(List_Buscar.ListIndex)
'    End Select
    Frame9.Visible = False
    Frame3.Enabled = True
End Sub

Private Sub txtCampo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSección_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
