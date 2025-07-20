VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmReportes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reportes"
   ClientHeight    =   2205
   ClientLeft      =   4125
   ClientTop       =   3615
   ClientWidth     =   3855
   Icon            =   "frmReportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2205
   ScaleWidth      =   3855
   Begin VB.Frame Frame_Base_Datos1 
      Caption         =   "Frame_Base_Datos"
      Height          =   735
      Left            =   1800
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1935
      Begin VB.Data DCentros_Pasantias 
         Caption         =   "Centros_Pasantias"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Centros_Pasantias"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Data DTablas 
         Caption         =   "DTablas"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   ""
         Top             =   2040
         Width           =   3375
      End
      Begin VB.Data DDocentes 
         Caption         =   "Docentes"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   1920
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Docentes"
         Top             =   1680
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Data DSolvencias 
         Caption         =   "Solvencias"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   1920
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Solvencias"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Data DNotas 
         Caption         =   "Notas"
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
         RecordSource    =   "Notas"
         Top             =   1320
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.Data DControl 
         Caption         =   "Control"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   1920
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Control"
         Top             =   960
         Visible         =   0   'False
         Width           =   1620
      End
      Begin VB.Data DAlumnos 
         Caption         =   "Alumnos"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   345
         Left            =   1800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Alumnos"
         Top             =   240
         Width           =   1725
      End
      Begin VB.Data DEspecialidad 
         Caption         =   "Especialidad"
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
         RecordSource    =   "Especialidad"
         Top             =   600
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.Data DTabla_Gerenar 
         Caption         =   "Tabla_Gerenar"
         Connect         =   "Access"
         DatabaseName    =   ""
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   1800
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Tabla_Gerenar"
         Top             =   600
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.TextBox adodc_Cedula 
         DataField       =   "Cedula"
         DataSource      =   "Adodc1_Alumnos"
         Height          =   285
         Left            =   120
         TabIndex        =   2
         Text            =   "Cedula"
         Top             =   960
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSAdodcLib.Adodc Adodc1_Alumnos 
         Height          =   330
         Left            =   120
         Top             =   240
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
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
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   "Reporte_Mov"
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
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
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   2880
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   2280
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame_Alumnos 
      Height          =   2175
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3855
      Begin VB.CommandButton cmdCerrar 
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
         Left            =   2040
         Picture         =   "frmReportes.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
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
         Left            =   720
         Picture         =   "frmReportes.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "General"
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
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Value           =   -1  'True
         Width           =   3015
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Individual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
'-- Autor : Kenner Y. Roa B. Telefono (0212) 433.55.14
'-- E-mail : kyrb2000@cantv.net / kyrb2000@hotmail.com
'-- Fecha : 22/05/2002 (Caracas / Venezuela)
'-----------------------------------------------------------
' Reportes Generados
' Tabla_Alumnos                     =Alumnos_G.rpt
'                                    Alumnos_I.rpt
' Tabla_Ficha_Pasantia              =Fichas_G.rpt
'                                    Fichas_I.rpt
' Tabla_Solvencia_de_los_Pasantes   =Formulario_G.rpt
' Tabla_Docentes                    =Docentes.rpt
' Tabla_Registro_del_Pasante        =Formulario_E.rpt
' Tabla_Registro_del_Tutor_Academico=Formulario_T.rpt
'-----------------------------------------------------------
Public Tipo_Reporte As String
Dim Campo1, Campo2, Campo3 As String
Dim Campo4, Campo5, Campo6, Campo7 As String
Dim CampoB1, CampoB2, CampoB3 As Boolean
Dim CampoC1, CampoC2, CampoC3 As String
Dim WDirecPrint As String
Dim SQL As String

Private Sub cmdImprimir_Click()
    Select Case Tipo_Reporte
    Case "ALU":
        Call Tabla_Alumnos
    Case "FIC":
        Call Tabla_Ficha_Pasantia
    Case "SOL":
        Call Tabla_Solvencia_de_los_Pasantes
    Case "DOC":
        Call Tabla_Docentes
    Case "REG":
        Call Tabla_Registro_del_Pasante
    Case "TUR":
        Call Tabla_Registro_del_Tutor_Academico
    Case "CAS":
        Call Reportes_de_Casos_Especiales
    Case "TGR":
        Call Reportes_Total_por_Seccion
    Case "TGP":
        Call Reportes_Total_por_Seccion_P
    Case "CEN":
        Call Reportes_Centros_de_Pasantias
    Case "RPP":
        Call Reportes_Pendientes_X_Pasantías
    Case "REP":
        Call Reportes_Realizando_Pasantías
    End Select
End Sub

'-----------------------------------------------------------
' Tabla_Alumnos                     =Alumnos_G.rpt
'                                    Alumnos_I.rpt
'-----------------------------------------------------------
Sub Tabla_Alumnos()
    If Option2.Value = True Then
        If FGeneral = True Then
            Set Rep_Alumnos.DataSource = DataEnvironment1.Connection1.Execute(SQL)
            Rep_Alumnos.Caption = Me.Caption
            Rep_Alumnos.Show
        End If
    Else
        GoTo Reporte_Individual
    End If

Exit Sub
Reporte_Individual:
    Dim Cedula, Buscar As String
    Cedula = InputBox("Introduzca La Cedula :", "Busqueda de Datos")
    If Cedula <> "" Then
        Buscar = "[Cedula]" & "=" & "'" & Cedula & "'"
        DAlumnos.Recordset.FindFirst Buscar
        If DAlumnos.Recordset.NoMatch Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    SQL = "SELECT Alumnos.Nacionalidad, Alumnos.Cedula, Alumnos.Nombres, Alumnos.Apellidos, Alumnos.Seccion, Especialidad.Descripcion," & _
            "IIf(Mid([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, Alumnos.Telefono, Alumnos.Direccion, Alumnos.Periodo " & _
            "FROM Alumnos INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2) = Especialidad.Codigo " & _
            " WHERE Alumnos.Cedula ='" & Cedula & "'"
'---------------------------------------------------------------------------------------------
    Set Rep_Alumnos.DataSource = DataEnvironment1.Connection1.Execute(SQL)
    Rep_Alumnos.Caption = Me.Caption
    Rep_Alumnos.Show
End Sub

'-----------------------------------------------------------
' Tabla_Ficha_Pasantia              =Fichas_G.rpt
'                                    Fichas_I.rpt
'-----------------------------------------------------------
Sub Tabla_Ficha_Pasantia()
    If Option2.Value = True Then
        If FGeneral = True Then
            Cargar_Datos (SQL)
            GoTo Reporte_General
        End If
    Else
        GoTo Reporte_Individual
    End If
Exit Sub

Reporte_General:
    CrystalReport2.ReportFileName = WDirecPrint + "Fichas_G.rpt"
    CrystalReport2.WindowTitle = Me.Caption
    CrystalReport2.Action = 1
Exit Sub

Reporte_Individual:
    Dim Buscar As String
    Buscar = InputBox("Introduzca La Cedula :", "Busqueda de Datos")
    If Buscar <> "" Then
        Buscar = "[Cedula]" & "=" & "'" & Buscar & "'"
        DAlumnos.Recordset.FindFirst Buscar
        If DAlumnos.Recordset.NoMatch Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
'---------------------------------------------------------------------------------------------
        Campo1 = DAlumnos.Recordset.Fields("Cedula")
        Campo2 = DAlumnos.Recordset.Fields("Seccion")
        Campo3 = DAlumnos.Recordset.Fields("Apellidos") & ", " & DAlumnos.Recordset.Fields("Nombres")
        Campo4 = DAlumnos.Recordset.Fields("Telefono")
        Campo5 = DAlumnos.Recordset.Fields("Periodo")
        Buscar = "[Codigo]" & "=" & "'" & Mid(Campo2, 1, 2) & "'"
        DEspecialidad.Recordset.FindFirst Buscar
        Campo6 = DEspecialidad.Recordset.Fields("Descripcion")
        Select Case Mid(Campo2, 5, 1)
        Case "1"
            Campo7 = "Mañana"
        Case "3"
            Campo7 = "Noche"
        End Select
        Buscar = "[Cedula]" & "=" & "'" & txtCedula & "'"
        DSolvencias.Recordset.FindFirst Buscar
        CampoB1 = IIf(DSolvencias.Recordset.Fields("Taller") = True, 1, 0)
        CampoB2 = IIf(DSolvencias.Recordset.Fields("Administrativo_Caja") = True, 1, 0)
        CampoB3 = IIf(DSolvencias.Recordset.Fields("Academica_Control_de_Estudio") = True, 1, 0)
'---------------------------------------------------------------------------------------------
    CrystalReport1.ReportFileName = WDirecPrint + "Fichas_I.rpt"
    CrystalReport1.WindowTitle = Me.Caption
    CrystalReport1.Formulas(0) = "Cedula='" & Campo1 & "'"
    CrystalReport1.Formulas(1) = "ApeNom='" & Campo3 & "'"
    CrystalReport1.Formulas(2) = "Telefono='" & Campo4 & "'"
    CrystalReport1.Formulas(3) = "Seccion='" & Campo2 & "'"
    CrystalReport1.Formulas(4) = "Especialidad='" & Campo6 & "'"
    CrystalReport1.Formulas(5) = "Turno='" & Campo7 & "'"
    CrystalReport1.Formulas(6) = "Periodo='" & Campo5 & "'"
    CrystalReport1.Formulas(7) = "SNum1='" & IIf(CampoB1 = 0, " ", "X") & "'"
    CrystalReport1.Formulas(8) = "SNum2='" & IIf(CampoB2 = 0, " ", "X") & "'"
    CrystalReport1.Formulas(9) = "SNum3='" & IIf(CampoB3 = 0, " ", "X") & "'"
    CrystalReport1.Action = 1
End Sub

'-----------------------------------------------------------
' Tabla_Solvencia_de_los_Pasantes   =Formulario_G.rpt
'-----------------------------------------------------------
Sub Tabla_Solvencia_de_los_Pasantes()
    If Option2.Value = True Then
        If FGeneral = True Then
            Cargar_Datos (SQL)
            GoTo Reporte_General
        End If
    End If
Exit Sub

Reporte_General:
    CrystalReport2.ReportFileName = WDirecPrint + "Formulario_G.rpt"
    CrystalReport2.WindowTitle = Me.Caption
    CrystalReport2.Formulas(0) = "Campo1='" & CampoC1 & "'"
    CrystalReport2.Formulas(1) = "Campo2='" & CampoC2 & "'"
    CrystalReport2.Formulas(2) = "Campo3='" & CampoC3 & "'"
    CrystalReport2.Action = 1
End Sub

'-----------------------------------------------------------
' Tabla_Docentes                    =Docentes.rpt
'-----------------------------------------------------------
Sub Tabla_Docentes()
    If Option2.Value = True Then
        SQL = "SELECT  Docentes.NumeroD_ID, Docentes.Cedula, Docentes.Nombres, Docentes.Apellidos, Docentes.Cargo , Docentes.Telefono" & _
              " From Docentes"
    Else
        Dim Cedula As String
        Dim Buscar As String
        Cedula = InputBox("Introduzca La Cedula :", "Busqueda de Datos")
        If Cedula <> "" Then
            Buscar = "[Cedula]" & "=" & "'" & Cedula & "'"
            DDocentes.Recordset.FindFirst Buscar
            If DDocentes.Recordset.NoMatch Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
        '---------------------------------------------------------------------------------------------
        SQL = "SELECT  Docentes.NumeroD_ID, Docentes.Cedula, Docentes.Nombres, Docentes.Apellidos, Docentes.Cargo, Docentes.Telefono" & _
              " From Docentes" & _
              " WHERE Docentes.Cedula ='" & Cedula & "'"
    End If
'    Cargar_Datos2 (SQL)
    GoTo Reporte_General
Exit Sub

Reporte_General:
'---------------------------------------------------------------------------------------------
    Set Rep_Docentes.DataSource = DataEnvironment1.Connection1.Execute(SQL)
    Rep_Docentes.Caption = Me.Caption
    Rep_Docentes.Show
    
'    CrystalReport2.ReportFileName = WDirecPrint + "Docentes.rpt"
'    CrystalReport2.WindowTitle = Me.Caption
'    CrystalReport2.Action = 1
End Sub

'-----------------------------------------------------------
' Tabla_Registro_del_Pasante        =Formulario_E.rpt
'-----------------------------------------------------------
Sub Tabla_Registro_del_Pasante()
    If Option2.Value = True Then
        If FGeneral = True Then
            Cargar_Datos (SQL)
            GoTo Reporte_General
        End If
    End If
Exit Sub

Reporte_General:
    CrystalReport2.ReportFileName = WDirecPrint + "Formulario_E.rpt"
    CrystalReport2.WindowTitle = Me.Caption
    CrystalReport2.Formulas(0) = "Campo1='" & CampoC1 & "'"
    CrystalReport2.Formulas(1) = "Campo2='" & CampoC2 & "'"
    CrystalReport2.Formulas(2) = "Campo3='" & CampoC3 & "'"
    CrystalReport2.Action = 1
End Sub

'-----------------------------------------------------------
' Tabla_Registro_del_Tutor_Academico=Formulario_T.rpt
'-----------------------------------------------------------
Sub Tabla_Registro_del_Tutor_Academico()
    If Option2.Value = True Then
        If FGeneral = True Then
            Cargar_Datos (SQL)
            GoTo Reporte_General
        End If
    End If
Exit Sub

Reporte_General:
    CrystalReport2.ReportFileName = WDirecPrint + "Formulario_T.rpt"
    CrystalReport2.WindowTitle = Me.Caption
    CrystalReport2.Formulas(0) = "Campo1='" & CampoC1 & "'"
    CrystalReport2.Formulas(1) = "Campo2='" & CampoC2 & "'"
    CrystalReport2.Formulas(2) = "Campo3='" & CampoC3 & "'"
    CrystalReport2.Action = 1
End Sub

Sub Cargar_Datos(SQLcad As String)
    BaseD = Base_de_Datos
    BD = Abrir_BaseDatos(BaseD, 1)
    Set DAlumnos.Recordset = MAESTRO.CreateDynaset(SQLcad)
    '-----------------------------------------------
        If Not Adodc1_Alumnos.Recordset.EOF Then
            Adodc1_Alumnos.Recordset.MoveFirst
            Do Until Adodc1_Alumnos.Recordset.EOF
                If Not Adodc1_Alumnos.Recordset.EOF Then
                    Adodc1_Alumnos.Recordset.Delete
                End If
                Adodc1_Alumnos.Recordset.MoveNext
            Loop
        End If
        If Not DAlumnos.Recordset.EOF Then
            DAlumnos.Recordset.MoveFirst
            Do Until DAlumnos.Recordset.EOF
                If Not DAlumnos.Recordset.EOF Then
                    Adodc1_Alumnos.Recordset.AddNew
'---------------------------------------------------------------------------------------------
                    Campo1 = DAlumnos.Recordset.Fields("Cedula")
                    Campo2 = DAlumnos.Recordset.Fields("Seccion")
                    Campo3 = DAlumnos.Recordset.Fields("Apellidos") & ", " & DAlumnos.Recordset.Fields("Nombres")
                    Campo4 = DAlumnos.Recordset.Fields("Telefono")
                    Campo5 = DAlumnos.Recordset.Fields("Periodo")
                    Buscar = "[Codigo]" & "=" & "'" & Mid(Campo2, 1, 2) & "'"
                    DEspecialidad.Recordset.FindFirst Buscar
                    Campo6 = DEspecialidad.Recordset.Fields("Descripcion")
                    Select Case Mid(Campo2, 5, 1)
                    Case "1"
                        Campo7 = "Mañana"
                    Case "3"
                        Campo7 = "Noche"
                    End Select
                    Buscar = "[Cedula]" & "=" & "'" & Campo1 & "'"
                    DSolvencias.Recordset.FindFirst Buscar
                    CampoB1 = IIf(DSolvencias.Recordset.Fields("Taller") = True, 1, 0)
                    CampoB2 = IIf(DSolvencias.Recordset.Fields("Administrativo_Caja") = True, 1, 0)
                    CampoB3 = IIf(DSolvencias.Recordset.Fields("Academica_Control_de_Estudio") = True, 1, 0)
'---------------------------------------------------------------------------------------------
                    Adodc1_Alumnos.Recordset.Fields("Cedula") = Campo1
                    Adodc1_Alumnos.Recordset.Fields("ApeNom") = Campo3
                    Adodc1_Alumnos.Recordset.Fields("Seccion") = Campo2
                    Adodc1_Alumnos.Recordset.Fields("Telefono") = Campo4
                    Adodc1_Alumnos.Recordset.Fields("Periodo") = Campo5
                    Adodc1_Alumnos.Recordset.Fields("Descripcion") = Campo6
                    Adodc1_Alumnos.Recordset.Fields("Turno") = Campo7
'---------------------------------------------------------------------------------------------
                    Buscar = "[Cedula]" & "=" & "'" & Campo1 & "'"
                    DSolvencias.Recordset.FindFirst Buscar
                    If Not DSolvencias.Recordset.NoMatch Then
                        Adodc1_Alumnos.Recordset.Fields("Taller") = IIf(DSolvencias.Recordset.Fields("Taller") = True, 1, 0)
                        Adodc1_Alumnos.Recordset.Fields("Administrativo_Caja") = IIf(DSolvencias.Recordset.Fields("Administrativo_Caja") = True, 1, 0)
                        Adodc1_Alumnos.Recordset.Fields("Academica_Control_de_Estudio") = IIf(DSolvencias.Recordset.Fields("Academica_Control_de_Estudio") = True, 1, 0)
                        Adodc1_Alumnos.Recordset.Fields("Academica_Departamento_de_Grado") = IIf(DSolvencias.Recordset.Fields("Academica_Departamento_de_Grado") = True, 1, 0)
                        Adodc1_Alumnos.Recordset.Fields("Notas_Entregadas") = IIf(DSolvencias.Recordset.Fields("Notas_Entregadas") = True, 1, 0)
                        Adodc1_Alumnos.Recordset.Fields("Casos_Especiales") = IIf(DSolvencias.Recordset.Fields("Casos_Especiales") = True, 1, 0)
                        Adodc1_Alumnos.Recordset.Fields("Entrego_Carta_Aceptacion") = IIf(DSolvencias.Recordset.Fields("Entrego_Carta_Aceptacion") = True, 1, 0)
                        Adodc1_Alumnos.Recordset.Fields("Realizando_Pasantia") = IIf(DSolvencias.Recordset.Fields("Realizando_Pasantia") = True, 1, 0)
                    End If
'---------------------------------------------------------------------------------------------
                    Buscar = "[Cedula]" & "=" & "'" & Campo1 & "'"
                    DNotas.Recordset.FindFirst Buscar
                    If Not DNotas.Recordset.NoMatch Then
                        Adodc1_Alumnos.Recordset.Fields("Acum_60") = DNotas.Recordset.Fields("Acum_60")
                        Adodc1_Alumnos.Recordset.Fields("Acum_40") = DNotas.Recordset.Fields("Acum_40")
                        Adodc1_Alumnos.Recordset.Fields("Cedula_Tutor") = Trim(DNotas.Recordset.Fields("Cedula_Tutor"))
                        Adodc1_Alumnos.Recordset.Fields("Fecha_Entrega_Plan") = Trim(DNotas.Recordset.Fields("Fecha_Entrega_Plan"))
                        Adodc1_Alumnos.Recordset.Fields("Fecha_Aceptacion_Plan") = Trim(DNotas.Recordset.Fields("Fecha_Aceptacion_Plan"))
                        Adodc1_Alumnos.Recordset.Fields("Fecha_Entrega_Informe") = Trim(DNotas.Recordset.Fields("Fecha_Entrega_Informe"))
                        Adodc1_Alumnos.Recordset.Fields("Fecha_Aceptacion_Informe") = Trim(DNotas.Recordset.Fields("Fecha_Aceptacion_Informe"))
                        Adodc1_Alumnos.Recordset.Fields("Fecha_Final_Proceso") = Trim(DNotas.Recordset.Fields("Fecha_Final_Proceso"))
                        Adodc1_Alumnos.Recordset.Fields("Numero_Visitas") = DNotas.Recordset.Fields("Numero_Visitas")
                        Buscar = "[Cedula]" & "=" & "'" & Trim(DNotas.Recordset.Fields("Cedula_Tutor")) & "'"
                        DDocentes.Recordset.FindFirst Buscar
                        If Not DDocentes.Recordset.NoMatch Then
                            Adodc1_Alumnos.Recordset.Fields("ApeNom_Tutor") = DDocentes.Recordset.Fields("Apellidos") & "," & DDocentes.Recordset.Fields("Nombres")
                        End If
                    End If
'---------------------------------------------------------------------------------------------
                    Buscar = "[Cedula]" & "=" & "'" & Campo1 & "'"
                    DCentros_Pasantias.Recordset.FindFirst Buscar
                    If Not DCentros_Pasantias.Recordset.NoMatch Then
                        Adodc1_Alumnos.Recordset.Fields("Numero_Oficio") = DCentros_Pasantias.Recordset.Fields("Numero_Oficio")
                        Adodc1_Alumnos.Recordset.Fields("Horario") = Trim(DCentros_Pasantias.Recordset.Fields("Horario"))
                        Adodc1_Alumnos.Recordset.Fields("Nombre_Emp") = Trim(DCentros_Pasantias.Recordset.Fields("Nombre_Emp"))
                        Adodc1_Alumnos.Recordset.Fields("Tutor_Emp") = Trim(DCentros_Pasantias.Recordset.Fields("Tutor_Emp"))
                        Adodc1_Alumnos.Recordset.Fields("Fecha_Inicio") = Trim(DCentros_Pasantias.Recordset.Fields("Fecha_Inicio"))
                        Adodc1_Alumnos.Recordset.Fields("Fecha_Culminacion") = Trim(DCentros_Pasantias.Recordset.Fields("Fecha_Culminacion"))
                        Adodc1_Alumnos.Recordset.Fields("Telefono_Emp") = Trim(DCentros_Pasantias.Recordset.Fields("Telefono_Emp"))
                        Adodc1_Alumnos.Recordset.Fields("Direccion_Emp") = Trim(DCentros_Pasantias.Recordset.Fields("Direccion_Emp"))
                    End If
                    Adodc1_Alumnos.Recordset.Update
                End If
                DAlumnos.Recordset.MoveNext
            Loop
        End If
    DAlumnos.Recordset.Close
    MAESTRO.Close
End Sub

Sub Cargar_Datos2(SQLcad As String)
    BaseD = Base_de_Datos
    BD = Abrir_BaseDatos(BaseD, 1)
    Set DTablas.Recordset = MAESTRO.CreateDynaset(SQLcad)
    '-----------------------------------------------
        If Not Adodc1_Alumnos.Recordset.EOF Then
            Adodc1_Alumnos.Recordset.MoveFirst
            Do Until Adodc1_Alumnos.Recordset.EOF
                If Not Adodc1_Alumnos.Recordset.EOF Then
                    Adodc1_Alumnos.Recordset.Delete
                End If
                Adodc1_Alumnos.Recordset.MoveNext
            Loop
        End If
        If Not DTablas.Recordset.EOF Then
            DTablas.Recordset.MoveFirst
            Do Until DTablas.Recordset.EOF
                If Not DTablas.Recordset.EOF Then
                    Adodc1_Alumnos.Recordset.AddNew
'---------------------------------------------------------------------------------------------
                    Campo1 = DTablas.Recordset.Fields("Cedula")
                    Campo2 = DTablas.Recordset.Fields("Apellidos") & ", " & DTablas.Recordset.Fields("Nombres")
                    Campo3 = DTablas.Recordset.Fields("Cargo")
                    Adodc1_Alumnos.Recordset.Fields("Cedula") = Campo1
                    Adodc1_Alumnos.Recordset.Fields("ApeNom") = Campo2
                    Adodc1_Alumnos.Recordset.Fields("Descripcion") = Campo3
                    Adodc1_Alumnos.Recordset.Update
                End If
                DTablas.Recordset.MoveNext
            Loop
        End If
    DTablas.Recordset.Close
    MAESTRO.Close
End Sub

Private Sub cmdCerrar_Click()
    If cmdCerrar.Caption = "&Cerrar" Then
        Unload Me
    End If
End Sub

Private Sub Form_Activate()
    Adodc1_Alumnos.ConnectionString = DSN_Reporte_Mov
    Adodc1_Alumnos.RecordSource = "Alumnos"
    Adodc1_Alumnos.Refresh
    DAlumnos.DatabaseName = Base_de_Datos
    DAlumnos.Refresh
    DCentros_Pasantias.DatabaseName = Base_de_Datos
    DCentros_Pasantias.Refresh
    DControl.DatabaseName = Base_de_Datos
    DControl.Refresh
    DDocentes.DatabaseName = Base_de_Datos
    DDocentes.Refresh
    DEspecialidad.DatabaseName = Base_de_Datos
    DEspecialidad.Refresh
    DNotas.DatabaseName = Base_de_Datos
    DNotas.Refresh
    DSolvencias.DatabaseName = Base_de_Datos
    DSolvencias.Refresh
    DTabla_Gerenar.DatabaseName = Base_de_Datos
    DTabla_Gerenar.Refresh
    DTablas.DatabaseName = Base_de_Datos
    DTablas.Refresh
    WDirecPrint = App.Path & "\Reportes\"
    Select Case Tipo_Reporte
    Case "SOL":
        Option1.Enabled = False
    Case "REG":
        Option1.Enabled = False
    Case "TUR":
        Option1.Enabled = False
    Case "CAS":
        Option1.Enabled = False
    Case "TGR":
        Option1.Enabled = False
    Case "TGP":
        Option1.Enabled = False
    End Select
End Sub

Sub Reportes_de_Casos_Especiales() 'DISTINCT
    SQL = "SELECT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Centros_Pasantias].[Fecha_Inicio], [Centros_Pasantias].[Fecha_Culminacion], [Solvencias].[Realizando_Pasantia], [Solvencias].[Entrego_Carta_Aceptacion] " & _
    "FROM (Solvencias INNER JOIN (Alumnos INNER JOIN Centros_Pasantias ON [Alumnos].[Cedula]=[Centros_Pasantias].[Cedula]) ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
    "WHERE (([Solvencias].[Administrativo_Caja])=False) or (([Solvencias].[Casos_Especiales])=True) order by [Alumnos].[Seccion];"
    Set Rep_Casos_Especiales.DataSource = DataEnvironment1.Connection1.Execute(SQL)
    Rep_Casos_Especiales.Caption = Me.Caption
    Rep_Casos_Especiales.Show
End Sub

Sub Reportes_Total_por_Seccion()
    SQL = "SELECT DISTINCT [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, COUNT([Alumnos].[Seccion]) AS Total_X_Seccion " & _
    "FROM (Solvencias INNER JOIN (Alumnos INNER JOIN Centros_Pasantias ON [Alumnos].[Cedula]=[Centros_Pasantias].[Cedula]) ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
    "GROUP BY [Alumnos].[Seccion], [Especialidad].[Descripcion] " & _
    "HAVING COUNT([Alumnos].[Seccion]);"
    Set Rep_Total_por_Seccion.DataSource = DataEnvironment1.Connection1.Execute(SQL)
    Rep_Total_por_Seccion.Show
    Rep_Total_por_Seccion.Caption = Me.Caption
    Rep_Total_por_Seccion.Show
End Sub

Sub Reportes_Total_por_Seccion_P()
    SQL = "SELECT DISTINCT [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, COUNT([Alumnos].[Seccion]) AS Total_X_Seccion " & _
    "FROM (Solvencias INNER JOIN Alumnos ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
    "Where ([Solvencias].[Realizando_Pasantia] = False) And ([Solvencias].[Administrativo_Caja] = True) " & _
    "GROUP BY [Alumnos].[Seccion], [Especialidad].[Descripcion] " & _
    "HAVING COUNT([Alumnos].[Seccion]);"
    Set Rep_Total_por_Seccion_Pendiente.DataSource = DataEnvironment1.Connection1.Execute(SQL)
    Rep_Total_por_Seccion_Pendiente.Show
    Rep_Total_por_Seccion_Pendiente.Caption = Me.Caption
    Rep_Total_por_Seccion_Pendiente.Show
End Sub

Sub Reportes_Centros_de_Pasantias()
    If Option2.Value = True Then
        If FGeneral = True Then
            Set Rep_Centros_de_Pasantias.DataSource = DataEnvironment1.Connection1.Execute(SQL)
            Rep_Centros_de_Pasantias.Caption = Me.Caption
            Rep_Centros_de_Pasantias.Show
        End If
    Else
        GoTo Reporte_Individual
    End If

Exit Sub
Reporte_Individual:
    Dim Empresa, Buscar As String
    Empresa = InputBox("Introduzca el Nombre de la Empresa :", "Busqueda de Datos")
    If Empresa <> "" Then
        SQL = "SELECT Alumnos.Nacionalidad, Alumnos.Cedula, Alumnos.Nombres, Alumnos.Apellidos, Alumnos.Seccion,[Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Especialidad].[Descripcion] " & _
        "FROM (Centros_Pasantias INNER JOIN Alumnos ON [Centros_Pasantias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
        "WHERE Centros_Pasantias.Nombre_Emp LIKE '" & UCase(Empresa) & "' " & _
        "ORDER BY [Centros_Pasantias].[Nombre_Emp];"
    Else
        SQL = "SELECT Alumnos.Nacionalidad, Alumnos.Cedula, Alumnos.Nombres, Alumnos.Apellidos, Alumnos.Seccion,[Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Especialidad].[Descripcion] " & _
        "FROM (Centros_Pasantias INNER JOIN Alumnos ON [Centros_Pasantias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
        "ORDER BY [Centros_Pasantias].[Nombre_Emp];"
    End If
'---------------------------------------------------------------------------------------------
    Set Rep_Centros_de_Pasantias.DataSource = DataEnvironment1.Connection1.Execute(SQL)
    Rep_Centros_de_Pasantias.Caption = Me.Caption
    Rep_Centros_de_Pasantias.Show
End Sub

Sub Reportes_Pendientes_X_Pasantías()
    If Option2.Value = True Then
        If FGeneral = True Then
            Set Rep_Pendientes_X_Pasantías.DataSource = DataEnvironment1.Connection1.Execute(SQL)
            Rep_Pendientes_X_Pasantías.Caption = Me.Caption
            Rep_Pendientes_X_Pasantías.Show
        End If
    Else
        GoTo Reporte_Individual
    End If

Exit Sub
Reporte_Individual:
    Dim Buscar As String
    Buscar = InputBox("Indique la Sección :", "Buscar por Sección")
    If Buscar <> "" Then
        SQL = "SELECT DISTINCT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Solvencias].[Realizando_Pasantia], IIf([Solvencias].[Administrativo_Caja]=True,' ','*') AS Solvente " & _
        "FROM (Solvencias INNER JOIN Alumnos ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
        "WHERE ([Solvencias].[Realizando_Pasantia]=False) And ([Solvencias].[Administrativo_Caja]=True) and [Alumnos].[Seccion] like '" & Buscar & "' " & _
        "ORDER BY [Alumnos].[Seccion];"
    Else
        SQL = "SELECT DISTINCT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Solvencias].[Realizando_Pasantia], IIf([Solvencias].[Administrativo_Caja]=True,' ','*') AS Solvente " & _
        "FROM (Solvencias INNER JOIN Alumnos ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
        "WHERE ([Solvencias].[Realizando_Pasantia]=False) And ([Solvencias].[Administrativo_Caja]=True) " & _
        "ORDER BY [Alumnos].[Seccion];"
    End If
'---------------------------------------------------------------------------------------------
    Set Rep_Pendientes_X_Pasantías.DataSource = DataEnvironment1.Connection1.Execute(SQL)
    Rep_Pendientes_X_Pasantías.Caption = Me.Caption
    Rep_Pendientes_X_Pasantías.Show
End Sub

Sub Reportes_Realizando_Pasantías()
    If Option2.Value = True Then
        If FGeneral = True Then
            Set Rep_Realizando_Pasantías.DataSource = DataEnvironment1.Connection1.Execute(SQL)
            Rep_Realizando_Pasantías.Caption = Me.Caption
            Rep_Realizando_Pasantías.Show
        End If
    Else
        GoTo Reporte_Individual
    End If

Exit Sub
Reporte_Individual:
    Dim Buscar As String
    Buscar = InputBox("Indique la Sección :", "Buscar por Sección")
    If Buscar <> "" Then
        SQL = "SELECT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Centros_Pasantias].[Fecha_Inicio], [Centros_Pasantias].[Fecha_Culminacion], [Solvencias].[Realizando_Pasantia], [Solvencias].[Entrego_Carta_Aceptacion]" & _
        "FROM (Solvencias INNER JOIN (Alumnos INNER JOIN Centros_Pasantias ON [Alumnos].[Cedula]=[Centros_Pasantias].[Cedula]) ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo]" & _
        "Where [Alumnos].[Seccion] Like '" & Buscar & "' " & _
        "ORDER BY [Alumnos].[Apellidos];"
    Else
        SQL = "SELECT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Centros_Pasantias].[Fecha_Inicio], [Centros_Pasantias].[Fecha_Culminacion], [Solvencias].[Realizando_Pasantia], [Solvencias].[Entrego_Carta_Aceptacion]" & _
        "FROM (Solvencias INNER JOIN (Alumnos INNER JOIN Centros_Pasantias ON [Alumnos].[Cedula]=[Centros_Pasantias].[Cedula]) ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo]" & _
        "ORDER BY [Alumnos].[Apellidos];"
    End If
'---------------------------------------------------------------------------------------------
    Set Rep_Realizando_Pasantías.DataSource = DataEnvironment1.Connection1.Execute(SQL)
    Rep_Realizando_Pasantías.Caption = Me.Caption
    Rep_Realizando_Pasantías.Show
End Sub

Function FGeneral() As Boolean
    Select Case Tipo_Reporte
'    Case "ALU":
'    Case "FIC":
'    Case "SOL":
'    Case "DOC":
'    Case "REG":
'    Case "TUR":
'    Case "CAS":
'    Case "TGR":
'    Case "TGP":
    Case "RPP":
        frmGenerarReport.Show vbModal
        If frmGenerarReport.Boton_Respuesta Then
            ' Consulta General Periodo,Seccion,Especialidad
            ' [*,*,*]
            If frmGenerarReport.pPeriodo = "*" And _
                frmGenerarReport.pSeccion = "*" And _
                frmGenerarReport.pEspecialidad = "*" Then
                CampoC1 = "Todos los Periodo"
                CampoC2 = "Todas las Secciones"
                CampoC3 = "Todas las Especialidad"
                SQL = "SELECT DISTINCT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Solvencias].[Realizando_Pasantia], IIf([Solvencias].[Administrativo_Caja]=True,' ','*') AS Solvente " & _
                "FROM (Solvencias INNER JOIN Alumnos ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
                "WHERE ([Solvencias].[Realizando_Pasantia]=False) And ([Solvencias].[Administrativo_Caja]=True) " & _
                "ORDER BY [Alumnos].[Seccion];"
            End If
            ' Consulta por Periodo General
            ' [<>,*,*]
            If frmGenerarReport.pPeriodo <> "*" And _
                frmGenerarReport.pSeccion = "*" And _
                frmGenerarReport.pEspecialidad = "*" Then
                CampoC1 = "Periodo=" & frmGenerarReport.pPeriodo
                CampoC2 = "Todas las Secciones"
                CampoC3 = "Todas las Especialidad"
                SQL = "SELECT DISTINCT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Solvencias].[Realizando_Pasantia], IIf([Solvencias].[Administrativo_Caja]=True,' ','*') AS Solvente " & _
                "FROM (Solvencias INNER JOIN Alumnos ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
                "WHERE ([Solvencias].[Realizando_Pasantia]=False) And ([Solvencias].[Administrativo_Caja]=True)"
                    SQL = SQL & " and [Alumnos].[Periodo]='" & frmGenerarReport.pPeriodo & "' " & _
                "ORDER BY [Alumnos].[Seccion];"
            End If
            ' Consulta por Periodo y Seccion
            ' [<>,<>,*]
            If frmGenerarReport.pPeriodo <> "*" And _
                frmGenerarReport.pSeccion <> "*" Then
                CampoC1 = "Periodo=" & frmGenerarReport.pPeriodo
                CampoC2 = "Sección=" & frmGenerarReport.pSeccion
                CampoC3 = "Especialidad=" & frmGenerarReport.pEspecialidad
                SQL = "SELECT DISTINCT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Solvencias].[Realizando_Pasantia], IIf([Solvencias].[Administrativo_Caja]=True,' ','*') AS Solvente " & _
                "FROM (Solvencias INNER JOIN Alumnos ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
                "WHERE ([Solvencias].[Realizando_Pasantia]=False) And ([Solvencias].[Administrativo_Caja]=True)"
                    SQL = SQL & " and [Alumnos].[Periodo]='" & frmGenerarReport.pPeriodo & "'"
                    SQL = SQL & " and [Alumnos].[Seccion]='" & frmGenerarReport.pSeccion & "' " & _
                "ORDER BY [Alumnos].[Seccion];"
            End If
            ' Consulta por solo Seccion
            ' [*,<>,*] y [*,<>,<>]
            If frmGenerarReport.pPeriodo = "*" And _
                frmGenerarReport.pSeccion <> "*" Then
                CampoC1 = "Todos los Periodo"
                CampoC2 = "Sección=" & frmGenerarReport.pSeccion
                CampoC3 = "Especialidad=" & frmGenerarReport.pEspecialidad
                SQL = "SELECT DISTINCT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Solvencias].[Realizando_Pasantia], IIf([Solvencias].[Administrativo_Caja]=True,' ','*') AS Solvente " & _
                "FROM (Solvencias INNER JOIN Alumnos ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
                "WHERE ([Solvencias].[Realizando_Pasantia]=False) And ([Solvencias].[Administrativo_Caja]=True)"
                    SQL = SQL & " and [Alumnos].[Seccion]='" & frmGenerarReport.pSeccion & "' " & _
                "ORDER BY [Alumnos].[Seccion];"
            End If
            ' Consulta por Periodo y Especialidad
            ' [<>,*,<>]
            If frmGenerarReport.pPeriodo <> "*" And _
                frmGenerarReport.pSeccion = "*" And _
                frmGenerarReport.pEspecialidad <> "*" Then
                CampoC1 = "Periodo=" & frmGenerarReport.pPeriodo
                CampoC2 = "Todas las Secciones"
                CampoC3 = "Especialidad=" & frmGenerarReport.pEspecialidad
                SQL = "SELECT DISTINCT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Solvencias].[Realizando_Pasantia], IIf([Solvencias].[Administrativo_Caja]=True,' ','*') AS Solvente " & _
                "FROM (Solvencias INNER JOIN Alumnos ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
                "WHERE ([Solvencias].[Realizando_Pasantia]=False) And ([Solvencias].[Administrativo_Caja]=True)"
                    SQL = SQL & " and [Alumnos].[Periodo]='" & frmGenerarReport.pPeriodo & "' "
                    SQL = SQL & " and [Especialidad].[Descripcion]='" & frmGenerarReport.pEspecialidad & "' " & _
                "ORDER BY [Alumnos].[Seccion];"
            End If
            ' Consulta por solo Especialidad
            ' [*,*,<>]
            If frmGenerarReport.pPeriodo = "*" And _
                frmGenerarReport.pSeccion = "*" And _
                frmGenerarReport.pEspecialidad <> "*" Then
                CampoC1 = "Todos los Periodo"
                CampoC2 = "Todas las Secciones"
                CampoC3 = "Especialidad=" & frmGenerarReport.pEspecialidad
                SQL = "SELECT DISTINCT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Solvencias].[Realizando_Pasantia], IIf([Solvencias].[Administrativo_Caja]=True,' ','*') AS Solvente " & _
                "FROM (Solvencias INNER JOIN Alumnos ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
                "WHERE ([Solvencias].[Realizando_Pasantia]=False) And ([Solvencias].[Administrativo_Caja]=True)"
                SQL = SQL & " and [Especialidad].[Descripcion]='" & frmGenerarReport.pEspecialidad & "' " & _
                "ORDER BY [Alumnos].[Seccion];"
            End If
            FGeneral = True
        Else
            FGeneral = False
        End If
    Case "REP":
        frmGenerarReport.Show vbModal
        If frmGenerarReport.Boton_Respuesta Then
            ' Consulta General Periodo,Seccion,Especialidad
            ' [*,*,*]
            If frmGenerarReport.pPeriodo = "*" And _
                frmGenerarReport.pSeccion = "*" And _
                frmGenerarReport.pEspecialidad = "*" Then
                CampoC1 = "Todos los Periodo"
                CampoC2 = "Todas las Secciones"
                CampoC3 = "Todas las Especialidad"
                SQL = "SELECT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Centros_Pasantias].[Fecha_Inicio], [Centros_Pasantias].[Fecha_Culminacion], [Solvencias].[Realizando_Pasantia], [Solvencias].[Entrego_Carta_Aceptacion]" & _
                "FROM (Solvencias INNER JOIN (Alumnos INNER JOIN Centros_Pasantias ON [Alumnos].[Cedula]=[Centros_Pasantias].[Cedula]) ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo]" & _
                "Where [Alumnos].[Seccion] Like '" & Buscar & "' " & _
                "ORDER BY [Alumnos].[Apellidos];"
            End If
            ' Consulta por Periodo General
            ' [<>,*,*]
            If frmGenerarReport.pPeriodo <> "*" And _
                frmGenerarReport.pSeccion = "*" And _
                frmGenerarReport.pEspecialidad = "*" Then
                CampoC1 = "Periodo=" & frmGenerarReport.pPeriodo
                CampoC2 = "Todas las Secciones"
                CampoC3 = "Todas las Especialidad"
                SQL = "SELECT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Centros_Pasantias].[Fecha_Inicio], [Centros_Pasantias].[Fecha_Culminacion], [Solvencias].[Realizando_Pasantia], [Solvencias].[Entrego_Carta_Aceptacion]" & _
                "FROM (Solvencias INNER JOIN (Alumnos INNER JOIN Centros_Pasantias ON [Alumnos].[Cedula]=[Centros_Pasantias].[Cedula]) ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo]" & _
                "WHERE "
                    SQL = SQL & "[Alumnos].[Periodo]='" & frmGenerarReport.pPeriodo & "' " & _
                "ORDER BY [Alumnos].[Apellidos];"
            End If
            ' Consulta por Periodo y Seccion
            ' [<>,<>,*]
            If frmGenerarReport.pPeriodo <> "*" And _
                frmGenerarReport.pSeccion <> "*" Then
                CampoC1 = "Periodo=" & frmGenerarReport.pPeriodo
                CampoC2 = "Sección=" & frmGenerarReport.pSeccion
                CampoC3 = "Especialidad=" & frmGenerarReport.pEspecialidad
                SQL = "SELECT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Centros_Pasantias].[Fecha_Inicio], [Centros_Pasantias].[Fecha_Culminacion], [Solvencias].[Realizando_Pasantia], [Solvencias].[Entrego_Carta_Aceptacion]" & _
                "FROM (Solvencias INNER JOIN (Alumnos INNER JOIN Centros_Pasantias ON [Alumnos].[Cedula]=[Centros_Pasantias].[Cedula]) ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo]" & _
                "WHERE "
                    SQL = SQL & "[Alumnos].[Periodo]='" & frmGenerarReport.pPeriodo & "'"
                    SQL = SQL & " and [Alumnos].[Seccion]='" & frmGenerarReport.pSeccion & "' " & _
                "ORDER BY [Alumnos].[Apellidos];"
            End If
            ' Consulta por solo Seccion
            ' [*,<>,*] y [*,<>,<>]
            If frmGenerarReport.pPeriodo = "*" And _
                frmGenerarReport.pSeccion <> "*" Then
                CampoC1 = "Todos los Periodo"
                CampoC2 = "Sección=" & frmGenerarReport.pSeccion
                CampoC3 = "Especialidad=" & frmGenerarReport.pEspecialidad
                SQL = "SELECT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Centros_Pasantias].[Fecha_Inicio], [Centros_Pasantias].[Fecha_Culminacion], [Solvencias].[Realizando_Pasantia], [Solvencias].[Entrego_Carta_Aceptacion]" & _
                "FROM (Solvencias INNER JOIN (Alumnos INNER JOIN Centros_Pasantias ON [Alumnos].[Cedula]=[Centros_Pasantias].[Cedula]) ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo]" & _
                "WHERE "
                    SQL = SQL & "[Alumnos].[Seccion]='" & frmGenerarReport.pSeccion & "' " & _
                "ORDER BY [Alumnos].[Apellidos];"
            End If
            ' Consulta por Periodo y Especialidad
            ' [<>,*,<>]
            If frmGenerarReport.pPeriodo <> "*" And _
                frmGenerarReport.pSeccion = "*" And _
                frmGenerarReport.pEspecialidad <> "*" Then
                CampoC1 = "Periodo=" & frmGenerarReport.pPeriodo
                CampoC2 = "Todas las Secciones"
                CampoC3 = "Especialidad=" & frmGenerarReport.pEspecialidad
                SQL = "SELECT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Centros_Pasantias].[Fecha_Inicio], [Centros_Pasantias].[Fecha_Culminacion], [Solvencias].[Realizando_Pasantia], [Solvencias].[Entrego_Carta_Aceptacion]" & _
                "FROM (Solvencias INNER JOIN (Alumnos INNER JOIN Centros_Pasantias ON [Alumnos].[Cedula]=[Centros_Pasantias].[Cedula]) ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo]" & _
                "WHERE "
                    SQL = SQL & "[Alumnos].[Periodo]='" & frmGenerarReport.pPeriodo & "' "
                    SQL = SQL & " and [Especialidad].[Descripcion]='" & frmGenerarReport.pEspecialidad & "' " & _
                "ORDER BY [Alumnos].[Apellidos];"
            End If
            ' Consulta por solo Especialidad
            ' [*,*,<>]
            If frmGenerarReport.pPeriodo = "*" And _
                frmGenerarReport.pSeccion = "*" And _
                frmGenerarReport.pEspecialidad <> "*" Then
                CampoC1 = "Todos los Periodo"
                CampoC2 = "Todas las Secciones"
                CampoC3 = "Especialidad=" & frmGenerarReport.pEspecialidad
                SQL = "SELECT [Alumnos].[Cedula], [Alumnos].[Nombres], [Alumnos].[Apellidos], [Alumnos].[Seccion], [Especialidad].[Descripcion], IIF(MID([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, [Alumnos].[Telefono], [Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Centros_Pasantias].[Fecha_Inicio], [Centros_Pasantias].[Fecha_Culminacion], [Solvencias].[Realizando_Pasantia], [Solvencias].[Entrego_Carta_Aceptacion]" & _
                "FROM (Solvencias INNER JOIN (Alumnos INNER JOIN Centros_Pasantias ON [Alumnos].[Cedula]=[Centros_Pasantias].[Cedula]) ON [Solvencias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo]" & _
                "WHERE "
                SQL = SQL & "[Especialidad].[Descripcion]='" & frmGenerarReport.pEspecialidad & "' " & _
                "ORDER BY [Alumnos].[Apellidos];"
            End If
            FGeneral = True
        Else
            FGeneral = False
        End If
    Case "CEN":
        frmGenerarReport.Show vbModal
        If frmGenerarReport.Boton_Respuesta Then
            ' Consulta General Periodo,Seccion,Especialidad
            ' [*,*,*]
            If frmGenerarReport.pPeriodo = "*" And _
                frmGenerarReport.pSeccion = "*" And _
                frmGenerarReport.pEspecialidad = "*" Then
                CampoC1 = "Todos los Periodo"
                CampoC2 = "Todas las Secciones"
                CampoC3 = "Todas las Especialidad"
                SQL = "SELECT [Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Especialidad].[Descripcion] " & _
                    "FROM (Centros_Pasantias INNER JOIN Alumnos ON [Centros_Pasantias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
                    "ORDER BY [Centros_Pasantias].[Nombre_Emp];"
            End If
            ' Consulta por Periodo General
            ' [<>,*,*]
            If frmGenerarReport.pPeriodo <> "*" And _
                frmGenerarReport.pSeccion = "*" And _
                frmGenerarReport.pEspecialidad = "*" Then
                CampoC1 = "Periodo=" & frmGenerarReport.pPeriodo
                CampoC2 = "Todas las Secciones"
                CampoC3 = "Todas las Especialidad"
                SQL = "SELECT [Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Especialidad].[Descripcion] " & _
                    "FROM (Centros_Pasantias INNER JOIN Alumnos ON [Centros_Pasantias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
                    " WHERE (((Mid([Alumnos].[Seccion],1,2))=[Especialidad].[Codigo]))"
                    SQL = SQL & " and [Alumnos].[Periodo]='" & frmGenerarReport.pPeriodo & "'" & _
                    "ORDER BY [Centros_Pasantias].[Nombre_Emp];"
            End If
            ' Consulta por Periodo y Seccion
            ' [<>,<>,*]
            If frmGenerarReport.pPeriodo <> "*" And _
                frmGenerarReport.pSeccion <> "*" Then
                CampoC1 = "Periodo=" & frmGenerarReport.pPeriodo
                CampoC2 = "Sección=" & frmGenerarReport.pSeccion
                CampoC3 = "Especialidad=" & frmGenerarReport.pEspecialidad
                SQL = "SELECT [Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Especialidad].[Descripcion] " & _
                    "FROM (Centros_Pasantias INNER JOIN Alumnos ON [Centros_Pasantias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
                    " WHERE (((Mid([Alumnos].[Seccion],1,2))=[Especialidad].[Codigo]))"
                    SQL = SQL & " and [Alumnos].[Periodo]='" & frmGenerarReport.pPeriodo & "'"
                    SQL = SQL & " and [Alumnos].[Seccion]='" & frmGenerarReport.pSeccion & "' " & _
                    "ORDER BY [Centros_Pasantias].[Nombre_Emp];"
            End If
            ' Consulta por solo Seccion
            ' [*,<>,*] y [*,<>,<>]
            If frmGenerarReport.pPeriodo = "*" And _
                frmGenerarReport.pSeccion <> "*" Then
                CampoC1 = "Todos los Periodo"
                CampoC2 = "Sección=" & frmGenerarReport.pSeccion
                CampoC3 = "Especialidad=" & frmGenerarReport.pEspecialidad
                SQL = "SELECT [Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Especialidad].[Descripcion] " & _
                    "FROM (Centros_Pasantias INNER JOIN Alumnos ON [Centros_Pasantias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
                    " WHERE (((Mid([Alumnos].[Seccion],1,2))=[Especialidad].[Codigo]))"
                    SQL = SQL & " and [Alumnos].[Seccion]='" & frmGenerarReport.pSeccion & "' " & _
                    "ORDER BY [Centros_Pasantias].[Nombre_Emp];"
            End If
            ' Consulta por Periodo y Especialidad
            ' [<>,*,<>]
            If frmGenerarReport.pPeriodo <> "*" And _
                frmGenerarReport.pSeccion = "*" And _
                frmGenerarReport.pEspecialidad <> "*" Then
                CampoC1 = "Periodo=" & frmGenerarReport.pPeriodo
                CampoC2 = "Todas las Secciones"
                CampoC3 = "Especialidad=" & frmGenerarReport.pEspecialidad
                SQL = "SELECT [Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Especialidad].[Descripcion] " & _
                    "FROM (Centros_Pasantias INNER JOIN Alumnos ON [Centros_Pasantias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
                    " WHERE (((Mid([Alumnos].[Seccion],1,2))=[Especialidad].[Codigo]))"
                    SQL = SQL & " and [Alumnos].[Periodo]='" & frmGenerarReport.pPeriodo & "' "
                    SQL = SQL & " and [Especialidad].[Descripcion]='" & frmGenerarReport.pEspecialidad & "' " & _
                    "ORDER BY [Centros_Pasantias].[Nombre_Emp];"
            End If
            ' Consulta por solo Especialidad
            ' [*,*,<>]
            If frmGenerarReport.pPeriodo = "*" And _
                frmGenerarReport.pSeccion = "*" And _
                frmGenerarReport.pEspecialidad <> "*" Then
                CampoC1 = "Todos los Periodo"
                CampoC2 = "Todas las Secciones"
                CampoC3 = "Especialidad=" & frmGenerarReport.pEspecialidad
                SQL = "SELECT [Centros_Pasantias].[Nombre_Emp], [Centros_Pasantias].[Tutor_Emp], [Centros_Pasantias].[Telefono_Emp], [Centros_Pasantias].[Direccion_Emp], [Centros_Pasantias].[Horario], [Especialidad].[Descripcion] " & _
                    "FROM (Centros_Pasantias INNER JOIN Alumnos ON [Centros_Pasantias].[Cedula]=[Alumnos].[Cedula]) INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2)=[Especialidad].[Codigo] " & _
                    " WHERE (((Mid([Alumnos].[Seccion],1,2))=[Especialidad].[Codigo]))"
                SQL = SQL & " and [Especialidad].[Descripcion]='" & frmGenerarReport.pEspecialidad & "' " & _
                    "ORDER BY [Centros_Pasantias].[Nombre_Emp];"
            End If
            FGeneral = True
        Else
            FGeneral = False
        End If
    Case Else
        frmGenerarReport.Show vbModal
        If frmGenerarReport.Boton_Respuesta Then
            ' Consulta General Periodo,Seccion,Especialidad
            ' [*,*,*]
            If frmGenerarReport.pPeriodo = "*" And _
                frmGenerarReport.pSeccion = "*" And _
                frmGenerarReport.pEspecialidad = "*" Then
                CampoC1 = "Todos los Periodo"
                CampoC2 = "Todas las Secciones"
                CampoC3 = "Todas las Especialidad"
                SQL = "SELECT Alumnos.Nacionalidad, Alumnos.Cedula, Alumnos.Nombres, Alumnos.Apellidos, Alumnos.Seccion, Especialidad.Descripcion," & _
                      "IIf(Mid([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, Alumnos.Telefono, Alumnos.Direccion, Alumnos.Periodo " & _
                      "FROM Alumnos INNER JOIN Especialidad ON MID([Alumnos].[Seccion],1,2) = Especialidad.Codigo ORDER BY Alumnos.Seccion;"
            End If
            ' Consulta por Periodo General
            ' [<>,*,*]
            If frmGenerarReport.pPeriodo <> "*" And _
                frmGenerarReport.pSeccion = "*" And _
                frmGenerarReport.pEspecialidad = "*" Then
                CampoC1 = "Periodo=" & frmGenerarReport.pPeriodo
                CampoC2 = "Todas las Secciones"
                CampoC3 = "Todas las Especialidad"
    '            SQL = "SELECT Alumnos.Nacionalidad,Alumnos.Cedula, Alumnos.Seccion, Alumnos.Apellidos, Alumnos.Nombres, Alumnos.Telefono, Alumnos.Periodo, Especialidad.Descripcion"
                SQL = "SELECT Alumnos.Nacionalidad, Alumnos.Cedula, Alumnos.Nombres, Alumnos.Apellidos, Alumnos.Seccion, Especialidad.Descripcion," & _
                      "IIf(Mid([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, Alumnos.Telefono, Alumnos.Direccion, Alumnos.Periodo " & _
                      " From Especialidad, Alumnos " & _
                      " WHERE (((Mid([Alumnos].[Seccion],1,2))=[Especialidad].[Codigo]))"
                SQL = SQL & " and [Alumnos].[Periodo]='" & frmGenerarReport.pPeriodo & "' ORDER BY Alumnos.Cedula+Alumnos.Seccion;"
            End If
            ' Consulta por Periodo y Seccion
            ' [<>,<>,*]
            If frmGenerarReport.pPeriodo <> "*" And _
                frmGenerarReport.pSeccion <> "*" Then
                CampoC1 = "Periodo=" & frmGenerarReport.pPeriodo
                CampoC2 = "Sección=" & frmGenerarReport.pSeccion
                CampoC3 = "Especialidad=" & frmGenerarReport.pEspecialidad
    '            SQL = "SELECT Alumnos.Nacionalidad,Alumnos.Cedula, Alumnos.Seccion, Alumnos.Apellidos, Alumnos.Nombres, Alumnos.Telefono, Alumnos.Periodo, Especialidad.Descripcion"
                SQL = "SELECT Alumnos.Nacionalidad, Alumnos.Cedula, Alumnos.Nombres, Alumnos.Apellidos, Alumnos.Seccion, Especialidad.Descripcion," & _
                      "IIf(Mid([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, Alumnos.Telefono, Alumnos.Direccion, Alumnos.Periodo " & _
                      " From Especialidad, Alumnos " & _
                      " WHERE (((Mid([Alumnos].[Seccion],1,2))=[Especialidad].[Codigo]))"
                SQL = SQL & " and [Alumnos].[Periodo]='" & frmGenerarReport.pPeriodo & "'"
                SQL = SQL & " and [Alumnos].[Seccion]='" & frmGenerarReport.pSeccion & "' ORDER BY Alumnos.Cedula+Alumnos.Seccion;"
            End If
            ' Consulta por solo Seccion
            ' [*,<>,*] y [*,<>,<>]
            If frmGenerarReport.pPeriodo = "*" And _
                frmGenerarReport.pSeccion <> "*" Then
                CampoC1 = "Todos los Periodo"
                CampoC2 = "Sección=" & frmGenerarReport.pSeccion
                CampoC3 = "Especialidad=" & frmGenerarReport.pEspecialidad
    '            SQL = "SELECT Alumnos.Nacionalidad,Alumnos.Cedula, Alumnos.Seccion, Alumnos.Apellidos, Alumnos.Nombres, Alumnos.Telefono, Alumnos.Periodo, Especialidad.Descripcion"
                SQL = "SELECT Alumnos.Nacionalidad, Alumnos.Cedula, Alumnos.Nombres, Alumnos.Apellidos, Alumnos.Seccion, Especialidad.Descripcion," & _
                      "IIf(Mid([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, Alumnos.Telefono, Alumnos.Direccion, Alumnos.Periodo " & _
                      " From Especialidad, Alumnos " & _
                      " WHERE (((Mid([Alumnos].[Seccion],1,2))=[Especialidad].[Codigo]))"
                SQL = SQL & " and [Alumnos].[Seccion]='" & frmGenerarReport.pSeccion & "' ORDER BY Alumnos.Cedula+Alumnos.Seccion;"
            End If
            ' Consulta por Periodo y Especialidad
            ' [<>,*,<>]
            If frmGenerarReport.pPeriodo <> "*" And _
                frmGenerarReport.pSeccion = "*" And _
                frmGenerarReport.pEspecialidad <> "*" Then
                CampoC1 = "Periodo=" & frmGenerarReport.pPeriodo
                CampoC2 = "Todas las Secciones"
                CampoC3 = "Especialidad=" & frmGenerarReport.pEspecialidad
    '            SQL = "SELECT Alumnos.Nacionalidad,Alumnos.Cedula, Alumnos.Seccion, Alumnos.Apellidos, Alumnos.Nombres, Alumnos.Telefono, Alumnos.Periodo, Especialidad.Descripcion"
                SQL = "SELECT Alumnos.Nacionalidad, Alumnos.Cedula, Alumnos.Nombres, Alumnos.Apellidos, Alumnos.Seccion, Especialidad.Descripcion," & _
                      "IIf(Mid([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, Alumnos.Telefono, Alumnos.Direccion, Alumnos.Periodo " & _
                      " From Especialidad, Alumnos " & _
                      " WHERE (((Mid([Alumnos].[Seccion],1,2))=[Especialidad].[Codigo]))"
                SQL = SQL & " and [Alumnos].[Periodo]='" & frmGenerarReport.pPeriodo & "' "
                SQL = SQL & " and [Especialidad].[Descripcion]='" & frmGenerarReport.pEspecialidad & "' ORDER BY Alumnos.Cedula+Alumnos.Seccion;"
            End If
            ' Consulta por solo Especialidad
            ' [*,*,<>]
            If frmGenerarReport.pPeriodo = "*" And _
                frmGenerarReport.pSeccion = "*" And _
                frmGenerarReport.pEspecialidad <> "*" Then
                CampoC1 = "Todos los Periodo"
                CampoC2 = "Todas las Secciones"
                CampoC3 = "Especialidad=" & frmGenerarReport.pEspecialidad
    '            SQL = "SELECT Alumnos.Nacionalidad,Alumnos.Cedula, Alumnos.Seccion, Alumnos.Apellidos, Alumnos.Nombres, Alumnos.Telefono, Alumnos.Periodo, Especialidad.Descripcion"
                SQL = "SELECT Alumnos.Nacionalidad, Alumnos.Cedula, Alumnos.Nombres, Alumnos.Apellidos, Alumnos.Seccion, Especialidad.Descripcion," & _
                      "IIf(Mid([Alumnos].[Seccion],5,1)='3','Noche','Mañana') AS Turno, Alumnos.Telefono, Alumnos.Direccion, Alumnos.Periodo " & _
                      " From Especialidad, Alumnos " & _
                      " WHERE (((Mid([Alumnos].[Seccion],1,2))=[Especialidad].[Codigo]))"
                SQL = SQL & " and [Especialidad].[Descripcion]='" & frmGenerarReport.pEspecialidad & "' ORDER BY Alumnos.Cedula+Alumnos.Seccion;"
            End If
            FGeneral = True
        Else
            FGeneral = False
        End If
    End Select
End Function
