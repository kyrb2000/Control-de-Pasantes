VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmAlumnos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alumnos"
   ClientHeight    =   5790
   ClientLeft      =   1350
   ClientTop       =   1575
   ClientWidth     =   9255
   Icon            =   "frmAlumnos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5790
   ScaleWidth      =   9255
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   4215
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   9255
      Begin VB.ComboBox txtNacionalidad 
         DataField       =   "Nacionalidad"
         DataSource      =   "mdbAlumnos"
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
         ItemData        =   "frmAlumnos.frx":0442
         Left            =   1200
         List            =   "frmAlumnos.frx":044C
         TabIndex        =   10
         Text            =   "V"
         Top             =   600
         Width           =   495
      End
      Begin VB.Frame Frame_Base_Datos 
         Caption         =   "Frame_Base_Datos"
         Height          =   975
         Left            =   3000
         TabIndex        =   8
         Top             =   3000
         Visible         =   0   'False
         Width           =   2775
         Begin VB.Data Data2 
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
            Top             =   240
            Visible         =   0   'False
            Width           =   2505
         End
         Begin VB.Data Data3 
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
            Top             =   600
            Visible         =   0   'False
            Width           =   2505
         End
      End
      Begin VB.TextBox txtTurno 
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
         Left            =   7680
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtEspecialidad 
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
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1080
         Width           =   4695
      End
      Begin VB.TextBox txtCedula 
         DataField       =   "Cedula"
         DataSource      =   "mdbAlumnos"
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
         MaxLength       =   15
         TabIndex        =   11
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtSeccion 
         DataField       =   "Seccion"
         DataSource      =   "mdbAlumnos"
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
         Left            =   4800
         MaxLength       =   5
         TabIndex        =   12
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox txtPeriodo 
         DataField       =   "Periodo"
         DataSource      =   "mdbAlumnos"
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
         MaxLength       =   10
         TabIndex        =   14
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtApellidos 
         DataField       =   "Apellidos"
         DataSource      =   "mdbAlumnos"
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
         MaxLength       =   50
         TabIndex        =   22
         Top             =   2040
         Width           =   5895
      End
      Begin VB.TextBox txtTelefono 
         DataField       =   "Telefono"
         DataSource      =   "mdbAlumnos"
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
         MaxLength       =   50
         TabIndex        =   24
         Top             =   2520
         Width           =   4695
      End
      Begin VB.TextBox txtNombres 
         DataField       =   "Nombres"
         DataSource      =   "mdbAlumnos"
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
         MaxLength       =   50
         TabIndex        =   20
         Top             =   1560
         Width           =   5895
      End
      Begin VB.TextBox txtNumero_ID 
         DataField       =   "Numero_ID"
         DataSource      =   "mdbAlumnos"
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
         Left            =   1320
         Locked          =   -1  'True
         MaxLength       =   9
         TabIndex        =   9
         Top             =   140
         Width           =   1695
      End
      Begin VB.TextBox txtDireccion 
         DataField       =   "Direccion"
         DataSource      =   "mdbAlumnos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1020
         Left            =   1920
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   3000
         Width           =   7215
      End
      Begin VB.Label Label6 
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
         Left            =   6720
         TabIndex        =   29
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label5 
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
         Left            =   120
         TabIndex        =   28
         Top             =   1080
         Width           =   1695
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
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
         Height          =   300
         Index           =   4
         Left            =   120
         TabIndex        =   27
         Top             =   2520
         Width           =   1215
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Apellidos :"
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
         TabIndex        =   25
         Top             =   2040
         Width           =   1260
      End
      Begin VB.Label lblLabels 
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
         Index           =   2
         Left            =   3600
         TabIndex        =   23
         Top             =   615
         Width           =   1125
      End
      Begin VB.Label lblLabels 
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
         Index           =   1
         Left            =   6120
         TabIndex        =   21
         Top             =   615
         Width           =   1080
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
         TabIndex        =   19
         Top             =   615
         Width           =   1005
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
         Caption         =   "Nombres :"
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
         TabIndex        =   17
         Top             =   1560
         Width           =   1230
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
         Index           =   6
         Left            =   120
         TabIndex        =   15
         Top             =   195
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         AutoSize        =   -1  'True
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
         Height          =   300
         Index           =   7
         Left            =   120
         TabIndex        =   13
         Top             =   3000
         Width           =   1290
      End
   End
   Begin MSAdodcLib.Adodc mdbAlumnos 
      Align           =   2  'Align Bottom
      Height          =   450
      Left            =   0
      Top             =   5340
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   794
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
      Connect         =   $"frmAlumnos.frx":0456
      OLEDBString     =   $"frmAlumnos.frx":04E2
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select Numero_ID,Nacionalidad,Cedula,Periodo,Seccion,Nombres,Apellidos,Telefono,Direccion from Alumnos Order by Numero_ID"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   1080
      TabIndex        =   6
      Top             =   4200
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
         Picture         =   "frmAlumnos.frx":056E
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
         Picture         =   "frmAlumnos.frx":0995
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
         Picture         =   "frmAlumnos.frx":0DB0
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
         Picture         =   "frmAlumnos.frx":11D8
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
         Picture         =   "frmAlumnos.frx":1615
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
         Picture         =   "frmAlumnos.frx":1A7E
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   150
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmAlumnos"
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
    Dim Buscar, Cedula As String
    Cedula = InputBox("Introduzca La Cedula :", "Busqueda de Datos")
    If Cedula <> "" Then
        Buscar = "[Cedula]" & "=" & "'" & Cedula & "'"
        mdbAlumnos.Recordset.Find Buscar
        If Not mdbAlumnos.Recordset.EOF Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    Frame1.Enabled = True
    mdbAlumnos.Recordset.AddNew
    txtCedula = Cedula
    txtPeriodo = Data3.Recordset.Fields("Periodo")
    txtNumero_ID = "."
    '--------- Botones ------------
    cmdUpdate.Caption = "&Actualizar"
    cmdUpdate.Enabled = True
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdBuscar.Enabled = False
    cmdImprimir.Enabled = False
    cmdClose.Caption = "&Cancelar"
    mdbAlumnos.Enabled = False
    '------------------------------'
    txtNacionalidad.SetFocus
End Sub

Private Sub cmdDelete_Click()
    Dim Respuesta
    If Not mdbAlumnos.Recordset.EOF Then
        'esto puede producir un error si elimina el último
        'registro o el único registro del recordset
        Respuesta = MsgBox("Desea Eliminar el Registro", vbYesNo + vbCritical, "Eliminación de Datos")
        If Respuesta = vbYes Then
            mdbAlumnos.Recordset.Delete
            mdbAlumnos.Recordset.MoveFirst
        End If
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim Buscar, Campo_Sel As String
    If (mdbAlumnos.Recordset.AbsolutePosition + 1) = 0 Then
        Exit Sub
    End If
    '************************************************
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
        If Campo_Sel = "Numero_ID" Then
            Buscar = "[" & Campo_Sel & "]" & "=" & Buscar
        Else
            Buscar = "[" & Campo_Sel & "]" & "=" & "'" & Buscar & "'"
        End If
        mdbAlumnos.Recordset.MoveFirst
        mdbAlumnos.Recordset.Find Buscar
        If mdbAlumnos.Recordset.EOF Then
            MsgBox Campo_Sel & " del Alumno No Existe", vbCritical, "Buscar " & Campo_Sel
        End If
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim fReportes As New frmReportes
    fReportes.Caption = "Reportes de Alumnos"
    fReportes.Tipo_Reporte = "ALU"
    Call centrarform(fReportes)
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
        If txtNumero_ID = "." Then
            txtNumero_ID = Data3.Recordset.Fields("Numero_ID")
            Numero_ID = Val(txtNumero_ID) + 1
            Data3.Recordset.Edit
            Data3.Recordset.Fields("Numero_ID") = Numero_ID
            Data3.Recordset.Update
        End If
        Frame1.Enabled = False
        mdbAlumnos.Recordset.Update '.UpdateRecord
'        mdbAlumnos.Recordset.Bookmark = mdbAlumnos.Recordset.LastModified
        '--------- Botones ------------
        cmdUpdate.Caption = "&Modificar"
        cmdAdd.Enabled = True
        cmdDelete.Enabled = True
        cmdBuscar.Enabled = True
        cmdImprimir.Enabled = True
        cmdClose.Caption = "&Cerrar"
        mdbAlumnos.Enabled = True
        '------------------------------'
    Else
        If Not mdbAlumnos.Recordset.EOF Then
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
            Frame1.Enabled = True
            'mdbAlumnos.Recordset.Edit
            '--------- Botones ------------
            cmdUpdate.Caption = "&Actualizar"
            cmdAdd.Enabled = False
            cmdDelete.Enabled = False
            cmdBuscar.Enabled = False
            cmdImprimir.Enabled = False
            cmdClose.Caption = "&Cancelar"
            mdbAlumnos.Enabled = False
            '------------------------------'
            txtNombres.SetFocus
        End If
    End If
End Sub

Private Sub cmdClose_Click()
    If cmdClose.Caption = "&Cerrar" Then
        Unload Me
    Else
        Frame1.Enabled = False
        '--------- Botones ------------
        cmdUpdate.Caption = "&Modificar"
        cmdAdd.Enabled = True
        cmdDelete.Enabled = True
        cmdBuscar.Enabled = True
        cmdImprimir.Enabled = True
        cmdClose.Caption = "&Cerrar"
        mdbAlumnos.Enabled = True
        '------------------------------'
        mdbAlumnos.Refresh
        mdbAlumnos.Recordset.MoveFirst
    End If
End Sub

Private Sub mdbAlumnos_Reposition()
  On Error Resume Next
  'Esto mostrará la posición del registro actual
  'para dynasets y snapshots
  mdbAlumnos.Caption = "Record: " & (mdbAlumnos.Recordset.AbsolutePosition + 1)
  Buscar_Datos
End Sub

Sub Buscar_Datos()
    Dim Buscar As String
    Buscar = "[Codigo]" & "=" & "'" & Mid(txtSeccion, 1, 2) & "'"
    If Mid(txtSeccion, 1, 2) = "" Then Exit Sub
    Data2.Recordset.FindFirst Buscar
    txtEspecialidad = Data2.Recordset.Fields("Descripcion")
    Select Case Mid(txtSeccion, 5, 1)
    Case "1"
        txtTurno = "Mañana"
    Case "3"
        txtTurno = "Noche"
    End Select
End Sub

Private Sub Form_Activate()
    mdbAlumnos.ConnectionString = DSN_Pasantias
    mdbAlumnos.RecordSource = "Alumnos"
    mdbAlumnos.Refresh
    If (mdbAlumnos.Recordset.AbsolutePosition + 1) <> 0 Then
        Call Buscar_Datos
    Else
        mdbAlumnos.Recordset.MoveFirst
        Call Buscar_Datos
    End If
End Sub

Private Sub Form_Load()
    If Me.WindowState <> 2 Then
        Me.Move (Screen.Width - Me.Width) / 2, 0
    End If
End Sub

Private Sub txtApellidos_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtTelefono.SetFocus
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtSeccion.SetFocus
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cmdUpdate.SetFocus
End Sub

Private Sub txtNacionalidad_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtCedula.SetFocus
End Sub

Private Sub txtNombres_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtApellidos.SetFocus
End Sub

Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtNombres.SetFocus
End Sub

Private Sub txtSeccion_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        Call Buscar_Datos
        txtPeriodo.SetFocus
    End If
End Sub

Private Sub mdbAlumnos_WillMove(ByVal adReason As ADODB.EventReasonEnum, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Call Buscar_Datos
End Sub

Private Sub mdbAlumnos_DragOver(Source As Control, X As Single, Y As Single, State As Integer)
    Call Buscar_Datos
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then txtDireccion.SetFocus
End Sub
