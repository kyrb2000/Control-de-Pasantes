VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCargar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carga de Datos"
   ClientHeight    =   3690
   ClientLeft      =   1995
   ClientTop       =   1980
   ClientWidth     =   4575
   Icon            =   "frmCargar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   60
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.Frame Frame_Base_Datos 
         Caption         =   "Frame_Base_Datos"
         Height          =   1095
         Left            =   480
         TabIndex        =   12
         Top             =   240
         Visible         =   0   'False
         Width           =   2295
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
         Height          =   360
         Left            =   2160
         MaxLength       =   10
         TabIndex        =   10
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtCedula 
         DataField       =   "Cedula"
         DataSource      =   "Adodc1"
         Height          =   285
         Left            =   3120
         TabIndex        =   7
         Text            =   "txtCedula"
         Top             =   1080
         Visible         =   0   'False
         Width           =   1215
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   480
         Top             =   720
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame Frame2 
         Height          =   1575
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   4215
         Begin MSComctlLib.ProgressBar ProgressBar2 
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Visible         =   0   'False
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   1
            Scrolling       =   1
         End
         Begin MSComctlLib.ProgressBar ProgressBar1 
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Visible         =   0   'False
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "&Cerrar"
            Height          =   615
            Left            =   2160
            Picture         =   "frmCargar.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "&Aceptar"
            Enabled         =   0   'False
            Height          =   615
            Left            =   360
            Picture         =   "frmCargar.frx":064C
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   840
            Width           =   1695
         End
         Begin VB.Label lblText2 
            Alignment       =   2  'Center
            Caption         =   "luego pulse [Aceptar]  ó  [Salir]"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   480
            Width           =   3975
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblText1 
            Alignment       =   2  'Center
            Caption         =   "Introduzca el disquette en la unidad A:"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   3975
            WordWrap        =   -1  'True
         End
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
         Left            =   960
         TabIndex        =   11
         Top             =   1440
         Width           =   900
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "CARGA DE DATOS"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Left            =   960
         TabIndex        =   1
         Top             =   120
         Width           =   2970
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   120
         Picture         =   "frmCargar.frx":098E
         Top             =   240
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmCargar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
'-- Autor : Kenner Y. Roa B. Telefono (0212) 433.55.14
'-- E-mail : kyrb2000@cantv.net / kyrb2000@hotmail.com
'-- Fecha : 22/05/2002 (Caracas / Venezuela)
'-----------------------------------------------------------

Private Sub cmdAceptar_Click()
    Dim Dbf_Archivo As String
    Dim Ruta_Directorio As String
    Dim Nombre_Archivo As String
    
    If cmdCancelar.Caption = "&Cerrar" Then
        cmdCancelar.Caption = "&Cancelar"
        ' Establecer CancelError a True
        CommonDialog1.CancelError = True
        On Error GoTo Error_Dialogo
        ' Establecer los indicadores
        CommonDialog1.Flags = cdlOFNHideReadOnly
        ' Establecer los filtros
        CommonDialog1.Filter = "Todos los archivos (*.dbf,*.mdb)|*.dbf;*.mdb"
        ' Especificar el filtro predeterminado
        CommonDialog1.FilterIndex = 2
        ' Presentar el cuadro de diálogo Abrir
        CommonDialog1.ShowOpen
        ' Presentar el nombre del archivo seleccionado
        Dbf_Archivo = CommonDialog1.FileName
        Ruta_Directorio = StripFileName(CommonDialog1.FileName)
        Nombre_Archivo = StripFile(CommonDialog1.FileTitle, ".")
        On Error GoTo 0
        Call Cargar_Datos(Ruta_Directorio, Nombre_Archivo)
    End If
    cmdCancelar.Caption = "&Cerrar"
    Call cmdCancelar_Click
    Exit Sub
    
Error_Dialogo:
    ' El usuario ha hecho clic en el botón Cancelar
    cmdCancelar.Caption = "&Salir"
    Exit Sub
End Sub

Sub Cargar_Datos(Ruta_Base_de_Datos As String, Base_de_Datos As String)
    Dim i1, i2 As Integer
    Dim Existe As Boolean
    
    ProgressBar1.Visible = True
    ProgressBar2.Visible = True
    Set MAESTRO = OpenDatabase(Ruta_Base_de_Datos, False, 0, "Dbase IV;")
    Set Ds1 = MAESTRO.OpenRecordset(Base_de_Datos, dbOpenDynaset)
    Ds1.MoveLast
    ProgressBar1.Max = Ds1.RecordCount
    ProgressBar2.Max = 100 'Data1.Recordset.RecordCount
    i1 = 0
    i2 = 0
    Ds1.MoveFirst
    While Not Ds1.EOF
        ProgressBar1.Value = i1
        i1 = i1 + 1
        If ProgressBar1.Value = ProgressBar1.Max Then
            i1 = 0
        End If
        Existe = False
        Buscar = "[Cedula]" & "=" & "'" & Ds1.Fields("Cedula") & "'"
        Data1.Recordset.FindFirst Buscar
        If Not Data1.Recordset.EOF Then
            If Ds1.Fields("Cedula") = Data1.Recordset.Fields("Cedula") Then
                Existe = True
            End If
        End If
            ProgressBar2.Value = i2
            i2 = i2 + 1
            If ProgressBar2.Value = ProgressBar2.Max Then
                i2 = 0
            End If
        If Not Existe Then
           ' MsgBox Ds1.Fields("Cedula"), , "No Existe"
            Nombres = FNombre(Ds1.Fields("Ape_Nom"))
            Apellidos = FApellidos(Ds1.Fields("Ape_Nom"))
            Numero_ID = Data2.Recordset.Fields("Numero_ID")
            Data1.Recordset.AddNew
            Data1.Recordset.Fields("Numero_ID") = Numero_ID
            Data1.Recordset.Fields("Nacionalidad") = Ds1.Fields("Nacio")
            Data1.Recordset.Fields("Cedula") = Ds1.Fields("Cedula")
            Data1.Recordset.Fields("Seccion") = Ds1.Fields("Turno") 'Ds1.Fields("Seccion")
            Data1.Recordset.Fields("Nombres") = Nombres
            Data1.Recordset.Fields("Apellidos") = Apellidos
            Data1.Recordset.Fields("Periodo") = txtPeriodo
            Data1.Recordset.Fields("Telefono") = Ds1.Fields("TelefHab")
            'Data1.Recordset.Fields("Telefono") = "-"
            Data1.Recordset.Fields("Direccion") = Ds1.Fields("DirHab")
            Data1.Recordset.Update
            Numero_ID = Numero_ID + 1
            Data2.Recordset.Edit
            Data2.Recordset.Fields("Numero_ID") = Numero_ID
            Data2.Recordset.Update
        End If
        Ds1.MoveNext
    Wend
    For c1 = 0 To ProgressBar1.Max
        ProgressBar1.Value = c1
    Next
    For c1 = 0 To ProgressBar2.Max
        ProgressBar2.Value = c1
    Next
    Ds1.Close
    MAESTRO.Close
    ProgressBar1.Visible = False
    ProgressBar2.Visible = False
End Sub

Private Sub cmdCancelar_Click()
    If cmdCancelar.Caption = "&Cerrar" Then
        Unload Me
    Else
        cmdCancelar.Caption = "&Cerrar"
    End If
End Sub

Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
    If Len(txtPeriodo) <> 0 Then
        cmdAceptar.Enabled = True
    Else
        cmdAceptar.Enabled = False
    End If
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

Private Sub txtPeriodo_LostFocus()
    If Len(txtPeriodo) <> 0 Then
        cmdAceptar.Enabled = True
    Else
        cmdAceptar.Enabled = False
    End If
End Sub
