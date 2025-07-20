VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmBuscar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selector de Queries"
   ClientHeight    =   5535
   ClientLeft      =   2310
   ClientTop       =   2115
   ClientWidth     =   8250
   Icon            =   "frmBuscar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5535
   ScaleWidth      =   8250
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraWhere 
      Caption         =   "Cláusula WHERE"
      Height          =   1035
      Left            =   2940
      TabIndex        =   5
      Top             =   1350
      Width           =   5175
      Begin VB.OptionButton optOr 
         Caption         =   "Or"
         Height          =   195
         Left            =   2610
         TabIndex        =   20
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton optAnd 
         Caption         =   "And"
         Height          =   195
         Left            =   1920
         TabIndex        =   19
         Top             =   720
         Width           =   645
      End
      Begin VB.CommandButton cmdBorrarWhere 
         Caption         =   "&Borrar Where"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   690
         Width           =   1275
      End
      Begin VB.ComboBox cmbWhereCampo 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   300
         Width           =   1905
      End
      Begin VB.ComboBox cmbWhereOperador 
         Height          =   315
         Left            =   2100
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   300
         Width           =   1005
      End
      Begin VB.TextBox txtCondicion 
         Height          =   315
         Left            =   3180
         TabIndex        =   7
         Text            =   "txtCondicion"
         Top             =   300
         Width           =   1875
      End
      Begin VB.CommandButton cmdWhere 
         Caption         =   "Asignar &Where"
         Height          =   255
         Left            =   3570
         TabIndex        =   6
         Top             =   690
         Width           =   1485
      End
   End
   Begin VB.Frame fraTabla 
      Caption         =   "Selección de tabla"
      Height          =   675
      Left            =   2940
      TabIndex        =   13
      Top             =   660
      Width           =   2745
      Begin VB.ComboBox cmbTabla 
         Height          =   315
         Left            =   150
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   270
         Width           =   2415
      End
   End
   Begin VB.CommandButton cmdEjecutarSelect 
      Caption         =   "&Ejecutar Select"
      Height          =   345
      Left            =   6240
      TabIndex        =   22
      Top             =   2400
      Width           =   1845
   End
   Begin VB.CommandButton cmsVisualizarSelect 
      Caption         =   "&Visualizar Select"
      Height          =   345
      Left            =   4320
      TabIndex        =   23
      Top             =   2400
      Width           =   1845
   End
   Begin VB.Frame fraSelect 
      Caption         =   "Selección (SQL Select)"
      Height          =   2145
      Left            =   90
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   8025
      Begin VB.Label lblSelect 
         AutoSize        =   -1  'True
         Caption         =   "Select * "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   150
         TabIndex        =   4
         Top             =   360
         Width           =   1035
      End
      Begin VB.Label lblFrom 
         AutoSize        =   -1  'True
         Caption         =   "From "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   510
         TabIndex        =   3
         Top             =   690
         Width           =   750
      End
      Begin VB.Label lblWhere 
         AutoSize        =   -1  'True
         Caption         =   "Where "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   510
         TabIndex        =   2
         Top             =   1050
         Width           =   7275
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame fraSeleccion 
      Caption         =   "Selección"
      Height          =   2655
      Left            =   90
      TabIndex        =   0
      Top             =   2760
      Width           =   8025
      Begin MSFlexGridLib.MSFlexGrid grdSeleccion 
         Height          =   2265
         Left            =   150
         TabIndex        =   21
         Top             =   270
         Width           =   7725
         _ExtentX        =   13626
         _ExtentY        =   3995
         _Version        =   393216
      End
   End
   Begin VB.Frame fraCampos 
      Caption         =   "Selección de campos"
      Height          =   1725
      Left            =   90
      TabIndex        =   10
      Top             =   660
      Width           =   2745
      Begin VB.ListBox lstCampos 
         Height          =   1035
         Left            =   150
         MultiSelect     =   1  'Simple
         TabIndex        =   12
         Top             =   270
         Width           =   2415
      End
      Begin VB.CommandButton cmdCampos 
         Caption         =   "&Asignar campos"
         Height          =   255
         Left            =   150
         TabIndex        =   11
         Top             =   1350
         Width           =   2415
      End
   End
   Begin VB.Frame fraBD 
      Caption         =   "Selección de la base de datos "
      Height          =   705
      Left            =   90
      TabIndex        =   15
      Top             =   30
      Width           =   7995
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   600
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox txtBD 
         Height          =   285
         Left            =   150
         TabIndex        =   17
         Text            =   "txtBD"
         Top             =   270
         Width           =   6075
      End
      Begin VB.CommandButton cmdBD 
         Caption         =   "Base de &Datos"
         Height          =   285
         Left            =   6390
         TabIndex        =   16
         Top             =   270
         Width           =   1485
      End
   End
End
Attribute VB_Name = "frmBuscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim BD As Database
Dim tdf As TableDefs
Dim fld As Fields
Dim NumTablas As Integer
Dim Where As String
'----------------------------------------------------------
'-- Abre la base de datos cuyo nombre es "DBName" (pasado
'-- como parámetro
'----------------------------------------------------------
Function openDB(DBName As String) As Integer
    If Trim(DBName) = "" Then GoTo OpenDBErr
  Dim MiNombreUsuario As String
  Dim MiNombreGrupo As String
  Dim criterio As String
  On Error GoTo OpenDBErr
  MiNombreUsuario = ""
  MiNombreGrupo = ""
  criterio = "ODBC;UID=" & MiNombreUsuario & ";PWS=" & MiNombreGrupo
  Set BD = OpenDatabase(DBName, , , criterio)
  'exito
  openDB = True
  GoTo OpenDBEnd
OpenDBErr:
  openDB = False
  Resume OpenDBEnd
OpenDBEnd:
End Function

'--------------------------------------------------------------------
'Esta función se encarga de centrar el formulario enviado como
'parámetro.
'--------------------------------------------------------------------
Sub centrarform(f As Form)
    f.Move (Screen.Width - f.Width) / 2, (Screen.Height - f.Height) / 2
End Sub





Private Sub CargarCampos(BaseDatos As Database, Tabla As String, lst As Control)
    Dim i As Integer
    Set fld = tdf(Tabla).Fields
    lst.Clear
    For i = 0 To fld.Count - 1
        lst.AddItem fld(i).Name
    Next i
End Sub



Private Sub cmbTabla_Click()
    lblFrom = "From [" & cmbTabla.Text & "]"
    lblSelect = "Select * "
    lblWhere.Visible = False
    lblWhere = "Where "
    cmdCampos.Enabled = True
    cmdEjecutarSelect.Enabled = True
    Call CargarCampos(BD, cmbTabla.Text, lstCampos)
    Call CargarCampos(BD, cmbTabla.Text, cmbWhereCampo)
End Sub

Private Sub cmbWhereCampo_Click()
    On Error GoTo ErrWhere
    txtCondicion = ""
    cmbWhereOperador.Clear
    Select Case fld(cmbWhereCampo.Text).Type
        Case 1: 'dbBoolean Yes / No
            cmbWhereOperador.AddItem "="
            cmbWhereOperador.AddItem "<>"
        Case 2 To 8: 'Númericos y Fecha/Hora
            cmbWhereOperador.AddItem "="
            cmbWhereOperador.AddItem "<>"
            cmbWhereOperador.AddItem ">"
            cmbWhereOperador.AddItem ">="
            cmbWhereOperador.AddItem "<"
            cmbWhereOperador.AddItem "<="
        Case 10, 12: 'Texto, Memo
            cmbWhereOperador.AddItem "Like"
        Case 11: 'Long Binary (Objeto OLE)
    End Select
    Exit Sub
ErrWhere:
    MsgBox "Error en la creación de la cláusula Where", , "Visual Basic 6.0. Práctica 35"
End Sub

Private Sub cmdBD_Click()
    On Error GoTo ErrBD
    cmbTabla.Clear
    lstCampos.Clear
    cmbWhereCampo.Clear
    cmbWhereOperador.Clear
    txtCondicion = ""
    cmdCampos.Enabled = False
    cmdEjecutarSelect.Enabled = False
    fraSeleccion.Visible = False
    If openDB(txtBD) Then
        Call CargarTablas(BD, cmbTabla)
        GoTo FinBD
     Else
        GoTo ErrBD
    End If
ErrBD:
    MsgBox "Error al abrir la base de datos. Verifique la ruta y nombre del archivo.", , Me.Caption
FinBD:
End Sub

Private Sub CargarTablas(BaseDatos As Database, lst As Control)
    Dim i As Integer
    Set tdf = BaseDatos.TableDefs
    lst.Clear
    NumTablas = tdf.Count - 1
    For i = 0 To NumTablas
        If tdf(i).Attributes = 0 Then
            lst.AddItem tdf(i).Name
        End If
    Next i
End Sub

Private Sub cmdBorrarWhere_Click()
    lblWhere = "Where "
    lblWhere.Visible = False
End Sub

Private Sub cmdCampos_Click()
    Dim i As Integer, X As Integer
    lblSelect = "Select "
    For i = 0 To lstCampos.ListCount - 1
        If lstCampos.Selected(i) Then
            X = X + 1
            lblSelect = lblSelect & "[" & lstCampos.List(i) & "], "
        End If
    Next i
    If X = 0 Then
        lblSelect = "Select * "
     Else
        lblSelect = Left$(lblSelect, Len(lblSelect) - 2)
    End If
    cmdWhere.Enabled = True
    cmdWhere.Enabled = True
    fraSeleccion.Visible = False
End Sub

Private Sub cmdEjecutarSelect_Click()
    On Error GoTo ErrEjecutarSelect
    Dim MyRecordSet As Recordset
    Dim SQLcad As String
    Dim X As Integer
    SQLcad = lblSelect & " " & lblFrom
    If lblWhere.Visible Then SQLcad = SQLcad & " " & lblWhere
    Set MyRecordSet = BD.OpenRecordset(SQLcad, dbOpenSnapshot)
    X = CargarGrid(Me, grdSeleccion, MyRecordSet)
    If X <> 0 Then
        fraSeleccion.Visible = True
    End If
    Exit Sub
ErrEjecutarSelect:
    MsgBox "Error en la construcción del Query.", , Me.Caption
End Sub

Function CargarGrid(f As Form, grd As Control, fds As Recordset) As Integer
   Dim tc As Integer               'Tipo de campo
   Dim i As Integer, j As Integer
   Dim fn As String                'Nombre de campo
   Dim rc As Integer               'Nº registros
   Dim gs As String                'Texto grid

   On Error GoTo LGErr
   
   'Configura el grid
   grd.Rows = 2       '
   grd.FixedRows = 1  'permite el siguiente paso
   grd.Rows = 1       'Limpia el grid completamente
   grd.Cols = fds.Fields.Count + 1
   grd.FixedCols = 1

     On Error GoTo LGErr
     'Carga los nombres de los campos
     grd.Row = 0
     For i = 0 To fds.Fields.Count - 1
       grd.Col = i + 1
       grd.Text = fds(i).Name
       'grd.Text = UCase(fds(i).Name)
       If grd.ColWidth(i + 1) < Len(fds(i).Name) * 120 Then
          grd.ColWidth(i + 1) = Len(fds(i).Name) * 120
       End If
     Next

   rc = 1
   ' Añade las filas del grid
   If fds.RecordCount <> 0 Then
       fds.MoveFirst
       While Not fds.EOF
         gs = CStr(rc + 0) + Chr$(9)
         For i = 0 To fds.Fields.Count - 1
             gs = gs + Format(fds(i), "############,###") + Chr$(9)
         Next
         gs = Mid(gs, 1, Len(gs) - 1)
         grd.AddItem gs
         fds.MoveNext
         rc = rc + 1
       Wend
    
       grd.FixedRows = 1   'Congela la fila de nombre de campos
       grd.FixedCols = 1   'Congela Segmento
       grd.Row = 1         'Fija la posicion inicial
       grd.Col = 1
   End If

   CargarGrid = rc       'Devuelve el nº de filas del grid.
   GoTo LGEnd

LGErr:
   CargarGrid = False    'Devuelve 0
   Resume LGEnd

LGEnd:
   
End Function

Private Sub cmdWhere_Click()
    Dim sqlWhere
    Dim Operador As String
    On Error GoTo ErrWhere
    If optOr = True Then
        Operador = " Or "
     Else
        Operador = " And "
    End If
    fraSeleccion.Visible = False
    If Trim(cmbWhereCampo.Text) = "" Or Trim(cmbWhereOperador.Text) = "" Or Trim(txtCondicion) = "" Then
        MsgBox "No es posible crear la cláusula Where", , Me.Caption
     Else
        If Trim(cmbWhereOperador.Text) = "Like" Then
            sqlWhere = "([" & Trim(cmbWhereCampo.Text) & "] " & _
                Trim(cmbWhereOperador.Text) & " '" & Trim(txtCondicion) & "')"
         Else
            sqlWhere = "([" & Trim(cmbWhereCampo.Text) & "] " & _
                Trim(cmbWhereOperador.Text) & " " & Trim(txtCondicion) & ")"
        End If
        If Right(lblWhere, 1) = ")" Then
            lblWhere = lblWhere & Operador & sqlWhere
         Else
            lblWhere = lblWhere & sqlWhere
        End If
        lblWhere.Visible = True
        cmdBorrarWhere.Enabled = True
    End If
    Exit Sub
ErrWhere:
    MsgBox "No es posible crear la cláusula Where", , Me.Caption
End Sub

Private Sub cmsVisualizarSelect_Click()
    If fraSelect.Visible = True Then
        fraSelect.Visible = False
    Else
        fraSelect.Visible = True
    End If
End Sub

Private Sub Form_Load()
'C:\Archivos de programa\DevStudio\VB\biblio.mdb
'Vaciamos todas las listas y preparamos el entorno
    lstCampos.Clear
    cmbTabla.Clear
    cmbWhereCampo.Clear
    cmbWhereOperador.Clear
    txtCondicion.Text = ""
    txtBD.Text = App.Path & "\Datos_Alumnos.mdb"
    fraSeleccion.Visible = False
    cmdEjecutarSelect.Enabled = False
    lblWhere.Visible = False
    cmdWhere.Enabled = False
    cmdBorrarWhere.Enabled = False
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'    Dim Respuesta As Integer
'    Respuesta = MsgBox("¿Desea salir de la aplicación?", 4, "Visual Basic 6.0. Práctica 35")
'    If Respuesta = 6 Then
'        End
'     Else
'        Cancel = True
'    End If
    Unload Me
End Sub

Private Sub fraSelect_Click()
    fraSelect.Visible = False
End Sub
Private Sub txtBD_Click()
    Dim Dbf_Archivo As String
    Dim Ruta_Directorio As String
    Dim Nombre_Archivo As String
    
        ' Establecer CancelError a True
        CommonDialog1.CancelError = True
        'On Error GoTo Error_Dialogo
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
        txtBD.Text = Dbf_Archivo
        On Error GoTo 0
        Exit Sub
    
Error_Dialogo:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub
End Sub

