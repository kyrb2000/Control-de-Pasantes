VERSION 5.00
Begin VB.Form frmGenerarReport 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3465
   ClientLeft      =   2130
   ClientTop       =   2205
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame_Base_Datos 
      Caption         =   "Frame_Base_Datos"
      Height          =   1095
      Left            =   480
      TabIndex        =   10
      Top             =   1680
      Visible         =   0   'False
      Width           =   2775
      Begin VB.Data DEspecialidad 
         Caption         =   "Especialidad"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   375
         Left            =   120
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Especialidad"
         Top             =   600
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.Data DTabla_Gerenar 
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
   End
   Begin VB.ComboBox cobEspecialidad 
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
      ItemData        =   "frmGenerarReport.frx":0000
      Left            =   1680
      List            =   "frmGenerarReport.frx":0007
      TabIndex        =   6
      Text            =   "Todos"
      Top             =   2040
      Width           =   2775
   End
   Begin VB.ComboBox cobSeccion 
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
      ItemData        =   "frmGenerarReport.frx":0012
      Left            =   1680
      List            =   "frmGenerarReport.frx":0019
      TabIndex        =   5
      Text            =   "Todos"
      Top             =   1560
      Width           =   1935
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
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   2
      Top             =   1080
      Width           =   1935
   End
   Begin VB.Frame Frame2 
      Height          =   855
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Width           =   4215
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   615
         Left            =   360
         Picture         =   "frmGenerarReport.frx":0024
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   170
         Width           =   1695
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   615
         Left            =   2160
         Picture         =   "frmGenerarReport.frx":0366
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   170
         Width           =   1695
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   3495
      Left            =   0
      Top             =   0
      Width           =   4695
   End
   Begin VB.Label lblEspecialidad 
      AutoSize        =   -1  'True
      Caption         =   "Especialidad :"
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
      TabIndex        =   4
      Top             =   2040
      Width           =   1395
   End
   Begin VB.Label lblSeccion 
      AutoSize        =   -1  'True
      Caption         =   "Sección :"
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
      TabIndex        =   3
      Top             =   1560
      Width           =   900
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
      TabIndex        =   1
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Selecciones los Item para Generar el Reporte"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmGenerarReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
'-- Autor : Kenner Y. Roa B. Telefono (0212) 433.55.14
'-- E-mail : kyrb2000@cantv.net / kyrb2000@hotmail.com
'-- Fecha : 22/05/2002 (Caracas / Venezuela)
'-----------------------------------------------------------
Public Boton_Respuesta As Boolean
Public pPeriodo As String
Public pSeccion As String
Public pEspecialidad As String

Private Sub cmdAceptar_Click()
    Dim Buscar As String
    pPeriodo = IIf(Len(txtPeriodo) = 0, "*", txtPeriodo)
    If cobSeccion <> "Todos" Then
        Buscar = "[Codigo]" & "=" & "'" & Mid(cobSeccion.List(cobSeccion.ListIndex), 1, 2) & "'"
        DEspecialidad.Recordset.FindFirst Buscar
        cobEspecialidad.Text = DEspecialidad.Recordset.Fields("Descripcion")
        pSeccion = cobSeccion
    Else
        pSeccion = "*"
    End If
    pEspecialidad = IIf(cobEspecialidad = "Todos", "*", cobEspecialidad)
    Boton_Respuesta = True
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    Boton_Respuesta = False
    Unload Me
End Sub

Private Sub cobSeccion_Click()
    Dim Buscar As String
    Buscar = "[Codigo]" & "=" & "'" & Mid(cobSeccion.List(cobSeccion.ListIndex), 1, 2) & "'"
    DEspecialidad.Recordset.FindFirst Buscar
    cobEspecialidad.Text = DEspecialidad.Recordset.Fields("Descripcion")
End Sub

Private Sub cobSeccion_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cobEspecialidad_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Form_Activate()
    DEspecialidad.DatabaseName = Base_de_Datos
    DEspecialidad.Refresh
    DTabla_Gerenar.DatabaseName = Base_de_Datos
    DTabla_Gerenar.Refresh
    Call Cargar_Datos
End Sub

Sub Cargar_Datos()
    Dim SQLcad As String
    Dim dynBaseD As Dynaset
    BaseD = Base_de_Datos
    BD = Abrir_BaseDatos(BaseD, 1)
    SQLcad = "SELECT DISTINCT Seccion FROM Alumnos"
    Set dynBaseD = MAESTRO.CreateDynaset(SQLcad)
        If Not dynBaseD.EOF Then
            dynBaseD.MoveFirst
            Do Until dynBaseD.EOF
                If Not dynBaseD.EOF Then
                    cobSeccion.AddItem dynBaseD.Fields("Seccion")
                End If
                dynBaseD.MoveNext
            Loop
        End If
    dynBaseD.Close
    If Not DEspecialidad.Recordset.EOF Then
        DEspecialidad.Recordset.MoveFirst
        Do Until DEspecialidad.Recordset.EOF
            If Not DEspecialidad.Recordset.EOF Then
                cobEspecialidad.AddItem DEspecialidad.Recordset.Fields("Descripcion")
            End If
            DEspecialidad.Recordset.MoveNext
        Loop
    End If
End Sub

Private Sub txtPeriodo_Click()
    txtPeriodo = DTabla_Gerenar.Recordset.Fields("Periodo")
End Sub

Private Sub txtPeriodo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then cobSeccion.SetFocus
End Sub
