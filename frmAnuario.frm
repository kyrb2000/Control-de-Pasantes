VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAnuario 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Diario de Mensajes"
   ClientHeight    =   3285
   ClientLeft      =   1755
   ClientTop       =   2025
   ClientWidth     =   4725
   Icon            =   "frmAnuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdActual 
      Caption         =   "Ac&tual"
      Height          =   300
      Left            =   3960
      TabIndex        =   12
      Top             =   2200
      Width           =   700
   End
   Begin VB.Data mdbDiario 
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   2640
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Diario"
      Top             =   2180
      Width           =   1260
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   2
      Top             =   2480
      Width           =   4455
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Cerrar"
         Height          =   540
         Left            =   3360
         Picture         =   "frmAnuario.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   150
         Width           =   975
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Modificar"
         Height          =   540
         Left            =   2280
         Picture         =   "frmAnuario.frx":064C
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   150
         Width           =   975
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "&Eliminar"
         Height          =   540
         Left            =   1200
         Picture         =   "frmAnuario.frx":098E
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Width           =   975
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "&Agregar"
         Height          =   540
         Left            =   120
         Picture         =   "frmAnuario.frx":0CD0
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdHoy 
      Caption         =   "&Hoy"
      Height          =   300
      Left            =   2640
      TabIndex        =   1
      Top             =   30
      Width           =   800
   End
   Begin MSComCtl2.MonthView MonthView1 
      DataSource      =   "mdbDiario"
      Height          =   2370
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   24444929
      CurrentDate     =   37652
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   2520
      TabIndex        =   7
      Top             =   -120
      Width           =   2295
      Begin VB.TextBox txtMensaje 
         DataField       =   "Mensaje"
         DataSource      =   "mdbDiario"
         Height          =   1455
         Left            =   120
         MaxLength       =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox cboHora 
         DataField       =   "Hora"
         DataSource      =   "mdbDiario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmAnuario.frx":1012
         Left            =   120
         List            =   "frmAnuario.frx":1014
         TabIndex        =   8
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label lblFecha 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         DataField       =   "Fecha"
         DataSource      =   "mdbDiario"
         Height          =   285
         Left            =   960
         TabIndex        =   10
         Top             =   150
         Width           =   1215
      End
   End
   Begin VB.Label lblfecha_M 
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   720
      Width           =   975
   End
End
Attribute VB_Name = "frmAnuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboHora_KeyPress(KeyAscii As Integer)
    'KeyAscii = 0
End Sub

Private Sub cmdActual_Click()
    MonthView1.Value = lblFecha.Caption
    MonthView1.Refresh
End Sub

Private Sub cmdAdd_Click()
    Frame2.Enabled = True
    mdbDiario.Recordset.AddNew
    lblFecha.Caption = DateClicked
    For h = 1 To 12
        For M = 0 To 59 Step 5
            cboHora.AddItem IIf(h < 10, "0" & Trim(Str(h)), Trim(Str(h))) & ":" & IIf(M < 10, "0" & Trim(Str(M)), Trim(Str(M))) & ":00 a.m."
        Next
    Next
    For h = 1 To 12
        For M = 0 To 59 Step 5
            cboHora.AddItem IIf(h < 10, "0" & Trim(Str(h)), Trim(Str(h))) & ":" & IIf(M < 10, "0" & Trim(Str(M)), Trim(Str(M))) & ":00 p.m."
        Next
    Next
    '--------- Botones ------------
    cmdUpdate.Caption = "&Actualizar"
    cmdUpdate.Enabled = True
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdClose.Caption = "&Cancelar"
    mdbDiario.Enabled = False
    '------------------------------'
    cboHora.SetFocus
End Sub

Private Sub cmdHoy_Click()
    MonthView1.Value = Date
    MonthView1.Refresh
End Sub

Private Sub cmdUpdate_Click()
    On Error GoTo Err_Reparar
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
        Frame2.Enabled = False
        mdbDiario.Recordset.Update '.UpdateRecord
        mdbDiario.Refresh
        '--------- Botones ------------
        cmdUpdate.Caption = "&Modificar"
        cmdAdd.Enabled = True
        cmdDelete.Enabled = True
        cmdClose.Caption = "&Cerrar"
        mdbDiario.Enabled = True
        '------------------------------'
    Else
        If Not mdbDiario.Recordset.EOF Then
            Hora_Mod = cboHora.Text
            For h = 1 To 12
                For M = 0 To 59 Step 5
                    cboHora.AddItem IIf(h < 10, "0" & Trim(Str(h)), Trim(Str(h))) & ":" & IIf(M < 10, "0" & Trim(Str(M)), Trim(Str(M))) & ":00 a.m."
                Next
            Next
            For h = 1 To 12
                For M = 0 To 59 Step 5
                    cboHora.AddItem IIf(h < 10, "0" & Trim(Str(h)), Trim(Str(h))) & ":" & IIf(M < 10, "0" & Trim(Str(M)), Trim(Str(M))) & ":00 p.m."
                Next
            Next
            ' Si el Control Text esta en blanco
            ' Reeplazalo con un "-"
            For i = 0 To Me.Controls.Count - 1
                If TypeOf Me.Controls(i) Is TextBox Then
                    If Me.Controls(i).Text = "" Then
                        Me.Controls(i).Text = "-"
                    End If
                End If
            Next i
            cboHora.Text = Hora_Mod
            ' ----------------------------------
            Frame2.Enabled = True
            mdbDiario.Recordset.Edit
            '--------- Botones ------------
            cmdUpdate.Caption = "&Actualizar"
            cmdAdd.Enabled = False
            cmdDelete.Enabled = False
            cmdClose.Caption = "&Cancelar"
            mdbDiario.Enabled = False
            '------------------------------'
            cboHora.SetFocus
        End If
    End If
    Exit Sub
Err_Reparar:
    mdbDiario.Refresh
End Sub

Private Sub cmdDelete_Click()
    Dim Respuesta
    On Error GoTo Err_Reparar
    If Not mdbDiario.Recordset.EOF Then
        'esto puede producir un error si elimina el último
        'registro o el único registro del recordset
        Respuesta = MsgBox("Desea Eliminar el Registro", vbYesNo + vbCritical, "Eliminación de Datos")
        If Respuesta = vbYes Then
            mdbDiario.Recordset.Delete
            mdbDiario.Recordset.MoveFirst
            mdbDiario.Refresh
        End If
    End If
Exit Sub
Err_Reparar:
    mdbDiario.Refresh
    Respuesta = MsgBox("No esta Eliminado el Registro", , "Eliminación de Datos")
End Sub

Private Sub cmdClose_Click()
    If cmdClose.Caption = "&Cerrar" Then
        If Not mdbDiario.Recordset.EOF Then
            Buscar = "[Fecha]=" & Date
            mdbDiario.Recordset.MoveFirst
            mdbDiario.Recordset.FindFirst Buscar
            If mdbDiario.Recordset.EOF Then
    '            MsgBox "No Existe la Fecha Actual en el Anuario", vbCritical, "Buscar Fecha Anuario"
            Else
                msg_Fecha = mdbDiario.Recordset.Fields("Fecha")
                msg_Hora = mdbDiario.Recordset.Fields("Hora")
                msg_Mensaje = mdbDiario.Recordset.Fields("Mensaje")
                msg_Tipo_Msg = "Urgente"
            End If
        End If
        Unload Me
    Else
        Frame2.Enabled = False
        '--------- Botones ------------
        cmdUpdate.Caption = "&Modificar"
        cmdAdd.Enabled = True
        cmdDelete.Enabled = True
        cmdClose.Caption = "&Cerrar"
        mdbDiario.Enabled = True
        '------------------------------'
        mdbDiario.Refresh
        mdbDiario.Recordset.MoveFirst
    End If
End Sub

Private Sub Form_Load()
    mdbDiario.DatabaseName = Base_de_Datos
    mdbDiario.Refresh
    Frame2.Enabled = False
    Call cmdHoy_Click
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    If Frame2.Enabled = True Then
        lblFecha.Caption = DateClicked
    End If
    lblfecha_M.Caption = DateClicked
End Sub

