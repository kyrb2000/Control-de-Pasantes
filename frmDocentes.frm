VERSION 5.00
Begin VB.Form frmDocentes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Docentes"
   ClientHeight    =   4275
   ClientLeft      =   1305
   ClientTop       =   1725
   ClientWidth     =   8925
   Icon            =   "frmDocentes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4275
   ScaleWidth      =   8925
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   0
      TabIndex        =   13
      Top             =   2760
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
         Picture         =   "frmDocentes.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   14
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
         Picture         =   "frmDocentes.frx":0869
         Style           =   1  'Graphical
         TabIndex        =   15
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
         Picture         =   "frmDocentes.frx":0C84
         Style           =   1  'Graphical
         TabIndex        =   16
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
         Picture         =   "frmDocentes.frx":10AC
         Style           =   1  'Graphical
         TabIndex        =   17
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
         Picture         =   "frmDocentes.frx":14E9
         Style           =   1  'Graphical
         TabIndex        =   18
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
         Picture         =   "frmDocentes.frx":1952
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.TextBox txtTelefono 
         DataField       =   "Telefono"
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
         MaxLength       =   50
         TabIndex        =   11
         Top             =   2160
         Width           =   6855
      End
      Begin VB.TextBox txtNumeroD_ID 
         DataField       =   "NumeroD_ID"
         DataSource      =   "Data1"
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
         Height          =   405
         Left            =   4680
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin VB.Data Data3 
         Caption         =   "Tabla_Gerenar"
         Connect         =   "Access"
         DatabaseName    =   "C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
         DefaultCursorType=   0  'DefaultCursor
         DefaultType     =   2  'UseODBC
         Exclusive       =   0   'False
         Height          =   350
         Left            =   3360
         Options         =   0
         ReadOnly        =   0   'False
         RecordsetType   =   1  'Dynaset
         RecordSource    =   "Tabla_Gerenar"
         Top             =   960
         Visible         =   0   'False
         Width           =   2290
      End
      Begin VB.TextBox txtCargo 
         DataField       =   "Cargo"
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
         MaxLength       =   50
         TabIndex        =   8
         Top             =   1680
         Width           =   6855
      End
      Begin VB.TextBox txtApellidos 
         DataField       =   "Apellidos"
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
         MaxLength       =   50
         TabIndex        =   7
         Top             =   1200
         Width           =   6855
      End
      Begin VB.TextBox txtNombres 
         DataField       =   "Nombres"
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
         MaxLength       =   50
         TabIndex        =   6
         Top             =   720
         Width           =   6855
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
         Left            =   1440
         MaxLength       =   9
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label lblLabels 
         Caption         =   "Telefono:"
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
         TabIndex        =   12
         Top             =   2160
         Width           =   1815
      End
      Begin VB.Label lblLabels 
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
         Height          =   375
         Index           =   6
         Left            =   3480
         TabIndex        =   10
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cargo:"
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
         TabIndex        =   4
         Top             =   1680
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Apellidos:"
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
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Nombres:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Cedula:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Data Data1 
      Align           =   2  'Align Bottom
      Connect         =   "Access"
      DatabaseName    =   "C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   0
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Docentes"
      Top             =   3930
      Width           =   8925
   End
End
Attribute VB_Name = "frmDocentes"
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
        Data1.Recordset.FindFirst Buscar
        If Not Data1.Recordset.NoMatch Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    Frame1.Enabled = True
    Data1.Recordset.AddNew
    txtCedula = Cedula
    txtNumeroD_ID = "."
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
    txtCedula.SetFocus
End Sub

Private Sub cmdBuscar_Click()
    Dim Buscar As String
    If (Data1.Recordset.AbsolutePosition + 1) = 0 Then
        Exit Sub
    End If
    Buscar = InputBox("Introduzca La Cedula :", "Busqueda de Datos")
    If Buscar <> "" Then
        Buscar = "[Cedula]" & "=" & "'" & Buscar & "'"
        Data1.Recordset.FindFirst Buscar
    End If
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

Private Sub cmdRefresh_Click()
  'esto sólo es necesario para aplicaciones multiusuario
  Data1.Refresh
End Sub

Private Sub cmdImprimir_Click()
    Dim fReportes As New frmReportes
    fReportes.Caption = "Reportes de Docentes"
    fReportes.Tipo_Reporte = "DOC"
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
        Frame1.Enabled = False
        If txtNumeroD_ID = "." Then
            txtNumeroD_ID = Data3.Recordset.Fields("NumeroD_ID")
            NumeroD_ID = Val(txtNumeroD_ID) + 1
            Data3.Recordset.Edit
            Data3.Recordset.Fields("NumeroD_ID") = NumeroD_ID
            Data3.Recordset.Update
        End If
        Data1.UpdateRecord
        Data1.Recordset.Bookmark = Data1.Recordset.LastModified
        '--------- Botones ------------
        cmdUpdate.Caption = "&Modificar"
        cmdAdd.Enabled = True
        cmdDelete.Enabled = True
        cmdBuscar.Enabled = True
        cmdImprimir.Enabled = True
        cmdClose.Caption = "&Cerrar"
        Data1.Enabled = True
        '------------------------------'
    Else
        If Not Data1.Recordset.EOF Then
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
        Data1.Enabled = True
        '------------------------------'
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
    Case vbDataActionMovePrevious
    Case vbDataActionMoveNext
    Case vbDataActionMoveLast
    Case vbDataActionAddNew
    Case vbDataActionUpdate
    Case vbDataActionDelete
    Case vbDataActionFind
    Case vbDataActionBookmark
    Case vbDataActionClose
  End Select
End Sub

Private Sub txtApellidos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtCargo.SetFocus
End Sub

Private Sub txtCargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then cmdUpdate.SetFocus
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNombres.SetFocus
End Sub

Private Sub txtNombres_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtApellidos.SetFocus
End Sub
