VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAlumnosPagoTaller2 
   Caption         =   "Ingreso de Alumnos al Taller de Pasantías"
   ClientHeight    =   5085
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8265
   Icon            =   "frmAlumnosPagoTaller2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   8265
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   3855
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   8055
      Begin VB.CheckBox txtTaller 
         Alignment       =   1  'Right Justify
         Caption         =   "Realizó el Taller"
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
         Left            =   480
         TabIndex        =   6
         Top             =   3360
         Width           =   2175
      End
      Begin VB.TextBox txtDireccion 
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
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   2400
         Width           =   6135
      End
      Begin VB.TextBox txtTelefono 
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
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1920
         Width           =   6135
      End
      Begin VB.TextBox txtApellidos 
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
         MaxLength       =   150
         TabIndex        =   3
         Top             =   1440
         Width           =   6135
      End
      Begin VB.TextBox txtSección 
         Alignment       =   2  'Center
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
         Left            =   5160
         MaxLength       =   5
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtCedula 
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
         MaxLength       =   10
         TabIndex        =   0
         Top             =   240
         Width           =   2055
      End
      Begin VB.TextBox txtNombres 
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
         MaxLength       =   150
         TabIndex        =   2
         Top             =   840
         Width           =   6135
      End
      Begin VB.Label Label4 
         Caption         =   "Dirección:"
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
         Left            =   480
         TabIndex        =   22
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Label Label3 
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
         Left            =   480
         TabIndex        =   21
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label2 
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
         Left            =   480
         TabIndex        =   20
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Sección:"
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
         Left            =   3960
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label1 
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
         Height          =   255
         Left            =   480
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
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
         Left            =   480
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1110
      Left            =   120
      MouseIcon       =   "frmAlumnosPagoTaller2.frx":0442
      TabIndex        =   15
      Top             =   3960
      Width           =   8055
      Begin MSForms.CommandButton CmdOpcion 
         Height          =   975
         Index           =   7
         Left            =   5400
         TabIndex        =   11
         Top             =   120
         Width           =   1215
         Caption         =   "IMPRIMIR"
         Size            =   "2143;1720"
         MousePointer    =   1
         Picture         =   "frmAlumnosPagoTaller2.frx":074C
         FontName        =   "Times New Roman"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CmdOpcion 
         Height          =   975
         Index           =   2
         Left            =   4080
         TabIndex        =   10
         Top             =   120
         Width           =   1215
         Caption         =   "BUSCAR"
         Size            =   "2143;1720"
         MousePointer    =   1
         Picture         =   "frmAlumnosPagoTaller2.frx":0B77
         FontName        =   "Times New Roman"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CmdOpcion 
         Height          =   975
         Index           =   4
         Left            =   2760
         TabIndex        =   9
         Top             =   120
         Width           =   1215
         Caption         =   "ELIMINAR"
         Size            =   "2143;1720"
         MousePointer    =   1
         Picture         =   "frmAlumnosPagoTaller2.frx":0FA3
         FontName        =   "Times New Roman"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CmdOpcion 
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   1215
         Caption         =   "AGREGAR"
         Size            =   "2143;1720"
         MousePointer    =   1
         Picture         =   "frmAlumnosPagoTaller2.frx":141C
         FontName        =   "Times New Roman"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CmdOpcion 
         Height          =   975
         Index           =   3
         Left            =   1440
         TabIndex        =   8
         Top             =   120
         Width           =   1215
         Caption         =   "MODIFICAR"
         Size            =   "2143;1720"
         MousePointer    =   1
         Picture         =   "frmAlumnosPagoTaller2.frx":182D
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CmdOpcion 
         Height          =   975
         Index           =   6
         Left            =   6720
         TabIndex        =   12
         Top             =   120
         Width           =   1215
         Caption         =   "SALIR"
         Size            =   "2143;1720"
         MousePointer    =   1
         Picture         =   "frmAlumnosPagoTaller2.frx":1C65
         FontName        =   "Arial"
         FontEffects     =   1073741825
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CmdOpcion 
         Height          =   975
         Index           =   1
         Left            =   1440
         TabIndex        =   13
         Top             =   120
         Width           =   1215
         VariousPropertyBits=   25
         Caption         =   "GUARDAR"
         Size            =   "2143;1720"
         MousePointer    =   1
         Picture         =   "frmAlumnosPagoTaller2.frx":209C
         FontName        =   "Times New Roman"
         FontEffects     =   1073750017
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
      Begin MSForms.CommandButton CmdOpcion 
         Height          =   975
         Index           =   5
         Left            =   6720
         TabIndex        =   14
         Top             =   120
         Width           =   1215
         VariousPropertyBits=   25
         Caption         =   "CANCELAR"
         Size            =   "2143;1720"
         MousePointer    =   1
         Picture         =   "frmAlumnosPagoTaller2.frx":251D
         FontName        =   "Times New Roman"
         FontEffects     =   1073750017
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         ParagraphAlign  =   3
         FontWeight      =   700
      End
   End
End
Attribute VB_Name = "frmAlumnosPagoTaller2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sw As Boolean
Private Puntero As Integer

Private Sub Limpiar()
    txtCedula = ""
    txtSección = ""
    txtNombres = ""
    txtApellidos = ""
    txtDireccion = ""
    txtTelefono = ""
    txtTaller.Value = 0
    Frame1.Enabled = False
End Sub

Private Sub Imprimir()
    On Error GoTo Error_SQL
    Sección = InputBox("Indique la Sección:", Me.Caption)
    If Sección = "" Then
        SQL = "SELECT *" & _
              "FROM Alumnos_Pago_Taller "
    Else
        SQL = "SELECT *" & _
              "FROM Alumnos_Pago_Taller " & _
              "WHERE Sección='" & Sección & "'"
    End If
'---------------------------------------------------------------------------------------------
    Set Rep_TAlumnos.DataSource = DEPago_de_Taller.Connection1.Execute(SQL)
    Rep_TAlumnos.Caption = Me.Caption
    Rep_TAlumnos.Show
Error_SQL:
End Sub

Private Sub Cancelar()
    Botones (True)
    txtCedula.Locked = False
    Limpiar
    CmdOpcion(0).SetFocus
End Sub

Private Sub Validar()
    If txtCedula = "" Or txtNombres = "" Or txtSección = "" Or txtApellidos = "" Then
        Sw = True
    Else
        Sw = False
    End If
End Sub

Private Sub Ubicar()
    If txtCedula = "" Then
        txtCedula.SetFocus
    ElseIf txtNombres = "" Then
        txtNombres.SetFocus
    ElseIf txtSección = "" Then
        txtSección.SetFocus
    ElseIf txtApellidos = "" Then
        txtApellidos.SetFocus
    ElseIf txtDireccion = "" Then
        txtDireccion.SetFocus
    ElseIf txtTelefono = "" Then
        txtTelefono.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Sw = False
    Cargar
End Sub

Private Sub CmdOpcion_Click(Index As Integer)
    Select Case Index
        Case o
            Agregar
        Case 1
            Guardar
        Case 2
            Buscar
        Case 3
            Modificar
        Case 4
            Eliminar
        Case 5
            Cancelar
        Case 6
            If MsgBox("¿Desea Salir de la Aplicación?", vbYesNo + vbExclamation, Me.Caption) = vbYes Then
                Unload Me
            End If
        Case 7
            Imprimir
    End Select
End Sub

Private Sub Agregar()
    Cancelar
    Botones (False)
    Frame1.Enabled = True
    txtCedula.SetFocus
    '---------------------------
End Sub

Private Sub Guardar()
    Validar
    If Sw = True Then
        MsgBox "Debe Completar los Datos", vbCritical, Me.Caption
        Ubicar
        Exit Sub
    Else
        If Puntero = 1 Then
            SQL = "Select * From Alumnos_Pago_Taller Where Cedula='" & txtCedula & "'"
            Recor.Open SQL, Conex, adOpenDynamic, adLockOptimistic
            With Recor
                !Sección = UCase(txtSección)
                !Nombres = UCase(txtNombres)
                !Apellidos = UCase(txtApellidos)
                !Direccion = UCase(txtDireccion)
                !Telefono = txtTelefono
                !Taller = IIf(txtTaller.Value = 0, False, True)
                .Update
            End With
            MsgBox "Los Datos Han Sido Actualizados", vbInformation, Me.Caption
            CmdOpcion(1).Enabled = False
            Call Limpiar
            Recor.Close
        Else
            SQL = "Select * From Alumnos_Pago_Taller Where Cedula='" & txtCedula & "'"
            Recor.Open SQL, Conex, adOpenDynamic
            If Not Recor.EOF Then
                MsgBox "La Cedula YA EXISTE los Datos NO Fueron Guardados", vbCritical, Me.Caption
                Recor.Close
                txtCedula.SetFocus
                Exit Sub
            End If
            Recor.Close
            SQL = ""
            SQL = "Insert Into Alumnos_Pago_Taller(Nacionalidad,Cedula,Sección,Nombres,Apellidos,Direccion,Telefono,Taller)"
                'x = IIf(txtTaller.Value = 0, False, True)
            SQL = SQL & "Values ('V','" & txtCedula & "','" & UCase(txtSección) & "','" & UCase(txtNombres) & "','" & UCase(txtApellidos) & "','" & UCase(txtDireccion) & "','" & UCase(txtTelefono) & "'," & txtTaller.Value & ")"
            Conex.Execute (SQL)
            MsgBox "Los Datos Fueron Guardados", vbInformation, Me.Caption
            CmdOpcion(1).Enabled = False
            Call Limpiar
        End If
        Botones (True)
    End If
End Sub

Private Sub Buscar()
    Cancelar
    Num = InputBox("Ingrese el Número de Cedula:", Me.Caption)
    If Num = "" Then
        CmdOpcion(1).Enabled = False
        Exit Sub
    End If
    SQL = "Select * From Alumnos_Pago_Taller Where Cedula='" & Num & "'"
    Recor.Open SQL, Conex, adOpenDynamic
    If Not Recor.EOF Then
        With Recor
            txtCedula = !Cedula
            txtSección = !Sección
            txtNombres = !Nombres
            txtApellidos = !Apellidos
            txtDireccion = IIf(IsNull(!Direccion), "-", !Direccion)
            txtTelefono = IIf(IsNull(!Telefono), "-", !Telefono)
            txtTaller.Value = IIf(!Taller, 1, 0)
            Recor.Close
        End With
    Else
        MsgBox "Cedula NO EXISTENTE", vbCritical, Me.Caption
        Recor.Close
        CmdOpcion(1).Enabled = False
        Num = ""
    End If
End Sub

Private Sub Modificar()
    If txtCedula.Text <> "" Then
'        If MsgBox("Desea Realmente Modificar este Registro", vbYesNo + vbExclamation, Me.Caption) = vbYes Then
            txtCedula.Locked = True
            Frame1.Enabled = True
            txtSección.SetFocus
            Botones (False)
            Puntero = 1
'        Else
'            Frame1.Enabled = False
'            CmdOpcion(1).Enabled = False
'            Puntero = 0
'        End If
    Else
        Buscar
        If Num = "" Then
            Exit Sub
        Else
            Modificar
        End If
    End If
End Sub
Private Sub Eliminar()
    If txtCedula.Text <> "" Then
        If MsgBox("Desea Realmente Eliminar este Registro", vbYesNo + vbCritical, Me.Caption) = vbYes Then
            SQL = "Delete * From Alumnos_Pago_Taller Where Cedula = '" & (txtCedula) & "'"
            Conex.Execute (SQL)
            Cancelar
            MsgBox "Los Datos Fueron Eliminados", vbInformation, Me.Caption
            Call Limpiar
        Else
            MsgBox "Se ha Cancelado la Eliminación", vbInformation, Me.Caption
            Call Limpiar
            Exit Sub
        End If
    Else
        Buscar
        If Num = "" Then
            Exit Sub
        Else
            Eliminar
        End If
    End If
End Sub

Sub Botones(Cancelar As Boolean)
    '--------[ Botones ]--------
    If Not Cancelar Then
        CmdOpcion(0).Enabled = False
        CmdOpcion(1).Enabled = True
        CmdOpcion(2).Enabled = False
        CmdOpcion(3).Visible = False
        CmdOpcion(3).Enabled = False
        CmdOpcion(4).Enabled = False
        CmdOpcion(5).Enabled = True
        CmdOpcion(6).Visible = False
        CmdOpcion(6).Enabled = False
        CmdOpcion(7).Enabled = False
    Else
        CmdOpcion(0).Enabled = True
        CmdOpcion(1).Enabled = False
        CmdOpcion(2).Enabled = True
        CmdOpcion(3).Visible = True
        CmdOpcion(3).Enabled = True
        CmdOpcion(4).Enabled = True
        CmdOpcion(5).Enabled = False
        CmdOpcion(6).Visible = True
        CmdOpcion(6).Enabled = True
        CmdOpcion(7).Enabled = True
    End If
    '---------------------------
End Sub

Private Sub txtApellidos_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTelefono.SetFocus
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtCedula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtSección.SetFocus
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtDireccion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtTaller.SetFocus
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtNombres_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtApellidos.SetFocus
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtSección_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtNombres.SetFocus
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtTaller_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CmdOpcion(1).SetFocus
End Sub

Private Sub txtTelefono_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then txtDireccion.SetFocus
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
