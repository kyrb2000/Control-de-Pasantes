VERSION 5.00
Begin VB.MDIForm MDIPasantia 
   BackColor       =   &H8000000C&
   Caption         =   "Sistema Automatizado de Control de Pasantes"
   ClientHeight    =   3195
   ClientLeft      =   4275
   ClientTop       =   3045
   ClientWidth     =   4680
   Icon            =   "MDIPasantia.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox ScrollBar 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      ScaleHeight     =   360
      ScaleMode       =   0  'User
      ScaleWidth      =   4620
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   4680
      Begin VB.CommandButton cmdDiario 
         Height          =   360
         Left            =   3720
         Picture         =   "MDIPasantia.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Diario..."
         Top             =   0
         Width           =   510
      End
      Begin VB.CommandButton cmdCalculadora 
         Height          =   360
         Left            =   3150
         Picture         =   "MDIPasantia.frx":0B58
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Calculadora..."
         Top             =   0
         Width           =   510
      End
      Begin VB.TextBox txtDiaActual 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1680
         TabIndex        =   2
         Text            =   "txtDiaActual"
         Top             =   50
         Width           =   1335
      End
      Begin VB.TextBox txtHora 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   0
         TabIndex        =   1
         Text            =   "txtHora"
         Top             =   50
         Width           =   1455
      End
      Begin VB.Line Line2 
         X1              =   3120
         X2              =   3120
         Y1              =   0
         Y2              =   360
      End
      Begin VB.Line Line1 
         X1              =   1560
         X2              =   1560
         Y1              =   0
         Y2              =   360
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   360
   End
   Begin VB.Menu mnuDatos 
      Caption         =   "&Datos"
      Begin VB.Menu mnuDatos_Ficha_de_Pasantía 
         Caption         =   "&Ficha de Pasantía"
      End
      Begin VB.Menu mnuDatos_Agregar_Alumnos 
         Caption         =   "&Agregar Alumnos"
      End
      Begin VB.Menu mnuDatos_Docentes 
         Caption         =   "&Docentes"
      End
      Begin VB.Menu mnuDatos_Centros_de_Pasantias 
         Caption         =   "Centros de Pasantías"
      End
      Begin VB.Menu mnuDatos_Especialidad 
         Caption         =   "&Especialidad"
      End
      Begin VB.Menu mnuLinea0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDatos_Carta_de_Postulación 
         Caption         =   "Carta de &Postulación"
      End
      Begin VB.Menu mnuDatos_Carta_de_Presentación 
         Caption         =   "&Carta de Presentación"
      End
   End
   Begin VB.Menu mnuReportes 
      Caption         =   "&Reportes"
      Begin VB.Menu mnuReportes_Formatos 
         Caption         =   "F&ormatos"
         Begin VB.Menu mnuFormatos_Ficha_Personal_de_Pasante 
            Caption         =   "&Ficha Personal de Pasante"
         End
         Begin VB.Menu mnuFormatos_Ficha_Personal_del_Proceso_de_Pasantía 
            Caption         =   "F&icha Personal del Proceso de Pasantía"
         End
         Begin VB.Menu mnuFormato_Evaluación_del_Pasante 
            Caption         =   "&Evaluación del Pasante"
         End
         Begin VB.Menu mnuFormato_Control_de_Asistencia 
            Caption         =   "&Control de Asistencia"
         End
         Begin VB.Menu mnuFormato_Cronograma_para_Taller_de_Pasantías 
            Caption         =   "C&ronograma para Taller de Pasantías"
         End
         Begin VB.Menu mnuFormato_Control_de_Visitas_al_Pasante 
            Caption         =   "Co&ntrol de Visitas al Pasante"
         End
         Begin VB.Menu mnuFormato_AutoEvaluación_del_Pasante 
            Caption         =   "&AutoEvaluación del Pasante"
         End
         Begin VB.Menu mnuConstancia_de_Asistencia_a_Taller_de_Pasantias 
            Caption         =   "Constancia de Asistencia al &Taller de Pasantías"
         End
         Begin VB.Menu mnuFormato_Control_de_Actividades_Diarias 
            Caption         =   "Control de Actividades Diarias"
         End
      End
      Begin VB.Menu mnuReportes_Fichas_de_Pasantía 
         Caption         =   "&Fichas de Pasantía"
      End
      Begin VB.Menu mnuReportes_Solvencia_de_los_Pasantes 
         Caption         =   "&Solvencia de los Pasantes"
      End
      Begin VB.Menu mnuReportes_Registro_del_Pasante 
         Caption         =   "&Registro del Pasante"
      End
      Begin VB.Menu mnuReportes_Registro_del_Tutor_Academico 
         Caption         =   "Registro del &Tutor Academico"
      End
      Begin VB.Menu mnuLinea1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportes_Casos_Especiales 
         Caption         =   "&Casos Especiales"
      End
      Begin VB.Menu mnuReportes_Centros_de_Pasantias 
         Caption         =   "C&entros de Pasantías"
      End
      Begin VB.Menu mnuReportes_Pendientes_X_Pasantías 
         Caption         =   "Pe&ndiente por Pasantías"
      End
      Begin VB.Menu mnuReportes_Total_Pendiente_Pasantias_x_Seccion 
         Caption         =   "&Total Pendiente Pasantías por Sección"
      End
      Begin VB.Menu mnuReportes_Realizando_Pasantías 
         Caption         =   "Rea&lizando Pasantías"
      End
      Begin VB.Menu mnuReportes_Total_Realizando_Pasantias_x_Seccion 
         Caption         =   "Tota&l Realizando Pasantías por Sección"
      End
      Begin VB.Menu mnuLinea3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportes_Alumnos 
         Caption         =   "&Alumnos..."
      End
      Begin VB.Menu mnuReportes_Docentes 
         Caption         =   "&Docentes..."
      End
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "&Opciones"
      Begin VB.Menu mnuOpciones_Cargar_Datos 
         Caption         =   "&Cargar Datos"
      End
      Begin VB.Menu mnuOpciones_Mantenimiento 
         Caption         =   "&Mantenimiento"
      End
      Begin VB.Menu mnuLinea2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpciones_Cambios_Generales 
         Caption         =   "Cambios &Generales"
      End
      Begin VB.Menu mnuOpciones_Selector_de_Queries 
         Caption         =   "&Selector de Queries"
      End
   End
   Begin VB.Menu mnuAyuda 
      Caption         =   "Ay&uda"
      Begin VB.Menu mnuAyuda_Contenido 
         Caption         =   "&Contenido..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAcerca_de 
         Caption         =   "A&cerca de ..."
      End
   End
   Begin VB.Menu mnuSalir 
      Caption         =   "&Salir"
   End
End
Attribute VB_Name = "MDIPasantia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
'-- Autor : Kenner Y. Roa B. Telefono (0212) 433.55.14
'-- E-mail : kyrb2000@cantv.net / kyrb2000@hotmail.com
'-- Fecha : 22/05/2002 (Caracas / Venezuela)
'-----------------------------------------------------------
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub cmdCalculadora_Click()
    Dim RetVal
    RetVal = Shell("C:\WINDOWS\CALC.EXE", vbHide)    ' Ejecuta Calculadora.
    AppActivate RetVal         ' Activa la Calculadora.
End Sub

Private Sub cmdDiario_Click()
    Call centrarform(frmAnuario)
End Sub

Private Sub MDIForm_Activate()
    On Error GoTo Error_de_Formulario
    Tipo_Fecha = "dd/mm/yyyy"
Error_de_Formulario:

End Sub

Private Sub MDIForm_DblClick()
    If ScrollBar.Visible = False Then
        ScrollBar.Visible = True
    Else
        ScrollBar.Visible = False
    End If
End Sub

Private Sub MDIForm_Load()
    DataEnvironment1.Connection1.Open "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\Datos_Alumnos.mdb"
    MDIPasantia.Caption = MDIPasantia.Caption & " " & X_Periodo
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim Respuesta As Integer
    Respuesta = MsgBox("¿Desea salir de la aplicación?", 4, Me.Caption)
    If Respuesta = vbYes Then
        End
    Else
        Cancel = True
    End If
End Sub

Private Sub mnuAcerca_de_Click()
    Acerca_de.Show 1
End Sub

Private Sub mnuAyuda_Contenido_Click()
    Dim nRet As Integer
    App.HelpFile = App.Path & "\Pasantías.hlp"
    If Len(App.HelpFile) = 0 Then
        MsgBox "No se puede mostrar el contenido de la Ayuda. No hay Ayuda asociada a este proyecto.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If
End Sub

Private Sub mnuDatos_Agregar_Alumnos_Click()
    Call centrarform(frmAlumnos)
End Sub

Private Sub mnuDatos_Carta_de_Postulación_Click()
    Dim fCartas_Postulación As New frmCartas
    fCartas_Postulación.Caption = "Cartas de Postulación"
    fCartas_Postulación.Tipo_Carta = "POS"
    Call centrarform(fCartas_Postulación)
End Sub

Private Sub mnuDatos_Carta_de_Presentación_Click()
    Dim fCartas_Presentación As New frmCartas
    fCartas_Presentación.Caption = "Cartas de Presentación"
    fCartas_Presentación.Tipo_Carta = "PRE"
    Call centrarform(fCartas_Presentación)
'    Call centrarForm(frmCartas_Presentacion)
End Sub

Private Sub mnuDatos_Centros_de_Pasantias_Click()
    Call centrarform(frmCentros_Pasantias)
End Sub

Private Sub mnuDatos_Docentes_Click()
    Call centrarform(frmDocentes)
End Sub

Private Sub mnuDatos_Especialidad_Click()
    Call centrarform(frmEspecialidad)
End Sub

Private Sub mnuDatos_Ficha_de_Pasantía_Click()
    Call centrarform(frmFicha_Pasantia)
End Sub

Private Sub mnuFormatos_Ficha_Personal_de_Pasante_Click()
        Shell "C:\Archivos de programa\Microsoft Office\Office\winword.exe C:\Archiv~1\Sistem~2\Reportes\Formato01.doc"
End Sub

Private Sub mnuFormatos_Ficha_Personal_del_Proceso_de_Pasantía_Click()
        Shell "C:\Archivos de programa\Microsoft Office\Office\winword.exe C:\Archiv~1\Sistem~2\Reportes\Formato02.doc"
End Sub

Private Sub mnuFormato_Evaluación_del_Pasante_Click()
        Shell "C:\Archivos de programa\Microsoft Office\Office\winword.exe C:\Archiv~1\Sistem~2\Reportes\Formato03.doc"
End Sub

Private Sub mnuFormato_Control_de_Asistencia_Click()
        Shell "C:\Archivos de programa\Microsoft Office\Office\winword.exe C:\Archiv~1\Sistem~2\Reportes\Formato04.doc"
End Sub

Private Sub mnuFormato_Cronograma_para_Taller_de_Pasantías_Click()
        Shell "C:\Archivos de programa\Microsoft Office\Office\winword.exe C:\Archiv~1\Sistem~2\Reportes\Formato05.doc"
End Sub

Private Sub mnuFormato_Control_de_Visitas_al_Pasante_Click()
        Shell "C:\Archivos de programa\Microsoft Office\Office\winword.exe C:\Archiv~1\Sistem~2\Reportes\Formato06.doc"
End Sub

Private Sub mnuFormato_AutoEvaluación_del_Pasante_Click()
        Shell "C:\Archivos de programa\Microsoft Office\Office\winword.exe C:\Archiv~1\Sistem~2\Reportes\Formato07.doc"
End Sub

Private Sub mnuConstancia_de_Asistencia_a_Taller_de_Pasantias_Click()
        Shell "C:\Archivos de programa\Microsoft Office\Office\winword.exe C:\Archiv~1\Sistem~2\Reportes\Formato08.doc"
End Sub

Private Sub mnuFormato_Control_de_Actividades_Diarias_Click()
    Shell "C:\Archivos de programa\Microsoft Office\Office\winword.exe C:\Archiv~1\Sistem~2\Reportes\Formato09.doc"
End Sub

Private Sub mnuOpciones_Cambios_Generales_Click()
    Call centrarform(frmCambios)
End Sub

Private Sub mnuOpciones_Cargar_Datos_Click()
    frmContrasena.Show vbModal
    If frmContrasena.Respuesta = True Then
        Call centrarform(frmCargar)
    End If
    Unload frmContrasena
End Sub

Private Sub mnuOpciones_Mantenimiento_Click()
    frmContrasena.Show vbModal
    If frmContrasena.Respuesta = True Then
        Call centrarform(frmOpciones)
    End If
    Unload frmContrasena
End Sub

Private Sub mnuOpciones_Selector_de_Queries_Click()
    Dim fBuscar As New frmBuscar
    Call centrarform(fBuscar)
End Sub

Private Sub mnuReportes_Alumnos_Click()
    Dim fReportes As New frmReportes
    fReportes.Caption = "Reportes de Alumnos"
    fReportes.Tipo_Reporte = "ALU"
    Call centrarform(fReportes)
End Sub

Private Sub mnuReportes_Casos_Especiales_Click()
    Dim fReportes As New frmReportes
    fReportes.Caption = "Reportes de Casos Especiales"
    fReportes.Tipo_Reporte = "CAS"
    Call centrarform(fReportes)
End Sub

Private Sub mnuReportes_Centros_de_Pasantias_Click()
    Dim fReportes As New frmReportes
    fReportes.Caption = "Reportes de Centros de Pasantías"
    fReportes.Tipo_Reporte = "CEN"
    Call centrarform(fReportes)
End Sub

Private Sub mnuReportes_Docentes_Click()
    Dim fReportes As New frmReportes
    fReportes.Caption = "Reportes de Docentes"
    fReportes.Tipo_Reporte = "DOC"
    Call centrarform(fReportes)
End Sub

Private Sub mnuReportes_Fichas_de_Pasantía_Click()
    Dim fReportes As New frmReportes
    fReportes.Caption = "Reportes Fichas de Pasantía"
    fReportes.Tipo_Reporte = "FIC"
    Call centrarform(fReportes)
End Sub

Private Sub mnuReportes_Pendientes_X_Pasantías_Click()
    Dim fReportes As New frmReportes
    fReportes.Caption = "Reportes Pendientes por Pasantías"
    fReportes.Tipo_Reporte = "RPP"
    Call centrarform(fReportes)
End Sub

Private Sub mnuReportes_Realizando_Pasantías_Click()
    Dim fReportes As New frmReportes
    fReportes.Caption = "Reportes Realizando Pasantías"
    fReportes.Tipo_Reporte = "REP"
    Call centrarform(fReportes)
End Sub

Private Sub mnuReportes_Registro_del_Pasante_Click()
    Dim fReportes As New frmReportes
    fReportes.Caption = "Reportes Registro del Pasante"
    fReportes.Tipo_Reporte = "REG"
    Call centrarform(fReportes)
End Sub

Private Sub mnuReportes_Registro_del_Tutor_Academico_Click()
    Dim fReportes As New frmReportes
    fReportes.Caption = "Reportes Registro del Tutor Academico"
    fReportes.Tipo_Reporte = "TUR"
    Call centrarform(fReportes)
End Sub

Private Sub mnuReportes_Solvencia_de_los_Pasantes_Click()
    Dim fReportes As New frmReportes
    fReportes.Caption = "Reportes Solvencia de los Pasantes"
    fReportes.Tipo_Reporte = "SOL"
    Call centrarform(fReportes)
End Sub

Private Sub mnuReportes_Total_Pendiente_Pasantias_x_Seccion_Click()
    Dim fReportes As New frmReportes
    fReportes.Caption = "Reporte Total Pendiente Pasantías por Sección"
    fReportes.Tipo_Reporte = "TGP"
    Call centrarform(fReportes)
End Sub

Private Sub mnuReportes_Total_Realizando_Pasantias_x_Seccion_Click()
    Dim fReportes As New frmReportes
    fReportes.Caption = "Reporte Total Realizando Pasantías por Sección"
    fReportes.Tipo_Reporte = "TGR"
    Call centrarform(fReportes)
End Sub

Private Sub mnuSalir_Click()
    Dim Respuesta As Integer
    Respuesta = MsgBox("¿Desea salir de la aplicación?", 4, Me.Caption)
    If Respuesta = vbYes Then
        End
    Else
        Cancel = True
    End If
End Sub

Private Sub Timer1_Timer()
    txtHora = Time
    txtDiaActual = Format(Date, "mm/dd/yyyy")
    msg_Fecha = Format(msg_Fecha, "mm/dd/yyyy")
    If txtDiaActual = msg_Fecha Then
        If txtHora = msg_Hora Or msg_Tipo_Msg = "Urgente" Then
            msg_Tipo_Msg = "-"
            Clipboard.Clear
            Clipboard.SetText msg_Mensaje
            frmMSNPopup.lblMensaje.Caption = msg_Mensaje
            frmMSNPopup.Show
        End If
    End If
End Sub
