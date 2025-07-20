VERSION 5.00
Begin VB.Form Acerca_de 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Acerca de ..."
   ClientHeight    =   3705
   ClientLeft      =   1470
   ClientTop       =   1875
   ClientWidth     =   5865
   ClipControls    =   0   'False
   Icon            =   "Acerca_de.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Tag             =   "Acerca de Project1"
   Begin VB.Frame Frame_Base_Datos 
      Caption         =   "Frame_Base_Datos"
      Height          =   1335
      Left            =   1800
      TabIndex        =   8
      Top             =   1320
      Visible         =   0   'False
      Width           =   2775
      Begin VB.TextBox txtRif_Empresa 
         DataField       =   "Rif_Empresa"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   2535
      End
      Begin VB.TextBox txtNombre_Empresa 
         DataField       =   "Nombre_Empresa"
         DataSource      =   "Data3"
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   2535
      End
      Begin VB.Data Data3 
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
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      Height          =   285
      Left            =   840
      TabIndex        =   7
      Text            =   "  E-mail : kyrb2000@gmail.com | Kenner Roa (0426)-832-75-74"
      Top             =   840
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00FFFF00&
      Cancel          =   -1  'True
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   4245
      TabIndex        =   0
      Tag             =   "Aceptar"
      Top             =   2745
      Width           =   1467
   End
   Begin VB.CommandButton cmdSysInfo 
      Caption         =   "Info. del &sistema..."
      Height          =   345
      Left            =   4260
      TabIndex        =   1
      Tag             =   "Info. del &sistema..."
      Top             =   3195
      Width           =   1452
   End
   Begin VB.Label lblVersion 
      Caption         =   "Versi�n"
      Height          =   225
      Left            =   1080
      TabIndex        =   3
      Tag             =   "Versi�n"
      Top             =   840
      Width           =   4095
   End
   Begin VB.Label lblAutorizado 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   960
      TabIndex        =   6
      Top             =   1320
      Width           =   4095
   End
   Begin VB.Image Gestion 
      BorderStyle     =   1  'Fixed Single
      Height          =   735
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   810
   End
   Begin VB.Label lblDescription 
      Caption         =   "Se autoriza el uso de este producto a :"
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   1050
      TabIndex        =   5
      Tag             =   "Descripci�n de la aplicaci�n"
      Top             =   1125
      Width           =   4095
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "T�tulo de la aplicaci�n"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   720
      Left            =   1050
      TabIndex        =   4
      Tag             =   "T�tulo de la aplicaci�n"
      Top             =   120
      Width           =   4095
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   225
      X2              =   5657
      Y1              =   2550
      Y2              =   2550
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   240
      X2              =   5657
      Y1              =   2565
      Y2              =   2565
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   $"Acerca_de.frx":0442
      ForeColor       =   &H00000000&
      Height          =   1065
      Left            =   255
      TabIndex        =   2
      Tag             =   "Advertencia: ..."
      Top             =   2625
      Width           =   3870
   End
End
Attribute VB_Name = "Acerca_de"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
'-- Autor : Kenner Y. Roa B. Telefono (0426)-832.75.74
'-- E-mail : kyrb2000@gmail.com / kyrb2000@hotmail.com
'-- Fecha : 22/05/2002 (Caracas / Venezuela)
'-----------------------------------------------------------
Dim WEmpresa As String
Dim WRif As String
' Opciones de seguridad de clave del Registro...
Const KEY_ALL_ACCESS = &H2003F

' Tipos ROOT de claves del Registro...
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1                         ' cadena terminada en valor nulo Unicode
Const REG_DWORD = 4                      ' n�mero de 32 bits

Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Sub Form_Activate()
    Data3.DatabaseName = Base_de_Datos
    Data3.Refresh
    WEmpresa = txtNombre_Empresa
    '"Instituto Universitario de Tecnolog�a de Administraci�n Industrial (IUTA)"
    WRif = txtRif_Empresa
    '"Regi�n Capital"
    lblAutorizado.Caption = Chr(10) & Chr(13) & WEmpresa & Chr(10) & Chr(13) & Chr(13) & WRif
End Sub

Private Sub Form_Load()
    lblVersion.Caption = "Versi�n " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = MDIPasantia.Caption '"Sistema de Control de Pasant�as" 'App.Title
    Gestion.Picture = LoadPicture(App.Path & "\Logo_Bat.bmp")
End Sub

Private Sub cmdSysInfo_Click()
        Call StartSysInfo
End Sub

Private Sub cmdOK_Click()
        Unload Me
End Sub

Public Sub StartSysInfo()
    On Error GoTo SysInfoErr
        Dim rc As Long
        Dim SysInfoPath As String
        ' Intentar obtener el nombre y la ruta del programa en el Registro...
        If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
        ' Intentar obtener s�lo la ruta del programa en el Registro...
        ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
                ' Validar la existencia de versi�n conocida de 32 bits de archivo
                If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
                        SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
                ' Error: no se encuentra el archivo...
                Else
                        GoTo SysInfoErr
                End If
        ' Error: no se encuentra la entrada del Registro...
        Else
                GoTo SysInfoErr
        End If
        Call Shell(SysInfoPath, vbNormalFocus)
        Exit Sub
SysInfoErr:
        MsgBox "La informaci�n del sistema no est� disponible en este momento", vbOKOnly
End Sub

Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
        Dim i As Long                                           ' Contador de bucle
        Dim rc As Long                                          ' C�digo de retorno
        Dim hKey As Long                                        ' Controlador a una clave de Registro abierta
        Dim hDepth As Long                                      '
        Dim KeyValType As Long                                  ' Tipo de datos de una clave del Registro
        Dim tmpVal As String                                    ' Almacenamiento temporal para un valor de clave del Registro
        Dim KeyValSize As Long                                  ' Tama�o de variable de clave del Registro
        '------------------------------------------------------------
        ' Abrir RegKey bajo KeyRoot {HKEY_LOCAL_MACHINE...}
        '------------------------------------------------------------
        rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Abrir la clave del Registro
        

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Controlar error...
        

        tmpVal = String$(1024, 0)                             ' Asignar espacio de variable
        KeyValSize = 1024                                       ' Marcar tama�o de variable
        

        '------------------------------------------------------------
        ' Obtener valor de clave del Registro...
        '------------------------------------------------------------
        rc = RegQueryValueEx(hKey, SubKeyRef, 0, KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                                                

        If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Controlar errores
        

        tmpVal = VBA.Left(tmpVal, InStr(tmpVal, VBA.Chr(0)) - 1)
        '------------------------------------------------------------
        ' Determinar el tipo de valor de clave para conversi�n...
        '------------------------------------------------------------
        Select Case KeyValType                                  ' Buscar tipos de datos...
        Case REG_SZ                                             ' Tipo de datos String de clave del Registro
                KeyVal = tmpVal                                     ' Copiar valor String
        Case REG_DWORD                                          ' Tipo de datos Double Word de clave del Registro
                For i = Len(tmpVal) To 1 Step -1                    ' Convertir cada bit
                        KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Generar valor car�cter a car�cter
                Next
                KeyVal = Format$("&h" + KeyVal)                     ' Convertir Double Word a String
        End Select
        

        GetKeyValue = True                                      ' Operaci�n realizada correctamente
        rc = RegCloseKey(hKey)                                  ' Cerrar clave del Registro
        Exit Function                                           ' Salir
        

GetKeyError:    ' Limpiar despu�s de que se produzca un error...
        KeyVal = ""                                             ' Establecer el valor de retonor a la cadena vac�a
        GetKeyValue = False                                     ' La operaci�n no se ha realizado correctamente
        rc = RegCloseKey(hKey)                                  ' Cerrar clave del Registro
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.Visible = False
End Sub

Private Sub Form_DblClick()
    Text1.BackColor = &H80000018
    Text1.Visible = True
End Sub

Private Sub Gestion_DblClick()
    Text1.BackColor = &H80000018
    Text1.Visible = True
End Sub

Private Sub Gestion_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.BackColor = &H80000018
    Text1.Visible = True
End Sub

Private Sub lblAutorizado_DblClick()
    Text1.BackColor = &H80000018
    Text1.Visible = True
End Sub

Private Sub lblAutorizado_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.BackColor = &H80000018
    Text1.Visible = True
End Sub

Private Sub lblDescription_DblClick()
    Text1.BackColor = &H80000018
    Text1.Visible = True
End Sub

Private Sub lblDisclaimer_DblClick()
    Text1.BackColor = &H80000018
    Text1.Visible = True
End Sub

Private Sub lblDisclaimer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Text1.Visible = False
End Sub

Private Sub lblTitle_DblClick()
    Text1.BackColor = &H80000018
    Text1.Visible = True
End Sub

Private Sub lblVersion_DblClick()
    Text1.BackColor = &H80000018
    Text1.Visible = True
End Sub
