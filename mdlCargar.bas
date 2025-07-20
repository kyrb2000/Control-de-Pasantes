Attribute VB_Name = "mdlCargar"
'-----------------------------------------------------------
'-- Autor : Kenner Y. Roa B. Telefono (0212) 433.55.14
'-- E-mail : kyrb2000@cantv.net
'-- Fecha : 22/05/2002 (Caracas / Venezuela)
'-----------------------------------------------------------
Public MAESTRO As Database         'Base de datos
Public MAESTRO2 As Database        'Base de datos
Public Ds1 As Dynaset
Public dynListin As Dynaset         'objeto Dynaset
Public BD As Integer
Public BaseD As String
Public Tipo_Fecha As String
Public Base_de_Datos As String
Public DSN_Reporte_Mov As String
Public DSN_Pasantias As String
Public X_Periodo As String

Sub Main()
    DSN_Pasantias = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.Path & "\Datos_Alumnos.mdb"
    Base_de_Datos = App.Path & "\Datos_Alumnos.mdb"
    DSN_Reporte_Mov = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Reportes\Movimiento.mdb;Persist Security Info=False"
    frmBienvenido.Show 1
    MDIPasantia.Show
End Sub

'------------------------------------------------------------------
'Esta función quita el nombre de archivo de una cadena ruta\archivo
'------------------------------------------------------------------
Function StripFileName(rsFileName As String) As String
  On Error Resume Next
  Dim i As Integer

  For i = Len(rsFileName) To 1 Step -1
    If Mid(rsFileName, i, 1) = "\" Then
      Exit For
    End If
  Next

  StripFileName = Mid(rsFileName, 1, i - 1)

End Function

'------------------------------------------------------------------
'Esta función quita el la extención de archivo
'------------------------------------------------------------------
Function StripFile(rsFileName As String, Caracter As String) As String
  On Error Resume Next
  Dim i As Integer

  For i = Len(rsFileName) To 1 Step -1
    If Mid(rsFileName, i, 1) = Caracter Then
      Exit For
    End If
  Next

  StripFile = Mid(rsFileName, 1, i - 1)

End Function

'-----------------------------------------------------------
'Esta subrutina Centra el formulario enviado (f)
'sobre el monitor
'-----------------------------------------------------------
Sub centrarform(f As Form)
    If f.WindowState = 1 Then
        f.WindowState = 0
    End If
    If f.WindowState <> 2 Then
        f.Move (Screen.Width - f.Width) / 2, 0 '(Screen.Height - f.Height) / 2
    End If
End Sub

'-----------------------------------------------------------
'-- Abre la base de datos cuyo nombre es "BDName" (pasado
'-- como parámetro
'----------------------------------------------------------
Public Function Abrir_BaseDatos(BDName As String, Master As Integer) As Integer
  Dim MiNombreUsuario As String
  Dim MiNombreGrupo As String
  Dim criterio As String
  Dim wrkJet As Workspace
  On Error GoTo Abrir_BDErr
  MiNombreUsuario = ""
  MiNombreGrupo = ""
  criterio = "ODBC;UID=" & MiNombreUsuario & ";PWS=" & MiNombreGrupo
  If Master = 1 Then
    Set MAESTRO = OpenDatabase(BDName, , , criterio)
  Else
    Set MAESTRO2 = OpenDatabase(BDName, , , criterio)
  End If
  'exito
  Abrir_BaseDatos = True
  GoTo Abrir_BDEnd

Abrir_BDErr:
  Abrir_BaseDatos = False
  Resume Abrir_BDEnd

Abrir_BDEnd:

End Function

'----------------------------------------------------------
'-- Abre las Tablas de la base de datos
'-- cuyo nombre es "TablaName" (pasado
'-- como parámetro
'----------------------------------------------------------
Function CrearDynaset(TablaName As String, Master As Integer) As Integer
    Dim SQLcad As String

    On Error GoTo CrearDynasetErr
    SQLcad = "SELECT * FROM " + TablaName
    Set dynListin = MAESTRO.CreateDynaset(SQLcad)
    
    'exito
    CrearDynaset = True
    GoTo CrearDynasetEnd

CrearDynasetErr:
  CrearDynaset = False
  Resume CrearDynasetEnd

CrearDynasetEnd:

End Function

'----------------------------------------------------------
'-- Texto=FNombre("ROA, KENNER")
'-- Texto="Kenner"
'----------------------------------------------------------
Function FNombre(ApeNom As String) As String
    Mayuscula = 1
    FNombre = ""
    For Letra = Len(ApeNom) To 1 Step -1
        If Mid(ApeNom, Letra, 1) = "," Then
            Exit For
        End If
        If Letra > 2 Then
            If Mid(ApeNom, Letra - 1, 1) = " " Then
                Mayuscula = 0
            End If
        End If
        If Mayuscula = 0 Then
            FNombre = UCase(Mid(ApeNom, Letra, 1)) + FNombre
            Mayuscula = 1
        Else
            FNombre = LCase(Mid(ApeNom, Letra, 1)) + FNombre
        End If
    Next Letra
    FNombre = Trim(FNombre)
End Function

'----------------------------------------------------------
'-- Texto=FApellidos("ROA, KENNER")
'-- Texto="Roa"
'----------------------------------------------------------
Function FApellidos(ApeNom As String) As String
    Mayuscula = 0
    FApellidos = ""
    For Letra = 1 To Len(ApeNom)
        If Mid(ApeNom, Letra, 1) = "," Then
            Exit For
        End If
        If Letra > 2 Then
            If Mid(ApeNom, Letra - 1, 1) = " " Then
                Mayuscula = 0
            End If
        End If
        If Mayuscula = 0 Then
            FApellidos = FApellidos + UCase(Mid(ApeNom, Letra, 1))
            Mayuscula = 1
        Else
            FApellidos = FApellidos + LCase(Mid(ApeNom, Letra, 1))
        End If
    Next Letra
End Function

