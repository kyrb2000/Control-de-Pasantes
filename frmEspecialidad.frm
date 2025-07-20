VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmEspecialidad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Especialidad"
   ClientHeight    =   2760
   ClientLeft      =   2250
   ClientTop       =   1500
   ClientWidth     =   8895
   Icon            =   "frmEspecialidad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2760
   ScaleWidth      =   8895
   Begin VB.Frame Frame3 
      Height          =   1095
      Left            =   0
      TabIndex        =   5
      Top             =   1320
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
         Picture         =   "frmEspecialidad.frx":0442
         Style           =   1  'Graphical
         TabIndex        =   6
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
         Picture         =   "frmEspecialidad.frx":0869
         Style           =   1  'Graphical
         TabIndex        =   7
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
         Picture         =   "frmEspecialidad.frx":0C84
         Style           =   1  'Graphical
         TabIndex        =   8
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
         Picture         =   "frmEspecialidad.frx":10AC
         Style           =   1  'Graphical
         TabIndex        =   9
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
         Picture         =   "frmEspecialidad.frx":14E9
         Style           =   1  'Graphical
         TabIndex        =   10
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
         Picture         =   "frmEspecialidad.frx":1952
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   150
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8895
      Begin VB.TextBox txtFields 
         DataField       =   "Codigo"
         DataSource      =   "mdbEspecialidad"
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
         Index           =   0
         Left            =   1800
         TabIndex        =   2
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtFields 
         DataField       =   "Descripcion"
         DataSource      =   "mdbEspecialidad"
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
         Index           =   1
         Left            =   1800
         TabIndex        =   1
         Top             =   795
         Width           =   6975
      End
      Begin VB.Label lblLabels 
         Caption         =   "Codigo:"
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
         TabIndex        =   4
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblLabels 
         Caption         =   "Descripcion:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   795
         Width           =   1815
      End
   End
   Begin MSAdodcLib.Adodc mdbEspecialidad 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      Top             =   2430
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   582
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
      Connect         =   "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb;"
      OLEDBString     =   "PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=C:\Archivos de programa\Sistema de Pasantías\Datos_Alumnos.mdb;"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select Codigo,Descripcion from Especialidad Order by Codigo"
      Caption         =   " "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "frmEspecialidad"
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
    Frame1.Enabled = True
    mdbEspecialidad.Recordset.AddNew
    '--------- Botones ------------
    cmdUpdate.Caption = "&Actualizar"
    cmdUpdate.Enabled = True
    cmdAdd.Enabled = False
    cmdDelete.Enabled = False
    cmdBuscar.Enabled = False
    'cmdImprimir.Enabled = False
    cmdClose.Caption = "&Cancelar"
    mdbEspecialidad.Enabled = False
    '------------------------------'
End Sub

Private Sub cmdDelete_Click()
    Dim Respuesta
    If Not mdbEspecialidad.Recordset.EOF Then
        'esto puede producir un error si elimina el último
        'registro o el único registro del recordset
        Respuesta = MsgBox("Desea Eliminar el Registro", vbYesNo + vbCritical, "Eliminación de Datos")
        If Respuesta = vbYes Then
            mdbEspecialidad.Recordset.Delete
            mdbEspecialidad.Recordset.MoveFirst
        End If
    End If
End Sub

Private Sub cmdBuscar_Click()
    Dim Buscar As String
    If (mdbEspecialidad.Recordset.AbsolutePosition + 1) = 0 Then
        Exit Sub
    End If
    Buscar = InputBox("Introduzca El Codigo :", "Busqueda de Datos")
    If Buscar <> "" Then
        Buscar = "[Codigo]" & "=" & "'" & Buscar & "'"
        mdbEspecialidad.Recordset.Find Buscar
        If mdbEspecialidad.Recordset.EOF Then
            MsgBox "Codigo No Existe", vbCritical, "Buscar Codigo"
        End If
    End If
End Sub

Private Sub cmdImprimir_Click()
'    Dim fReportes As New frmReportes
'    fReportes.Caption = "Reportes de Alumnos"
'    fReportes.Tipo_Reporte = "ALU"
'    Call centrarForm(fReportes)
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
        mdbEspecialidad.Recordset.Update '.UpdateRecord
'        mdbEspecialidad.Recordset.Bookmark = mdbEspecialidad.Recordset.LastModified
        '--------- Botones ------------
        cmdUpdate.Caption = "&Modificar"
        cmdAdd.Enabled = True
        cmdDelete.Enabled = True
        cmdBuscar.Enabled = True
        'cmdImprimir.Enabled = True
        cmdClose.Caption = "&Cerrar"
        mdbEspecialidad.Enabled = True
        '------------------------------'
    Else
        If Not mdbEspecialidad.Recordset.EOF Then
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
            'mdbEspecialidad.Recordset.Edit
            '--------- Botones ------------
            cmdUpdate.Caption = "&Actualizar"
            cmdAdd.Enabled = False
            cmdDelete.Enabled = False
            cmdBuscar.Enabled = False
            'cmdImprimir.Enabled = False
            cmdClose.Caption = "&Cancelar"
            mdbEspecialidad.Enabled = False
            '------------------------------'
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
        'cmdImprimir.Enabled = True
        cmdClose.Caption = "&Cerrar"
        mdbEspecialidad.Enabled = True
        '------------------------------'
        mdbEspecialidad.Refresh
        mdbEspecialidad.Recordset.MoveFirst
    End If
End Sub
