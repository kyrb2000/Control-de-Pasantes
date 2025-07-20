VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Adodc1 
   Caption         =   "Form1"
   ClientHeight    =   4125
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5250
   LinkTopic       =   "Form1"
   ScaleHeight     =   4125
   ScaleWidth      =   5250
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Carta de Postulación"
      Height          =   495
      Left            =   3000
      TabIndex        =   4
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Carta de Presentación"
      Height          =   495
      Left            =   3000
      TabIndex        =   3
      Top             =   1800
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      DataField       =   "ApeNom"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1200
      TabIndex        =   1
      Top             =   840
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      DataField       =   "Cedula"
      DataSource      =   "Adodc1"
      Height          =   285
      Left            =   1200
      TabIndex        =   0
      Text            =   "Cedula"
      Top             =   480
      Width           =   1095
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   240
      Top             =   2400
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   3
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "DSN=Pasantias"
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   "Pasantias"
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Alumnos"
      Caption         =   "Adodc1"
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
Attribute VB_Name = "Adodc1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Shell "C:\Archivos de programa\Microsoft Office\Office\winword.exe C:\Misdoc~1\Pasant~1\Postulación.doc"
End Sub

Private Sub Command2_Click()
    Shell "C:\Archivos de programa\Microsoft Office\Office\winword.exe C:\Misdoc~1\Pasant~1\Presentación.doc"
End Sub

Private Sub Text3_LostFocus()
    Adodc1.Recordset.MoveFirst
    While Not Adodc1.Recordset.EOF
        If Adodc1.Recordset.Fields(0) = Text3 Then
            Exit Sub
        End If
        Adodc1.Recordset.MoveNext
    Wend
End Sub
