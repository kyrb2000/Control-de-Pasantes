VERSION 5.00
Begin VB.Form FrmResultados 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2790
   ClientLeft      =   1650
   ClientTop       =   1935
   ClientWidth     =   5535
   Icon            =   "FrmResultados.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2790
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Apreciación: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   2850
      TabIndex        =   8
      Top             =   1800
      Width           =   2355
      Begin VB.Label LblApreciacion 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   135
         TabIndex        =   12
         Top             =   390
         Width           =   2130
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Porcentajes: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1005
      Left            =   2820
      TabIndex        =   7
      Top             =   750
      Width           =   2370
      Begin VB.Label lblpor1 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1230
         TabIndex        =   16
         Top             =   210
         Width           =   930
      End
      Begin VB.Label lblpor2 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1230
         TabIndex        =   15
         Top             =   570
         Width           =   930
      End
      Begin VB.Label Label6 
         Caption         =   "Incorrectas:"
         Height          =   255
         Left            =   165
         TabIndex        =   11
         Top             =   615
         Width           =   780
      End
      Begin VB.Label Label5 
         Caption         =   "Correctas:"
         Height          =   195
         Left            =   165
         TabIndex        =   10
         Top             =   285
         Width           =   915
      End
   End
   Begin VB.Frame Frame2 
      Height          =   705
      Left            =   240
      TabIndex        =   1
      Top             =   15
      Width           =   4965
      Begin VB.Label Lblnotafinal 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   3870
         TabIndex        =   9
         Top             =   180
         Width           =   885
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   75
         Picture         =   "FrmResultados.frx":0442
         Top             =   120
         Width           =   480
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "RESULTADO:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   630
         TabIndex        =   2
         Top             =   30
         Width           =   3450
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preguntas: "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1875
      Left            =   285
      TabIndex        =   0
      Top             =   765
      Width           =   2400
      Begin VB.Label lblincorrectas 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1290
         TabIndex        =   14
         Top             =   1290
         Width           =   930
      End
      Begin VB.Label lblcorrectas 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1290
         TabIndex        =   13
         Top             =   825
         Width           =   930
      End
      Begin VB.Label Lblnumerop 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1290
         TabIndex        =   6
         Top             =   360
         Width           =   930
      End
      Begin VB.Label Label4 
         Caption         =   "Respuestas Incorrectas:"
         Height          =   480
         Left            =   150
         TabIndex        =   5
         Top             =   1185
         Width           =   930
      End
      Begin VB.Label Label3 
         Caption         =   "Respuestas Correctas:"
         Height          =   480
         Left            =   150
         TabIndex        =   4
         Top             =   690
         Width           =   930
      End
      Begin VB.Label Label2 
         Caption         =   "Nª Preguntas:"
         Height          =   315
         Left            =   150
         TabIndex        =   3
         Top             =   360
         Width           =   1155
      End
   End
End
Attribute VB_Name = "FrmResultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Lblnumerop.Caption = 10
    lblcorrectas.Caption = 9
    lblincorrectas.Caption = 1
    lblpor1 = lblcorrectas * 100 / 10
    lblpor2 = lblincorrectas * 100 / 10
    LblApreciacion = "Muy Bien"
End Sub
