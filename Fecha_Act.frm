VERSION 5.00
Begin VB.Form Fecha_Act 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Fechero"
   ClientHeight    =   1290
   ClientLeft      =   3180
   ClientTop       =   3105
   ClientWidth     =   4710
   Icon            =   "Fecha_Act.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1290
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
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
      Height          =   1215
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
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
         Left            =   3300
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtFecha 
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
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "&Aceptar"
         Default         =   -1  'True
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
         Left            =   1920
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox lblYear 
         Alignment       =   2  'Center
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
         Height          =   420
         Left            =   3480
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox lblDay 
         Alignment       =   1  'Right Justify
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
         Left            =   240
         TabIndex        =   0
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox lblMonth 
         Alignment       =   2  'Center
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
         Height          =   420
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   2055
      End
      Begin VB.VScrollBar vsbMonth 
         Height          =   375
         Left            =   3120
         Max             =   1
         Min             =   12
         TabIndex        =   3
         Top             =   240
         Value           =   1
         Width           =   255
      End
      Begin VB.VScrollBar vsbDay 
         Height          =   375
         Left            =   720
         Max             =   1
         Min             =   31
         TabIndex        =   1
         Top             =   240
         Value           =   1
         Width           =   255
      End
      Begin VB.VScrollBar vsbYear 
         Height          =   375
         Left            =   4320
         Max             =   1800
         Min             =   2100
         TabIndex        =   5
         Top             =   240
         Value           =   2001
         Width           =   255
      End
   End
End
Attribute VB_Name = "Fecha_Act"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------
'-- Autor : Kenner Y. Roa B. Telefono (0426)-832.75.74
'-- E-mail : kyrb2000@gmail.com / kyrb2000@hotmail.com
'-- Fecha : 22/05/2002 (Caracas / Venezuela)
'-----------------------------------------------------------
Option Explicit
Dim Months(12) As String
Dim Days(12) As Integer
Public XFecha_Actual, XDia_Min, XDia_Max As String      'Variables para Fecha_Act.FRM

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Call cmdCancelar_Click
End Sub

Private Sub cmdCancelar_Click()
    txtFecha = "-"
    Me.Hide
End Sub

Private Sub cmdUpdate_Click()
    Dim Is_Leap As Integer
    Dim Msg As String, Age As Integer, Pronoun As String
    Dim M As Integer, D As Integer, Y As Integer
    
    'Check for leap year and if February is current month
    If vsbMonth.Value = 2 And vsbYear.Value Mod 4 = 0 And vsbYear.Value Mod 100 <> 0 Then
        Is_Leap = 1
    Else
        Is_Leap = 0
    End If

    'Check to make sure current day doesn't exceed number of days in month
    If vsbDay.Value > Days(vsbMonth.Value) + Is_Leap Then
        MsgBox "El Mes de " + Trim(Months(vsbMonth.Value)) + " Tiene " + Trim(Str(Days(vsbMonth.Value)) + Is_Leap) + " Dias", vbOKOnly + vbCritical, "La Fecha no es v�lida"
        lblDay.Text = Trim(Str(Days(vsbMonth.Value) + Is_Leap))
        Exit Sub
    End If

    txtFecha = vsbDay & "/" & vsbMonth.Value & "/" & vsbYear
    txtFecha = Format(txtFecha, Tipo_Fecha)
    
    Select Case DateDiff("d", Date, txtFecha)
    Case Is < Val(XDia_Min) '    If DateDiff("d", Date, txtFecha) < 0 Then
        MsgBox "Es Menor de la Fecha Actual " & DateDiff("d", Date, txtFecha) & " dias " & Format(txtFecha, Tipo_Fecha), vbOKOnly + vbCritical, "La Fecha no es v�lida"
        txtFecha = Format(XFecha_Actual, Tipo_Fecha)
        Call Actual_de_Fecha
        Exit Sub
    Case Is > Val(XDia_Max) '    If DateDiff("d", Date, txtFecha) > 90 Then
        MsgBox "Es Mayor de " & XDia_Max & " Dias Tiene " & DateDiff("d", Date, txtFecha) & " dias " & Format(txtFecha, Tipo_Fecha), vbOKOnly + vbCritical, "La Fecha no es v�lida"
        txtFecha = Format(XFecha_Actual, Tipo_Fecha)
        Call Actual_de_Fecha
        Exit Sub
    End Select
    Me.Hide
End Sub

Private Sub Form_Load()
    'Set arrays for dates and initialize labels
    Months(1) = "Enero": Days(1) = 31
    Months(2) = "Febrero": Days(2) = 28
    Months(3) = "Marzo": Days(3) = 31
    Months(4) = "Abril": Days(4) = 30
    Months(5) = "Mayo": Days(5) = 31
    Months(6) = "Junio": Days(6) = 30
    Months(7) = "Julio": Days(7) = 31
    Months(8) = "Agosto": Days(8) = 31
    Months(9) = "Septiembre": Days(9) = 30
    Months(10) = "Octubre": Days(10) = 31
    Months(11) = "Noviembre": Days(11) = 30
    Months(12) = "Diciembre": Days(12) = 31
    Call Actual_de_Fecha
End Sub

Function Actual_de_Fecha()
    If XFecha_Actual = "" Then
        XFecha_Actual = Format(Date, Tipo_Fecha)
    End If
    vsbMonth.Value = Val(Format(XFecha_Actual, "mm"))
    vsbDay.Value = Val(Format(XFecha_Actual, "dd"))
    vsbYear.Value = Val(Format(XFecha_Actual, "yyyy"))
    lblMonth.Text = Trim(Months(vsbMonth.Value))
    lblDay.Text = Trim(Str(vsbDay.Value))
    lblYear.Text = Trim(Str(vsbYear.Value))
End Function

Private Sub lblDay_Change()
    Select Case Val(lblDay)
    Case Is > 31
        lblDay = "31"
    Case Is < 1
        lblDay = "1"
    End Select
    vsbDay.Value = Val(lblDay)
    txtFecha = vsbDay & "/" & vsbMonth.Value & "/" & vsbYear
    txtFecha = Format(txtFecha, Tipo_Fecha)
End Sub

Private Sub lblDay_Click()
    lblDay.SelStart = 0
    lblDay.SelLength = Len(lblDay)
End Sub

Private Sub lblDay_KeyPress(KeyAscii As Integer)
    '== s�lo se dejan introducir letras
    Select Case KeyAscii
    Case Asc("1") To Asc("0"):
'      KeyAscii = KeyAscii
    Case Asc("A") To Asc("Z"):
        KeyAscii = 0
    Case Asc("a") To Asc("z"):
        KeyAscii = 0
    Case 32: '-- Spc y Tab, capturados para que solo salga pulsando CR.
        KeyAscii = 0
    Case 13: '-- CR
    End Select
End Sub

Private Sub lblMonth_Change()
    txtFecha = vsbDay & "/" & vsbMonth.Value & "/" & vsbYear
    txtFecha = Format(txtFecha, Tipo_Fecha)
End Sub

Private Sub lblYear_Change()
    vsbYear.Value = Val(lblYear)
    txtFecha = vsbDay & "/" & vsbMonth.Value & "/" & vsbYear
    txtFecha = Format(txtFecha, Tipo_Fecha)
End Sub

Private Sub lblYear_KeyPress(KeyAscii As Integer)
    '== s�lo se dejan introducir letras
    Select Case KeyAscii
    Case Asc("1") To Asc("0"):
'      KeyAscii = KeyAscii
    Case Asc("A") To Asc("Z"):
        KeyAscii = 0
    Case Asc("a") To Asc("z"):
        KeyAscii = 0
    Case 32: '-- Spc y Tab, capturados para que solo salga pulsando CR.
        KeyAscii = 0
    Case 13: '-- CR
    End Select
End Sub

Private Sub vsbDay_Change()
    lblDay.Text = Trim(Str(vsbDay.Value))
End Sub

Private Sub vsbMonth_Change()
    lblMonth.Text = Trim(Months(vsbMonth.Value))
End Sub

Private Sub vsbYear_Change()
    lblYear.Text = Trim(Str(vsbYear.Value))
End Sub

