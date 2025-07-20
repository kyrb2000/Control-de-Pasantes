VERSION 5.00
Begin VB.Form frmMSNPopup 
   BackColor       =   &H00D0FFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1695
   ClientLeft      =   8580
   ClientTop       =   11745
   ClientWidth     =   2775
   LinkTopic       =   "Form1"
   ScaleHeight     =   1695
   ScaleWidth      =   2775
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   0
      Top             =   360
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "¿ Hola ?"
      ForeColor       =   &H000000FF&
      Height          =   1755
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmMSNPopup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DirectionIsUp As Boolean ' Up is True, Down is False

Private Sub Form_Load()
  'Move it below the visible screen (and a little just in case)
  frmMSNPopup.Top = Screen.Height + 10
  
  'Move it to the far right of the visible screen (minus a little, just for esthetics)
  frmMSNPopup.Left = Screen.Width - (frmMSNPopup.Width + 100)
  
  'We're gonna move it up
  DirectionIsUp = True
End Sub

Private Sub Timer1_Timer()
  
  'Move at 10 millisecond intervals (100 times a second, 3 times what the eye can see)
  Timer1.Interval = 10
  
  ' If it's moving up
  If DirectionIsUp Then
    
    'Move it up 50 twips every 10 milliseconds
    frmMSNPopup.Top = frmMSNPopup.Top - 50
    
    'Move until the whole form is shown (minus 10 twips to make sure it still touches the bottom of the screen)
    If (frmMSNPopup.Top <= Screen.Height - (frmMSNPopup.Height - 10)) Then
      
      ' This specifies how long it will stay shown (Unmoving)
      Timer1.Interval = 3000
      
      ' We're gonna move it down next...
      DirectionIsUp = False
    End If
  
  Else
    'Move it down 50 twips every 10 milliseconds
    frmMSNPopup.Top = frmMSNPopup.Top + 50
    
    'Move until the whole form is shown (plus 10 twips to make sure it's hidden)
    If frmMSNPopup.Top >= Screen.Height + 10 Then
      Timer1.Enabled = False
      Unload Me
    End If
  End If
End Sub

