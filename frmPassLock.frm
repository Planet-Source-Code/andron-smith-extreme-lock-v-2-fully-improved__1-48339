VERSION 5.00
Begin VB.Form frmPassLock 
   BackColor       =   &H00C00000&
   BorderStyle     =   0  'None
   ClientHeight    =   3090
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmPassLock.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FF0000&
      Caption         =   "Frame1"
      Height          =   3090
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtPassword 
         BackColor       =   &H00FF8080&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   2
         Text            =   "Password"
         Top             =   1560
         Width           =   2175
      End
      Begin VB.TextBox txtUser 
         BackColor       =   &H00FF8080&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   "Username"
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   1320
         TabIndex        =   3
         Top             =   720
         Width           =   45
      End
   End
   Begin VB.Timer timPause 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   720
      Top             =   1200
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   840
      Top             =   360
   End
End
Attribute VB_Name = "frmPassLock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
#Const ProgramVersion = "1.0.1010"

Private inMessagebox

Private Sub DisplayMessage(ByVal Msg As String)
    
    lblStatus.Caption = Msg
    lblStatus.Left = Frame1.Width / 2 - lblStatus.Width / 2
    timPause.Enabled = True
End Sub

Private Sub Redraw_Form()
Me.Height = Screen.Height
Me.Width = Screen.Width
Me.Top = 0
Me.Left = 0
Frame1.Caption = App.Title
Frame1.Top = Me.ScaleHeight / 2 - Frame1.Height / 2
Frame1.Left = Me.ScaleWidth / 2 - Frame1.Width / 2

ActivateForm

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 95 Then KeyCode = 0




End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
txtPassword.SetFocus

End Sub

Private Sub Form_Load()
If txtUser.Text = "" Then txtUser.Locked = False


lngI = SetFocuses(Me.hWnd)



End Sub

Private Sub Form_Resize()
Redraw_Form


End Sub

Private Sub Timer2_Timer()

End Sub

Private Function ReturnPassCLue()
   Dim PassClue
      PassClue = Left$(PasswordInMemory.strPassword, 1)
      For i = 2 To Len(PasswordInMemory.strPassword)
        PassClue = PassClue & "*"
      Next i
      ReturnPassCLue = PassClue
      
End Function

Private Sub txtPassword_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyReturn Then
   If IsPasswordCorrect(txtPassword.Text, txtUser.Text) <> yesitdoes Then
    DisplayMessage "Username/Password is incorrect: (" & ReturnPassCLue() & ")"
      Else
      End
      
   End If
End If

    
       


End Sub

Private Sub txtUser_GotFocus()
If txtUser.Locked = True Then txtPassword.SetFocus

End Sub
