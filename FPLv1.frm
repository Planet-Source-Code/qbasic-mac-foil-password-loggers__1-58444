VERSION 5.00
Begin VB.Form FPLv11 
   Caption         =   "Enter Password"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7890
   ControlBox      =   0   'False
   LinkTopic       =   "FPLv11"
   ScaleHeight     =   5850
   ScaleWidth      =   7890
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAbort 
      Caption         =   "Abort"
      Height          =   495
      Left            =   1680
      TabIndex        =   11
      Top             =   3840
      Width           =   1215
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   495
      Left            =   240
      TabIndex        =   10
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox txt12 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   7080
      MultiLine       =   -1  'True
      TabIndex        =   9
      Top             =   3720
      Width           =   675
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   3960
      Top             =   4920
   End
   Begin VB.TextBox txtShow 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   2895
   End
   Begin VB.TextBox txtL2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   2760
      Width           =   4335
   End
   Begin VB.TextBox txtU2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   4335
   End
   Begin VB.CommandButton cmdFlip 
      Caption         =   "Secure It"
      Height          =   345
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   2895
   End
   Begin VB.TextBox txtIn 
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   4680
      Width           =   2895
   End
   Begin VB.TextBox txtN 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   3720
      Width           =   3375
   End
   Begin VB.TextBox txtU1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
   Begin VB.TextBox txtL1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   3480
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   1920
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "Use translation tables at the right to enter password above"
      Height          =   855
      Left            =   1560
      TabIndex        =   7
      Top             =   2520
      Width           =   1695
   End
End
Attribute VB_Name = "FPLv11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const xHi = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
Const xLo = "abcdefghijklmnopqrstuvwxyz"
Const xNo = "+-23456789"
Dim x As String * 62
Dim y As String * 62
Dim c(62) As String * 1
Dim Password As String
Const ModeS = "Change back to unsecure mode"
Const ModeU = "Use secure password entry method"
Dim ShowCursor As String * 1
Dim ShowThis As String

Private Sub cmdAbort_Click()
Call ResetMe
PwdEntry = ""
Unload Me
End Sub

Private Sub ResetMe()
txtIn.Text = ""
Password = ""
ShowThis = ""
End Sub

Private Sub cmdFlip_Click()
If cmdFlip.Caption = ModeU Then
  FPLv11.Height = 5040
  FPLv11.Width = 7980
  cmdFlip.Caption = ModeS
  Call ShuffleShow
Else
  Call NoShuffle
End If
Call ResetMe
txtIn.SetFocus
End Sub

Private Sub cmdHelp_Click()
Dim dbl As String
Dim M As String
M = M + vbLf + "Let's say your password is 'Apple31'"
M = M + vbLf
M = M + vbLf + "To enter the upper case 'A', look at the"
M = M + vbLf + "translation table and note that under the"
M = M + vbLf + "'A' is '" + Left$(y, 1) + "'. Enter '"
M = M + Left$(y, 1) + "' instead!"
M = M + vbLf
M = M + vbLf + "Similarly, say your password is '10MeN'"
M = M + vbLf
M = M + vbLf + "To enter the numeric '1', look at the"
M = M + vbLf + "translation table and note that under the"
M = M + vbLf + "'1' is '" + Mid$(y, 54, 1) + "'. Enter '"
M = M + Mid$(y, 54, 1) + "' instead!"
M = M + vbLf

M = M + vbLf
M = M + vbLf + "Finally, say your password is '"
M = M + Left$(xNo, 1) + "Jack3'"
M = M + vbLf + "To enter that special symbol, look at the"
M = M + vbLf + "small translation table and find that you"
M = M + vbLf + "would enter '0' (zero) instead"
M = M + vbLf
M = M + vbLf + "The two special characters are encoded as"
M = M + vbLf + "0 and 1 in order to not use 0 or 1 in the"
M = M + vbLf + "translation table, as they look like the"
M = M + vbLf + "alphabetic characters O and l"
M = M + vbLf
M = M + vbLf + "If you have a symbol other than the two"
M = M + vbLf + "special characters shown, just enter the"
M = M + vbLf + "symbol. It will be accepted and not changed."
M = M + vbLf
M = M + vbLf + "Note: Translations change after every"
M = M + vbLf + "character, so be careful to do one at a time"
MsgBox M, , "How to enter secure passwords"
txtIn.SetFocus
End Sub

Private Sub Form_Load()
txt12.Text = Expand(Left$(xNo, 2) + "01")
x = xHi + xLo + xNo
Dim i As Integer
For i = 1 To 62: c(i) = Mid$(x, i, 1): Next i
Call NoShuffle: ' Default is not so much security
End Sub

Private Sub NoShuffle()
FPLv11.Height = 1530
FPLv11.Width = 3300
txtU1.Text = "Enter password"
txtU2.Text = "(Not secured)"
txtL1.Text = ""
txtL2.Text = ""
txtN.Text = ""
cmdFlip.Caption = ModeU
End Sub

Private Function Expand(x As String) As String
Dim y As String
Dim i As Integer
For i = 1 To Len(x)
  y = y + Mid$(x, i, 1) + " "
Next i
Expand = y
End Function

Private Sub ShuffleShow()
Dim i As Integer
Dim Sort As Integer
Dim r As Integer
Randomize Timer
For Sort = 1 To 5
  For i = 1 To 62
    r = 1 + Int(Rnd * 62)
    Dim w As String * 1
    w = c(i)
    c(i) = c(r)
    c(r) = w
  Next i
Next Sort
Dim V As String * 62
For i = 1 To 62: Mid$(y, i, 1) = c(i): Next i
txtU1.Text = Expand(Mid$(x, 1, 13)) + Expand(Mid$(y, 1, 13))
txtU2.Text = Expand(Mid$(x, 14, 13)) + Expand(Mid$(y, 14, 13))
txtL1.Text = Expand(Mid$(x, 27, 13)) + Expand(Mid$(y, 27, 13))
txtL2.Text = Expand(Mid$(x, 40, 13)) + Expand(Mid$(y, 40, 13))
txtN.Text = Expand("0123456789") + Expand(Right$(y, 10))
End Sub

Private Function Less1(l As String) As String
Less1 = Left$(l, Len(l) - 1)
End Function

Private Sub Timer1_Timer()
If Me.Visible = False Then Exit Sub
Timer1.Enabled = False
If InStr(Password, Chr$(13)) > 0 Then
  Password = ""
  Timer1.Interval = 1
  ShowThis = ""
End If
If Timer1.Interval = 1 Then
  Timer1.Interval = 600
  ShowCursor = " "
End If
If ShowCursor = " " Then
  ShowCursor = "_"
Else
  ShowCursor = " "
End If
txtShow.Text = ShowThis + ShowCursor
Timer1.Enabled = True
End Sub

Private Sub txtIn_KeyPress(KeyAscii As Integer)
Timer1.Enabled = False
GoSub DoIt
Call ShuffleShow
Timer1.Interval = 1
Timer1.Enabled = True
Exit Sub

DoIt:
Dim q As Integer
Dim c As String * 1
If KeyAscii = 13 Then PwdEntry = Password: Unload Me
If KeyAscii = 27 Then Call ResetMe: Return
If KeyAscii = 8 Then
  If Len(Password) > 0 Then
    Password = Less1(Password)
    ShowThis = Less1(ShowThis)
  End If
  Return
End If
If cmdFlip.Caption = ModeU Then
  Password = Password + Chr$(KeyAscii)
Else
  q = InStr(y, Chr$(KeyAscii))
  If q = 0 Then
    If KeyAscii = 48 Then
      Password = Password + Left$(xNo, 1)
    ElseIf KeyAscii = 49 Then
      Password = Password + Mid$(xNo, 2, 1)
    Else
      Password = Password + Chr$(KeyAscii)
    End If
  Else
    c = Mid$(x, q, 1)
    If c = Left$(xNo, 1) Then
      Password = Password + "0"
    ElseIf c = Mid$(xNo, 2, 1) Then
      Password = Password + "1"
    Else
      Password = Password + Mid$(x, q, 1)
    End If
  End If
End If
ShowThis = String$(Len(Password), "*")
Return
End Sub
