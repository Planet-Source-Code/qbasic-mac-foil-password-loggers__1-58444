VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo of FPLv1"
   ClientHeight    =   2865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Instructions"
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   5295
      Begin VB.Label Label1 
         Caption         =   "Enter password            (Just press Enter if none)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1800
         TabIndex        =   2
         Top             =   1800
         Width           =   3375
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdRetry 
      Caption         =   "Retry"
      Height          =   495
      Left            =   1680
      TabIndex        =   4
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Enter password            (Just press Enter if none)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   3375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim GotPass As Boolean

Private Sub cmdExit_Click()
End
End Sub

Private Sub cmdRetry_Click()
GotPass = False
Call Form_Activate
cmdExit.SetFocus
End Sub

Private Sub Form_Load()
Frame1.Top = 120
End Sub

Private Sub Form_Activate()
If GotPass Then Exit Sub
GotPass = True
Frame1.Visible = True
FPLv11.Show vbModal
Dim M As String
M = "Password Entered was" + String$(6, " ") + "-->"
Label2.Caption = M + PwdEntry + "<--"
Frame1.Visible = False
End Sub

