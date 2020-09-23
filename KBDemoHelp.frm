VERSION 5.00
Begin VB.Form KBDemoHelp 
   Caption         =   "clsKeyboardPicture and ClsVirtualScoreSheet Help"
   ClientHeight    =   3765
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7545
   LinkTopic       =   "Form1"
   ScaleHeight     =   3765
   ScaleWidth      =   7545
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   3240
      TabIndex        =   1
      Top             =   3240
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "KBDemoHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Command1_Click()

    Me.Hide

End Sub

Private Sub Form_Load()

  Dim Fnum As Integer

    On Error GoTo Filemissing
    Fnum = FreeFile
    Open App.Path & "\" & "keyboardhelp.txt" For Input As Fnum
    Text1.Text = Input(LOF(Fnum), Fnum)
    Close

Exit Sub

Filemissing:
    MsgBox "The helpfile keyboardhelp.txt is missing or in wrong directory. Please move it to the same directory as the program."

End Sub

Private Sub Form_Resize()

    Text1.Width = KBDemoHelp.ScaleWidth
    Command1.Top = KBDemoHelp.ScaleHeight - Command1.Height
    Command1.Left = KBDemoHelp.ScaleWidth - Command1.Width
    Text1.Height = KBDemoHelp.ScaleHeight - Command1.Height

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyCode = 0

End Sub

':) Ulli's VB Code Formatter V2.13.6 (27/09/2002 3:19:06 PM) 1 + 40 = 41 Lines
