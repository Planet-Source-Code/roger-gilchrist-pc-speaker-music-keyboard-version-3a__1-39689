VERSION 5.00
Begin VB.Form FrmKBDemo 
   Caption         =   "Key Board Class Demo"
   ClientHeight    =   11520
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   10905
   LinkTopic       =   "Form1"
   ScaleHeight     =   11520
   ScaleWidth      =   10905
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Play RTTTL"
      Enabled         =   0   'False
      Height          =   375
      Index           =   4
      Left            =   9360
      TabIndex        =   20
      Top             =   300
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Index           =   2
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   19
      Top             =   0
      Width           =   9375
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Auto-Repair"
      Height          =   195
      Left            =   9360
      TabIndex        =   18
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CheckBox Check3 
      Alignment       =   1  'Right Justify
      Caption         =   "Nokia Composer keys only"
      Height          =   255
      Left            =   5040
      TabIndex        =   17
      Top             =   4920
      Width           =   2175
   End
   Begin VB.CheckBox Check2 
      Alignment       =   1  'Right Justify
      Caption         =   "Show Nokia Composer Safe Range"
      Height          =   255
      Left            =   5160
      TabIndex        =   16
      Top             =   5880
      Width           =   2895
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   15
      Top             =   3120
      Width           =   7695
   End
   Begin VB.CheckBox Check1 
      Alignment       =   1  'Right Justify
      Caption         =   "Show Bass Stave(F clef)"
      Height          =   255
      Left            =   8640
      TabIndex        =   14
      Top             =   5880
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000014&
      Height          =   5175
      Left            =   120
      ScaleHeight     =   5115
      ScaleWidth      =   10635
      TabIndex        =   13
      Top             =   6240
      Width           =   10695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Look"
      Height          =   375
      Index           =   6
      Left            =   9720
      TabIndex        =   12
      ToolTipText     =   "Save Random colour set to file 'KeyBoardLooks.txt'"
      Top             =   4890
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "clsKeyboardPicture"
      Height          =   1215
      Left            =   120
      TabIndex        =   10
      Top             =   3600
      Width           =   10455
      Begin VB.PictureBox Picture1 
         Height          =   855
         Left            =   120
         ScaleHeight     =   795
         ScaleWidth      =   10155
         TabIndex        =   11
         Top             =   240
         Width           =   10215
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Delete Last"
      Enabled         =   0   'False
      Height          =   375
      Index           =   3
      Left            =   9360
      TabIndex        =   7
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ComboBox cmbTempo 
      Height          =   315
      Left            =   9360
      TabIndex        =   6
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear Texts"
      Enabled         =   0   'False
      Height          =   375
      Index           =   2
      Left            =   8040
      TabIndex        =   5
      Top             =   3120
      Width           =   1095
   End
   Begin VB.ComboBox cmbLooks 
      Height          =   315
      Left            =   8520
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   4920
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play Basica"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   9360
      TabIndex        =   3
      Top             =   2400
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Index           =   1
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2160
      Width           =   9375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Play Nokia"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   9360
      TabIndex        =   1
      Top             =   1680
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Index           =   0
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   9375
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"FrmKBDemo.frx":0000
      Height          =   1250
      Left            =   120
      TabIndex        =   9
      Top             =   4920
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   "KeyBoard Look"
      Height          =   255
      Left            =   7320
      TabIndex        =   8
      Top             =   4950
      Width           =   1215
   End
   Begin VB.Menu mnuBasica 
      Caption         =   "BasicaTricks"
      Begin VB.Menu mnuWarning 
         Caption         =   "WARNING NO ERROR CHECKING"
      End
      Begin VB.Menu mnuBasicaOpt 
         Caption         =   "Octave Up        Max=10"
         Index           =   0
      End
      Begin VB.Menu mnuBasicaOpt 
         Caption         =   "Octave Down   Min=0"
         Index           =   1
      End
      Begin VB.Menu mnuBasicaOpt 
         Caption         =   "Note Up           A>B..... G>A"
         Index           =   2
      End
      Begin VB.Menu mnuBasicaOpt 
         Caption         =   "Note Down      A>G......G>F"
         Index           =   3
      End
      Begin VB.Menu mnuBasicaOpt 
         Caption         =   "Random Music"
         Index           =   4
      End
      Begin VB.Menu mnuBasicaOpt 
         Caption         =   "RealTranspose  (assumes Key of C)"
         Index           =   5
         Begin VB.Menu MnuTrans 
            Caption         =   " > D"
            Index           =   0
         End
         Begin VB.Menu MnuTrans 
            Caption         =   " > G-"
            Index           =   1
         End
         Begin VB.Menu MnuTrans 
            Caption         =   " > B-"
            Index           =   2
         End
         Begin VB.Menu MnuTrans 
            Caption         =   " > E-"
            Index           =   3
         End
         Begin VB.Menu MnuTrans 
            Caption         =   " > G#"
            Index           =   4
         End
         Begin VB.Menu MnuTrans 
            Caption         =   " > C#"
            Index           =   5
         End
         Begin VB.Menu MnuTrans 
            Caption         =   " > G-"
            Index           =   6
         End
         Begin VB.Menu MnuTrans 
            Caption         =   " > B"
            Index           =   7
         End
         Begin VB.Menu MnuTrans 
            Caption         =   " > E"
            Index           =   8
         End
         Begin VB.Menu MnuTrans 
            Caption         =   " > A"
            Index           =   9
         End
         Begin VB.Menu MnuTrans 
            Caption         =   " > D2"
            Index           =   10
         End
         Begin VB.Menu MnuTrans 
            Caption         =   " >G-2"
            Index           =   11
         End
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuhelpopt 
         Caption         =   "Help"
         Index           =   0
      End
      Begin VB.Menu mnuhelpopt 
         Caption         =   "About"
         Index           =   1
         Begin VB.Menu mnuaboutopt 
            Caption         =   "clsKeyBoardPicture"
            Index           =   0
         End
         Begin VB.Menu mnuaboutopt 
            Caption         =   "ClsVirtualScoreSheet"
            Index           =   1
         End
      End
   End
End
Attribute VB_Name = "FrmKBDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright 2002 Roger Gilchrist
'email: rojagilkrist@hotmail.com

'This is just a quick and dirty demo of ClsKeyBoardPicture
'Thanks to Nokia Ringtone Player by Ovidiu Daniel Diaconescu for the original inspiration and the
'Data in the NoteArray and some of the code to make noise


Option Explicit
Public Enum LimitType
    IntegerLimit
    LongLimit
    SingleLimit
    DoubleLimit
End Enum
Rem Mark Off
'Stops Code formatter complaining about these
#If False Then 'Enforce Case For Enums (does not compile but fools IDE)
Dim IntegerLimit
Dim LongLimit
Dim SingleLimit
Dim DoubleLimit
#End If  'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1
Rem Mark On

Private KB As New ClsKeyBoardPicture
Private VSS As New ClsVirtualScoreSheet
Private Declare Sub InitCommonControls Lib "comctl32" ()

Private Sub Check1_Click()

    VSS.ShowBassStave = Check1.Value = vbChecked
    VSS.ClearScore

End Sub

Private Sub Check2_Click()

    VSS.ShowNokiaRange = Check2.Value = vbChecked

End Sub

Private Sub Check3_Click()

    KB.ShowNokiaSafeOctaves = Check3.Value = vbChecked
    KB.Resize

End Sub

Private Sub Check4_Click()

    KB.NokiaAutoRepair = Check4.Value = vbChecked

End Sub

Private Sub cmbLooks_Click()

    KB.KeyBoardLook cmbLooks.ListIndex
    Command1(6).Enabled = cmbLooks.Text = "Random"

End Sub

Private Sub cmbTempo_Change()

    KB.Tempo = cmbTempo.Text

End Sub

Private Sub cmbTempo_Click()

    KB.Tempo = cmbTempo.Text

End Sub

Private Sub Combo1_Click()

  Dim tmp As String

    tmp = Combo1.Text
    Select Case Left$(tmp, 1)
      Case "N"
        tmp = Mid$(tmp, 7)
        KB.RTTTLTitle = Left$(tmp, InStr(tmp, "::Tempo") - 1)
        '"Nokia:Name::Tempo125:::
        tmp$ = Mid$(tmp, InStr(tmp, "::Tempo") + 8)
        cmbTempo.Text = Trim$(Left$(tmp$, InStr(tmp$, ":::") - 1))
        Text1(0).Text = UCase$(Mid$(tmp$, InStr(tmp$, ":::") + 3))
        Text1(1).Text = ""
        Text1(2).Text = ""
      Case "B"

        '"Basica:We Shall Overcome::T112 L4
        tmp = Mid$(tmp, InStr(tmp, ":") + 1)
        KB.RTTTLTitle = Left$(tmp, InStr(tmp, "::") - 1)
        Text1(1).Text = UCase$(Mid$(tmp, InStr(tmp, "::") + 2))
        Text1(0).Text = ""
        Text1(2).Text = ""
      Case "R"
        '"RTTL:VanessaMae:
        '"RTTL:Tubular:d=4,o=5,b=285:
        Text1(2).Text = Mid$(tmp, InStr(tmp, "RTTL:") + 5)
        KB.RTTTLTitle = Left$(Text1(2).Text, InStr(Text1(2).Text, ":") - 1)
        Text1(0).Text = ""
        Text1(1).Text = ""
        cmbTempo.Text = KB.Tempo
    End Select
    VSS.ClearScore

End Sub

Private Sub WriteTexts()

    If Text1(1).Text <> KB.CleanTextOutPut(Basica) Then
        Text1(1).Text = KB.CleanTextOutPut(Basica)
    End If
    If Text1(0).Text <> KB.CleanTextOutPut(Nokia) Then 'don't show clean if it matches already
        Text1(0).Text = KB.CleanTextOutPut(Nokia)

    End If
    If Text1(2).Text <> KB.CleanTextOutPut(RTTTL) Then 'don't show clean if it matches already
        Text1(2).Text = KB.CleanTextOutPut(RTTTL)

    End If
    VSS.AddNoteArray Split(Trim$(KB.CleanTextOutPut(Basica)))

End Sub

Private Sub Command1_Click(Index As Integer)

    Select Case Index
      Case 0 'Play Nokia
        Select Case Command1(0).Caption
          Case "Abort"
            KB.Abort
            Command1(0).Caption = "Play Nokia"
          Case Else
            Command1(0).Caption = "Abort"
            VSS.ClearScore
            KB.NokiaRead Text1(0)
            WriteTexts
            Command1(0).Caption = "Play Nokia"
        End Select
      Case 1 'Play Basica
        Select Case Command1(1).Caption
          Case "Abort"
            KB.Abort
            Command1(1).Caption = "Play Basica"
          Case Else
            Command1(1).Caption = "Abort"
            VSS.ClearScore
            '            If InStr(Text1(1).Text, "T") = 0 Then
            '                Text1(1).Text = "T" & cmbTempo.Text & Text1(1).Text
            '            End If
            KB.BasicaRead Text1(1)
            WriteTexts
            cmbTempo = KB.Tempo
            Command1(1).Caption = "Play Basica"
        End Select
      Case 2 'Clear Texts
        Text1(0) = ""
        Text1(1) = ""
        Text1(2) = ""
        KB.RTTTLString = ""
        VSS.ClearScore

      Case 4 'Play RTTL
        Select Case Command1(4).Caption
          Case "Abort"
            KB.Abort
            Command1(4).Caption = "Play RTTTL"
          Case Else
            Command1(4).Caption = "Abort"
            VSS.ClearScore
            If KB.RTTTLValid(Text1(2)) Then
                KB.RTTTLString = Text1(2)
            End If
            KB.RTTTLRead KB.RTTTLString
            WriteTexts
            cmbTempo = KB.Tempo
            Command1(4).Caption = "Play RTTTL"
        End Select

      Case 3 'Delete last
        If Len(Text1(0)) Then ' is text
            If InStr(Text1(0), " ") Then 'more than one note
                Text1(0).Text = Left$(Text1(0).Text, InStrRev(Text1(0).Text, " ") - 1)
              Else 'only one note'NOT INSTR(TEXT1(0),...
                Text1(0) = ""
                VSS.ClearScore
            End If
        End If

        If Len(Text1(1)) Then ' is text
            If InStr(Text1(1), " ") Then 'more than one note
                Text1(1).Text = Left$(Text1(1).Text, InStrRev(Text1(1).Text, " ") - 1)
                VSS.AddNoteArray Split(Trim$(Text1(1).Text))
              Else 'only one note'NOT INSTR(TEXT1(0),...'NOT INSTR(TEXT1(1),...
                Text1(1) = ""
                VSS.ClearScore
            End If
        End If

        If Len(Text1(2)) Then ' is text
            If InStr(Text1(2), ",") Then 'more than one note
                Text1(2).Text = Left$(Text1(2).Text, InStrRev(Text1(2).Text, ",") - 1)

              Else 'only one note'NOT INSTR(TEXT1(0),...'NOT INSTR(TEXT1(1),...'NOT INSTR(TEXT1(2),...
                Text1(1) = ""
                VSS.ClearScore
            End If
        End If
        VSS.AddNoteArray Split(Trim$(Text1(0).Text))
      Case 6
        KB.KeyBoardLookPrint
    End Select

End Sub

Private Sub Form_Load()

  Dim i As Integer

    Me.Width = Screen.Width * 2 / 3
    With cmbLooks
        .AddItem "Antique"
        .AddItem "Classical"
        .AddItem "Default"
        .AddItem "Random"
        .Text = "Default"
    End With 'CMBLOOKS
    With cmbTempo
        For i = 25 To 900
            .AddItem i
        Next i
        .Text = 120
    End With 'CMBTEMPO
    With KB
        Set .AssignControl = Picture1
        .PauseKeyOn = True

    End With 'KB
    Label2.Caption = "Mouse Note Durations ( Beats per Note)                              " & _
                     "Left=" & KB.Fractional(4 / KB.GetButtonValue(ButtonL)) & "             Right=" & KB.Fractional(4 / KB.GetButtonValue(ButtonR)) & "        Mid=" & KB.Fractional(4 / KB.GetButtonValue(ButtonM)) & _
                     "          [Shift] + Left=" & KB.Fractional(4 / KB.GetButtonValue(ShiftButtonL)) & "      +Right=" & KB.Fractional(4 / KB.GetButtonValue(ShiftButtonR)) & "   +Mid=" & KB.Fractional(4 / KB.GetButtonValue(ShiftButtonM)) & _
                     "          [Ctrl]+  [Any Button]=" & KB.Fractional(4 / KB.GetButtonValue(CtrlButtonL)) & _
                     " -----------------------------------------------------------------------------   " & _
                     "Dotted note ( times 1.5)   Use [Alt] while clicking"

    'Nokia format ="Nokia:<title>::Tempo:<tempoValue>:::<tune>
    'Basica format = "Basica:<title::<tune>
    ' the program reads the type (Nokia or Basica) to decide how to read the values into controls
    'the number and position of the ":" is very important
    With Combo1
        .AddItem "Nokia:Mission Impossible - Movie Theme::Tempo:112:::16G2 8P 16G2 8P 16F2 16P 16#F2 16P 16G2 8P 16G2 8P 16#A2 16P 16C3 16P 16G2 8P 16G2 8P 16F2 16P 16#F2 16P 16G2 8P 16G2 8P 16#A2 16P 16C3 16P 16#A2 16G2 2D2 32P 16#A2 16G2 2#C2 32P 16#A2 16G2 2C2 16P 16#A1 16C2"
        .AddItem "Nokia:Britney Spears - Anticipation::Tempo:125:::16F2 16- 8.C2 16- 16A1 16- 16A1 8.C2 8- 16C2 16C2 16C2 16C2 8C2 16C2 16C2 16C2 16C2 16C2 16C2 8D2 16C2 8- 16- 8F2 16C2 16- 16C2 16- 8A1 16C2 8.C2 16A1 16- 16C2 16- 16C2 16C2 8C2 16.C2 32- 16.A1 32- 16G1 16A1 8G1 4- 16A1 16- 16C2 16- 8C2 16.A1 32- 16C2 8.C2 8- 16C2 16- 16C2 16- 8C2 8- 16C2 16C2 16C2 16C2 16D2 16D2 8C2 8- 16C2 16- 8C2 16C2 16C2 16.A1 32- 16C2 8C2 16- 16A1 16A1 8C2 16C2 16- 16C2 16- 16C2 16- 16.A1 32- 16G1 16F1 16G1 8F1"
        .AddItem "Nokia:Barney Bralligans(Slip Jig)::Tempo:140::: 4#f1 8a1 8a1 8b1 8a1 8a1 8b1 8a1 4#f1 8a1 8a1 8b1 8a1 4d2 8#f2 4#f1 8a1 8a1 8b1 8a1 8a1 8b1 8a1 8b1 8a1 8b1 4e2 8d2 8#c2 8b1 8a1"
        'Keypresses: 4#. 68, 6, 7, 6, 6, 7, 6, 49#, 68, 6, 7, 6, 29*, 48#, 49#**, 68, 6, 7, 6, 6, 7, 6, 7, 6, 7, 39*, 28, 1#, 7**, 6
        .AddItem "Nokia:The Butterfly (Slip Jig)::Tempo:140:::4b1 8e1 4g1 8e1 4.#f1 4b1 8e1 4g1 8e1 8#f1 8e1 8d1 4b1 8e1 4g1 8e1 4.#f1 4b1 8d2 4d2 8b1 8a1 8#f1 8d1 4-"
        'Keypresses: 7, 38, 59, 38, hold49#, 7, 38, 59, 38, 4#, 3, 2,                  79, 38, 59, 38, hold49#, 7, 28*, 29, 78**, 6, 4#, 2, 09
        .AddItem "Nokia:Foxhunter's Jig Part One (Slip Jig)::Tempo:140:::8#f1 8g1 8#f1 4#f1 8d1 4g1 8e1 8#f1 8g1 8#f1 4#f1 8d1 4e1 8d1 8#f1 8g1 8#f1 4#f1 8d1 4g1 8b1 8a1 8#f1 8d1 8d1 8e1 8#f1 4e1 8d1"
        'Keypresses: 48#, 5, 4#, 49#, 28, 59, 38, 4#, 5, 4#, 49#, 28,                  39, 28, 4#, 5, 4#, 49#, 28, 59, 78, 6, 4#, 2, 2, 3, 4#, 39, 28
        .AddItem "Nokia:Foxhunter's Jig Part Two (Slip Jig)::Tempo:140:::4.b1 8b1 8a1 8g1 8#f1 8g1 8a1 4b1 8e1 4e1 8#f1 4g1 8b1 8a1 8b1 8#c2 8d2 8#c2 8b1 8a1 8b1 8#c2 4d2 8d1 8d1 8e1 8#f1 4e1 8d1"
        'Keypresses: hold 7, 78, 6, 5, 4#, 5, 6, 79, 38, 39, 48#, 59,                  78, 6, 7, 1#*, 2, 1#, 7**, 6, 7, 1#*, 29, 28**, 2, 3, 4#, 39,28
        .AddItem "Nokia:Glasgow Reel(Reel)::Tempo: 112::: 8a1 16d2 16a1 16f2 16a1 16d2 16a1 8#a1 16d2 16#a1 16f2 16#a1 16d2 16#a1 8c2 16e2 16c2 16g2 16c2 16e2 16g2 16f2 16e2 16d2 16f2 16e2 16d2 16c2 16#a1 16a1"
        'Submitted by Raymond Chambers of Clicky Feet. Check his site for more ringtones.
        'Keypresses: 68, 28*, 6**, 4*, 6**, 2*, 6**, 69#, 28*, 6**#, 4*, 6**#, 2*, 6**#, 19*, 38, 1, 5, 1, 3, 5, 4, 3, 2, 4, 3, 2,1, 6**#, 6
        .AddItem "Nokia:Kid on the Mountain (Slip Jig)::Tempo: 140:::8e1 8d1 8e1 8#f1 8e1 8#f1 4g1 8#f1 8e1 8#f1 8e1 4c2 8a1 8b1 8a1 8g1 8e1 8d1 8e1 8#f1 8e1 8#f1 4g1 8a1 8b1 8a1 8g1 8#f1 8a1 8g1 8#f1 8e1 8d1"
        'Keypresses:38, 2, 3, 4#, 3, 4#, 59, 48#, 3, 4#, 3, 19*, 68**, 7, 6, 5, 3, 2, 3, 4#, 3, 4#, 59, 68, 7, 6, 5, 4#, 6, 5, 4#, 3,2
        .AddItem "Nokia:King of the Fairies (Hornpipe Set Dance)::Tempo: 140:::4d1 8.e1 8d1 8.e1 8#f1 8.g1 8#f1 8.g1 8a1 4b1 16- 4b116- 8.g1 8#f1 8.g1 8a1 4b1 16- 4e1 16- 8.e1 8#f1 8.g1 8e1 8.#f1 8g1 8.#f1 8e1 4d1"
        'Keypresses: 2, (hold 3)8, 2, (hold 3), 4#, (hold 5), 4#, (hold                  5), 6, 79, 088, 7, 088, (hold 5)8, 4#, (hold 5), 6, 79, 088,                  3, 088, (hold 3)8, 4#, (hold 5), 3, (hold 4)#, 5, (hold 4)#,                  3, 29
        .AddItem "Nokia:Lark in the Morning (Jig)::Tempo:160:::4.a1 8a1 8#f1 8a1 4.b1 8b1 8d2 8b1 4.a1 8a1 8#f1 8a1 8#f2 8e2 8d2 8b1 8d2 8b1 4.a1 8a1 8#f1 8a1 4.b1 8b1 8d2 8b1 8d2 8e2 8#f2 8a2 8#f2 8e2 8#f2 8d2 8b1 4d2"
        'Keypresses:hold 6, 68, 4#, 6, hold 79, 78, 2*, 7**, hold 69,                  68, 4#, 6, 4#*, 3, 2, 7**, 2*, 7**, hold 6 9, 68, 4#, 6, hold                  7 9, 78, 2*, 7**, 2*, 3, 4#, 6, 4#, 3, 4#, 2, 7**, 29*
        .AddItem "Nokia:Madame Bonaparte (Set dance)::Tempo:140:::8e2 8d2 4#c2 8#c2 8b1 8#c2 8e2 8#c2 8a1 4d2 8d2 8#c2 8d2 8#f2 8e2 8b1 8a1 8#c2 8e2 8#g2 8a2 8#g2 8a2 8#f2 4e2 8e2 8#f2 8e2 8d2 8#c2 8b1"
        'Keypresses: 38*, 2, 19#*, 18#, 7**, 1#*, 3, 1#, 6**, 29*, 28,                  1#, 2, 4#, 3, 7**, 6, 1#*, 3, 5#, 6, 5#, 6, 4#, 39, 38, 4#, 3,                  2, 1#, 7**
        .AddItem "Nokia:Monaghan Jig(Jig)::Tempo:180:::8b1 8g1 8e1 4#f1 8e1 8b1 8e1 8#f1 8g1 8a1 8b1 8g1 8e1 4#f1 8e1 8a1 8#f1 8d1 8#f1 8g1 8a1 8b1 8g1 8e1 4#f1 8e1 8b1 8g1 8e1 8#f1 8g1 8a1 8d2 8c2 8b1 8a1 8b1 8g1 8#f1 8d1 8#f1 8a1 8#f1 8d1"
        'Keypresses: 78, 5, 3, 49#, 38, 7, 3, 4#, 5, 6, 7, 5, 3, 49#,                  38, 6, 4#, 2, 4#, 5, 6, 7, 5, 3, 49#, 38, 7, 5, 3, 4#, 5, 6,                  2*, 1, 7**, 6, 7, 5, 4#, 2, 4#, 6, 4#, 2
        .AddItem "Nokia:Morrison's Jig (Jig)::Tempo:160:::4.e1 4.b1 8e1 8#f1 8g1 8a1 8#f1 8d1 4.e1 8b1 8a1 8b1 8d2 8c2 8b1 8a1 8#f1 8d1 4.e1 8b1 8a1 8b1 4.e1 8a1 8#f1 8d1 4.g1 8#f1 8g1 8a1 8d2 8a1 8g1 8#f1 8e1 8d1"
        'Keypresses: hold 3, hold 7, 38, 4#, 5, 6, 4#, 2, hold 3 9, 78,                  6, 7, 2*, 1, 7**, 6, 4#, 2, hold 3 9, 78, 6, 7, hold 3 9, 78,                  4#, 2, hold 5 9, 48#, 5, 6, 2*, 6**, 5, 4#, 3, 2
        .AddItem "Nokia:Sally Gardens(Reel)::Tempo: 160:::4g1 8d1 8g1 4b1 8g1 8b1 8d2 8b1 8e2 8b1 8d2 8b1 8a1 8b1 4d2 8b1 8d2 8e2 8g2 8d2 8b1 8a1 8g1 8a1 8b1 8a1 8g1 8e1 8d12 g1"
        'Keypresses: 5, 28, 5, 79, 58, 7, 2*, 7**, 3*, 7**, 2*, 7**, 6,                  7, 29*, 78**, 2*, 3, 5, 2, 7**, 6, 5, 6, 7, 6, 5, 3, 2, 599
        .AddItem "Nokia:Star of Munster (Reel)::Tempo:160:::4c2 8a1 8c2 8b1 8a1 8g1 8b1 8a1 8g1 8e1 8#f1 8g1 8e1 4d1 8e1 8a1 8a1 8g1 8a1 8b1 8c2 8d2 8e2 8c2 8d2 8b1 8c2 8a1 4a1 4-"
        'Keypresses:1*, 68**, 1*, 7**, 6, 5, 7, 6, 5, 3, 4#, 5, 3, 29,                  38, 6, 6, 5, 6, 7, 1*, 2, 3, 1, 2, 7**, 1*, 6**, 69,
        .AddItem "Nokia:St Patrick's Day (Set dance)::Tempo: 140:::8d1 8g1 8a1 8g1 8g1 8b1 8d2 8g2 8#f2 8e2 8d2 8b1 8g1 8a1 8g1 8a1 8b1 8g1 8d1 8e1 8#f1 8e1 4e1 8d1 8g1 8a1 8g1 8g1 8b1 8d2 8g2 8#f2 8e2 8d2 8b1 8g1 8a1 8g1 8a1 8b1 8g1 8d1 4e1 8#f1 4g1"
        'Keypresses: 28, 5, 6, 5, 5, 7, 2*, 5, 4#, 3, 2, 7**, 5, 6, 5,                  6, 7, 5, 2, 3, 4#, 3, 39, 28, 5, 6, 5, 5, 7, 2*, 5, 4#, 3, 2,                  7**, 5, 6, 5, 6, 7, 5, 2, 39, 48#, 59

        .AddItem "Basica:Mary Had a Little Lamb::MN T100 O3 L8 GFE-FGGGP8 FFF4GB-B-4GFE-FGGGGFFGFE-."
        .AddItem "Basica:Kodok Ngorek::MN T100 L8 O3 GEEEGEEEGAGFED4 FDDDFDDDFGFEDC4"
        .AddItem "Basica:Topi Saya Bundar::MS T120 L4 O2 GG2E>C2<ED2....P4 EF2AG3FE2....P4  GG2E>C2<ED2....P4 >C<B2GA2B>C2...."
        .AddItem "Basica:Suwe Ora Jamu::MN T120 L4 O2 F#8G8A.A8F+GA2. F+G.G8AF#G2. A>D#.C#8DDC#. C#8<G.G8F#>D2."
        .AddItem "Basica:Gambang Suling::MN T108 L8 O2 >GA-G>C4.<GA-GFE-D-2 P8>CE-CD-4.<GA-GE-D-C2 P8A-GA->C4. E-D-E-C<A-G4F4.FGA-G4 E-4.E-GE-D-4 F4.FGA--G4.>D<A-GE-D-C2"
        .AddItem "Basica:Surilang::MN T120 L8 O2 P8<16B-16>E-FGA-16G16A->C1<P8<16B-16>E-FGA-16G16A->C1< P4F4FFA-GE-1 C16 C16CD-C<A-16A-B->D4C16<16A-16G16F2P8F16F16F16FF16F16A-GE-1 P4GA-B-GE-C1<"
        .AddItem "Basica:Meyong::MN T108 L8 O2 BBGB4.B>C<GFG2P8BBB4.B>C<GFG4.GBGFEF4.FGB>C<GB>C<GFEGFE2"
        .AddItem "Basica:Ngusik-Asik::MN T150 L4 O2 DE-FA2B->E-D2<2A1P4B-AFE-FDE-FAFE-D1P4E-D<E-1P4FDE-FE-B-FE-D<E-D1"
        .AddItem "Basica:Sounds of Silence-Paul Simon::T120 L8 P4 O3 D D F F A A L1 G P8 L8 C C C E E G G L1 F P8 L8 F F F A A O4 C C L2 D C P4 L8 O3 F F A A O4 C C L2 D C P4 L8 O3 F F O4 D L2 D. L4 D L8 D F E L4 D. L2 C L4 C. L8 D C L2 O3 A A P8 L8 F F F L2 O4 C. P8 L8 O3 E F L4 D. L2 D"
        .AddItem "Basica:We Shall Overcome::T112 L4 O3 G G A A L2 G E L4 G G A A L2 G E L4 G G A B L2 O4 C D O3 B"
        .AddItem "Basica:Nearer My God to Thee::T120 L2 O3 A L4 G F F. L8 D L2 D C L4 F A L2 G. P4 A L4 G F F. L8 D L2 D L4 C F F G L2 G. P4 O4 E L4 E D D. O3 B L2 O4 D D L4 E D D. L8 O3 B L2 G A L4 G F F. L8 D L2 D C L4 F A L2 G. P4"
        .AddItem "Basica:Michael haul the boat ashore::T120 L4 O3 D F A. L8 F A L4 B A L2 A L4 F L1 B L2 A L4 F A A. L8 F G L4 F L8 E L2 E L4 D E L2 F F L1 D"
        .AddItem "Basica:Silent Night::T110 L2 O3 F. L8 G L4 F L2 D L4 F. L8 G L4 F L2 D. O4 D L4 D L2 O3 B. B L4 B L2 F. G L4 G B. L8 A L4 G F. L8 G L4 F L2 D"
        .AddItem "RTTL:Tubular:d=4,o=5,b=285:c,f,c,g,c,d#,f,c,g#,c,a#,c,g,g#,c,g,c,f,c,g,c,d#,f,c,g#,c,a#,c,g,g#,c,g,c,f,c,g,c,d#,f,c,g#,c,a#,c,g,g#,c,g,c,f,c,g,c,d#,f,c,g#,c,a#,c,g,g#,c,g"
        .AddItem "RTTL:Popcorn:d=4,o=5,b=160:8c6,8a#,8c6,8g,8d#,8g,c,8c6,8a#,8c6,8g,8d#,8g,c,8c6,8d6,8d#6,16c6,8d#6,16c6,8d#6,8d6,16a#,8d6,16a#,8d6,8c6,8a#,8g,8a#,c6"
        .AddItem "RTTL:Money:d=4,o=5,b=112:8e6,8e6,8e6,8e6,8e6,8e6,16e,16a,16c6,16e6,8d#6,8d#6,8d#6,8d#6,8d#6,8d#6,16f,16a,16c6,16d#6,d6,8c6,8a,8c6,c6,2a,32a,32c6,32e6,8a6"
        .AddItem "RTTL:90210:d=4,o=5,b=140:8f,8a#,8c6,d.6,2d6,p,8f,8a#,8c6,8d6,8d#6,f6,f.6,2a#.,8f,8a#,8c6,8d6,8d#6,8f6,8g6,f6,8d#6,d#6,d6,2c.6,8a#,a,a#.,g6,8f6,8d#6,8d6,8d#6,8d6,8a#,f"
        .AddItem "RTTL:Abdelazer:d=4,o=5,b=160:2d,2f,2a,d6,8e6,8f6,8g6,8f6,8e6,8d6,2c#6,a6,8d6,8f6,8a6,8f6,d6,2a6,g6,8c6,8e6,8g6,8e6,c6,2a6,f6,8b,8d6,8f6,8d6,b,2g6,e6,8a,8c#6,8e6,8c6,a,2f6,8e6,8f6,8e6,8d6,c#6,f6,8e6,8f6,8e6,8d6,a,d6,8c#6,8d6,8e6,8d6,2d6"
        .AddItem "RTTL:Agadoo:d=4,o=5,b=125:8b,8g#,e,8e,8e,e,8e,8e,8e,8e,8d#,8e,f#,8a,8f#,d#,8d#,8d#,d#,8d#,8d#,8d#,8d#,8c#,8d#,e"
        .AddItem "RTTL:axelf:d=4,o=5,b=160:f#,8a.,8f#,16f#,8a#,8f#,8e,f#,8c.6,8f#,16f#,8d6,8c#6,8a,8f#,8c#6,8f#6,16f#,8e,16e,8c#,8g#,f#."
        .AddItem "RTTL:Barbie girl:d=4,o=5,b=125:8g#,8e,8g#,8c#6,a,p,8f#,8d#,8f#,8b,g#,8f#,8e,p,8e,8c#,f#,c#,p,8f#,8e,g#,f# "
        .AddItem "RTTL:bond:d=4,o=5,b=320:c,8d,8d,d,2d,c,c,c,c,8d#,8d#,2d#,d,d,d,c,8d,8d,d,2d,c,c,c,c,8d#,8d#,d#,2d#,d,c#,c,c6,1b.,g,f,1g."
        .AddItem "RTTL:Bulletme:d=4,o=5,b=112:b.6,g.6,16f#6,16g6,16f#6,8d.6,8e6,p,16e6,16f#6,16g6,8f#.6,8g6,8a6,b.6,g.6,16f#6,16g6,16f#6,8d.6,8e6,p,16c6,16b,16a,16b"
        .AddItem "RTTL:Star Trek:d=4,o=5,b=63:8f.,16a#,d#.6,8d6,16a#.,16g.,16c.6,f6"
        .AddItem "RTTL:VanessaMae:d=4,o=6,b=70:32c7,32b,16c7,32g,32p,32g,32p,32d#,32p,32d#,32p,32c,32p,32c,32p,32c7,32b,16c7,32g#,32p,32g#,32p,32f,32p,16f,32c,32p,32c,32p,32c7,32b,16c7,32g,32p,32g,32p,32d#,32p,32d#,32p,32c,32p,32c,32p,32g,32f,32d#,32d,32c,32d,32d#,32c,32d#,32f,16g,8p,16d7,32c7,32d7,32a#,32d7,32a,32d7,32g,16d7,32p,32d7,32p,32d7,32p,16d7,32c7,32d7,32a#,32d7,32a,32d7,32g,16d7,32p,32d7,32p,32d7,32p,32g,32f,32d#, 32d,32c,32d,32d#,32c,32d#,32d,8c"

        .AddItem "RTTL:Walk of Life:d=4,o=5,b=160:b.,b.,p,8p,8f#,8g,b,8g,8f,e.,e.,p,2p,p,8f,8g,b.,b.,p,8p,8f,8g,b,8g,f,e.,e.,p,8p,8f,8g,b,8g,8f,8e"

    End With 'COMBO1

    Set VSS.VirtualScoreSheet = Picture2

    Command1(3).Enabled = Len(Text1(0)) > 0 Or Len(Text1(1)) > 0

End Sub

Private Sub Form_Resize()

    With FrmKBDemo
        If .WindowState <> vbMinimized Then
            Frame1.Left = 20
            Frame1.Width = FrmKBDemo.ScaleWidth - 50
            Picture1.Width = Frame1.Width - 180
            Picture2.Width = Frame1.Width - 180
            Text1(0).Width = FrmKBDemo.ScaleWidth - Command1(0).Width
            Text1(1).Width = FrmKBDemo.ScaleWidth - Command1(1).Width
            Text1(2).Width = FrmKBDemo.ScaleWidth - Command1(1).Width
            Command1(0).Left = Text1(0).Width
            Check4.Left = Command1(0).Left

            cmbTempo.Left = Command1(0).Left
            Command1(1).Left = Text1(1).Width
            Command1(4).Left = Text1(1).Width
            Command1(2).Left = Text1(1).Width
            Command1(6).Left = Text1(1).Width - Command1(6).Width + Command1(2).Width
            cmbLooks.Left = Command1(6).Left - cmbLooks.Width

            Label1.Left = cmbLooks.Left - Label1.Width
            Check3.Left = Label1.Left - Check3.Width - 200
            Check2.Left = Label1.Left - Check2.Width - 200
            Check1.Left = FrmKBDemo.ScaleWidth - Check1.Width
            Command1(3).Left = Text1(1).Width - Command1(3).Width
            Combo1.Left = 120
            Combo1.Width = Command1(3).Left - 120
            'PICTURE1'FRMKBDEMO
            KB.Resize
            If .Height > Picture2.Top + 600 Then
                Picture2.Height = .Height - Picture2.Top - 600
                VSS.Resize
            End If
        End If
    End With 'FRMKBDEMO

End Sub

Private Sub Form_Unload(Cancel As Integer)

    KB.Abort
    End

End Sub

Private Sub mnuaboutopt_Click(Index As Integer)

    Select Case Index
      Case 0
        KB.About
      Case 1
        VSS.About
    End Select

End Sub

Private Sub mnuBasicaOpt_Click(Index As Integer)

    Select Case Index
      Case 0
        Text1(1).Text = KB.BasicaOctaveShift(Text1(1).Text, 1)
      Case 1
        Text1(1).Text = KB.BasicaOctaveShift(Text1(1).Text, -1)
      Case 2
        Text1(1).Text = KB.BasicaNoteShift(Text1(1).Text, 1)
      Case 3
        Text1(1).Text = KB.BasicaNoteShift(Text1(1).Text, -1)
      Case 4
        Text1(1).Text = KB.BasicaRandomTune
     '  Case 5
      'Text1(1).Text = KB.BasicaRealTranspose(Text1(1).Text, "C", "D")
    End Select

End Sub

Private Sub mnuhelpopt_Click(Index As Integer)

    Select Case Index
      Case 0
        KBDemoHelp.Show
      Case 1

    End Select

End Sub

Private Sub MnuTrans_Click(Index As Integer)
Dim Keys As Variant
Keys = Array("D", "G-", "B-", "E-", "G#", "C#", "B", "E", "A", "D2", "G-2")
Text1(1).Text = KB.BasicaRealTranspose(Text1(1).Text, "C", CStr(Keys(Index - 1)))

End Sub

Private Sub mnuWarning_Click()

    MsgBox "The stuff in this menu is experimental and has no serious error checking." & vbNewLine & _
           "If you drive the octave off the key board it will be auto-reset to the limiter." & vbNewLine & _
           "Shift sharp/flat notes to illegal notes and the sharp/flat will be deleted." & vbNewLine & _
           "I may add this stuff to the class at some time in the future.", _
           vbInformation, _
           "Basica Tricks"

End Sub

''Uncomment if you want to used XP Theme in compiled program
'remember that this will make the program very dangerous to install in win95
'see my upload 'WARNING XP styles and Win95' at PSC for details
'Private Sub Form_Initialize() ':) Line inserted by Formatter
'
'    InitCommonControls ':) Line inserted by Formatter
'
'End Sub ':) Line inserted by Formatter

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    KB.MouseDown Button, Shift, X, Y

End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    KB.MouseMove Button, Shift, X, Y

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  Dim AddNote As String

    KB.MouseUp Button, Shift, X, Y
    AddNote = KB.NoteTextOutPut(Nokia)
    If Len(AddNote) Then
        Text1(0).Text = Text1(0).Text & " " & AddNote
    End If
    AddNote = KB.NoteTextOutPut(Basica)
    If Len(AddNote) Then
        Text1(1).Text = Text1(1).Text & " " & AddNote
        VSS.AddNoteArray Split(Trim$(Text1(1).Text)) ', Basica)
    End If
    AddNote = KB.NoteTextOutPut(RTTTL)
    If Len(AddNote) Then
        Text1(2).Text = Text1(2).Text & IIf(Len(Text1(2)), ",", "") & AddNote

    End If

End Sub

Private Sub Text1_Change(Index As Integer)

    Command1(3).Enabled = Len(Text1(0)) > 0 Or Len(Text1(1)) > 0 Or Len(Text1(2)) > 0
    Command1(2).Enabled = Len(Text1(0)) > 0 Or Len(Text1(1)) > 0 Or Len(Text1(2)) > 0

    Command1(0).Enabled = Len(Text1(0)) > 0
    Command1(1).Enabled = Len(Text1(1)) > 0
    Command1(4).Enabled = Len(Text1(2)) > 0

End Sub

':) Ulli's VB Code Formatter V2.13.6 (9/10/2002 11:45:48 AM) 27 + 444 = 471 Lines
