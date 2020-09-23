VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsVirtualScoreSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
Private M_VSS As PictureBox
Private m_TopLine As Long
Private m_LinSep As Long
Private m_LineWidth As Integer
Private m_MidCpos As Long
Private BarLine As Single
Private BeatsPerBar As Integer
Private NOffset As Long
Private PI As Single
Private m_BassShow As Boolean
Private m_ShowNokiaRange As Boolean
Private Avg_Barlen As Long
Private LastTune As String
Private Bars As Long
Private CurOctave As Integer
Private CurTempo As Integer
Private CurLength As Single
Private BarlenFactor As Single
'Make this pulbic if you are not using it in conjunction with clsKeyBoardPicture
Private Enum TextTypeOutPut ' Enum for NoteTextOutPut
    Basica
    Nokia
End Enum

Rem Mark Off
'Stops Code formatter complaining about these
#If False Then 'Enforce Case For Enums (does not compile but fools IDE)
Dim Basica
Dim Nokia

#End If  'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1
Rem Mark On
Public Sub About()

  Dim msg As String

    msg = "This Class is still too experimental to have stable documentation." & vbNewLine & _
          "Copyright 2002 Roger Gilchrist"
    MsgBox msg, vbInformation, "ClsVirtualScoreSheet"

End Sub

Public Sub AddNote(Note As String)

  'read and print a single basica note at a time

    LastTune = Trim$(LastTune & " " & Note)

    DrawTune Trim$(Note), 0

End Sub

Public Sub AddNoteArray(Notes)

  'read and print a complete line of Basica Text

  Dim i As Long

    ClearScore
    Avg_Barlen = 0
    LastTune = Join(Notes)
    BarlenFactor = 1
    For i = LBound(Notes) To UBound(Notes)
        DrawTune Notes(i), 0
    Next i

End Sub

Private Sub Class_Initialize()

    PI = 4 * Atn(1)
    LineSeperationSize = 120
    ShowBassStave = True
    LineWidth = 1
    ShowNokiaRange = False

End Sub

Public Sub ClearScore()

    m_TopLine = 800
    m_MidCpos = m_TopLine + (5.5 * m_LinSep)
    BarLine = 0
    NOffset = 0
    BarLine = 0
    BeatsPerBar = 4

    If M_VSS Is Nothing Then
        'not initialised so don't try to do anything
      Else 'NOT M_VSS...
        M_VSS.Cls
        DrawStave
    End If

End Sub

Private Function CountChar(a, c$) As Long

  Dim i As Integer

    For i = 1 To Len(a)
        If Mid$(a, i, 1) = c$ Then
            CountChar = CountChar + 1
        End If
    Next i

End Function

Private Sub DrawANote(NOffset As Long, NPos As Long, Dur As Single, dot As Integer, sharp As Boolean, Flat As Boolean, SharpOffset As Long)

  Dim High As Single, i As Integer, UpDown As Boolean

    M_VSS.DrawWidth = m_LineWidth
    UpDown = (NPos > m_TopLine + m_LinSep * 2.5)
    If sharp Then
        DrawSharp NOffset, NPos
        SharpOffset = m_LinSep
    End If
    If Flat Then
        DrawFlat NOffset, NPos
        SharpOffset = m_LinSep * 0.25
    End If
    NOffset = NOffset + SharpOffset + 100
    M_VSS.FillColor = IIf(Dur > 3, vbBlack, vbWhite)
    M_VSS.FillStyle = IIf(Dur > 3, vbFSSolid, vbFSTransparent)

    High = 3.5
    If Dur > 32 Then
        High = 5.5
      ElseIf Dur > 16 Then 'NOT DUR...
        High = 4.5
    End If
    If Dur = 1 Or Dur = 2 Then
        M_VSS.DrawWidth = m_LineWidth * 2.5
    End If
    M_VSS.Circle (NOffset, NPos), m_LinSep / 1.5, vbBlack, , , 0.65
    M_VSS.DrawWidth = m_LineWidth
    If Dur > 1 Then
        If UpDown Then
            M_VSS.Line Step(m_LinSep / 1.5, 0)-Step(0, -m_LinSep * High)
          Else 'UPDOWN = FALSE
            M_VSS.Line Step(-m_LinSep / 1.5, 0)-Step(0, m_LinSep * High)
        End If
        M_VSS.DrawWidth = 1
        If Dur > 4 Then
            'M_VSS.Line Step(0, 0)-Step(m_LinSep / 2, m_LinSep * 2.5)
            If UpDown Then
                M_VSS.Circle Step(0, m_LinSep * IIf(Dur = 8, 3.1, 2)), m_LinSep * IIf(Dur = 8, 3.3, 2), vbBlack, 2 * PI, PI / 2, IIf(Dur = 8, 3.5, 2.5)
              Else 'UPDOWN = FALSE
                M_VSS.Circle Step(0, -m_LinSep * IIf(Dur = 8, 2.7, 2)), m_LinSep * IIf(Dur = 8, 2.7, 2), vbBlack, 3 * PI / 2, 2 * PI, IIf(Dur = 8, 3.5, 2.5)
            End If
            If Dur > 8 Then
                If UpDown Then
                    M_VSS.Circle Step(0, m_LinSep), m_LinSep * 2, vbBlack, 2 * PI, PI / 2, 2.5
                  Else 'UPDOWN = FALSE
                    M_VSS.Circle Step(0, -m_LinSep), m_LinSep * 2, vbBlack, 3 * PI / 2, 2 * PI, 2.5
                End If
                If Dur > 16 Then
                    If UpDown Then
                        M_VSS.Circle Step(0, m_LinSep), m_LinSep * 2, vbBlack, 2 * PI, PI / 2, 2.5
                      Else 'UPDOWN = FALSE
                        M_VSS.Circle Step(0, -m_LinSep), m_LinSep * 2, vbBlack, 3 * PI / 2, 2 * PI, 2.5
                    End If

                    If Dur > 32 Then

                        If UpDown Then
                            M_VSS.Circle Step(0, m_LinSep), m_LinSep * 2, vbBlack, 2 * PI, PI / 2, 2.5
                          Else 'UPDOWN = FALSE
                            M_VSS.Circle Step(0, -m_LinSep), m_LinSep * 2, vbBlack, 3 * PI / 2, 2 * PI, 2.5
                        End If
                    End If
                End If
            End If
        End If
    End If
    If dot > 0 Then
        M_VSS.FillStyle = vbFSSolid
        M_VSS.FillColor = vbBlack
        For i = 1 To dot
            M_VSS.Circle (NOffset + m_LinSep / 2 + m_LinSep * i, NPos), m_LinSep / 5, vbBlack
        Next i
        M_VSS.FillStyle = vbFSTransparent
    End If

    M_VSS.DrawWidth = 2

End Sub

Private Sub DrawBarLine()

    NOffset = NOffset + m_LinSep
    M_VSS.Line (NOffset, m_TopLine)-Step(0, m_LinSep * 4), vbBlack
    If m_BassShow Then
        M_VSS.Line (NOffset, m_TopLine + m_LinSep * 7)-Step(0, m_LinSep * 4), vbBlack
    End If
    NOffset = NOffset + m_LinSep
    If NOffset + Avg_Barlen > M_VSS.Width Then
        Avg_Barlen = 0
        Bars = 0
        BarLine = 0
        m_TopLine = m_TopLine + m_LinSep * IIf(m_BassShow, 16, 10)
        m_MidCpos = m_MidCpos + m_LinSep * IIf(m_BassShow, 16, 10)
        ' NPos = NPos + m_MidCpos 'move notes to next stave
        DrawStave
        NOffset = NOffset + m_LinSep

    End If

    M_VSS.DrawWidth = m_LineWidth

End Sub

Private Sub DrawFlat(X, Y)

  Dim Yoffset As Long

    M_VSS.DrawWidth = m_LineWidth
    Yoffset = Y - m_LinSep
    M_VSS.Line (X, Yoffset)-Step(0, m_LinSep * 1.5), vbBlack  'LeftVert
    M_VSS.Circle Step(0, -m_LinSep / 2), m_LinSep / 4, vbBlack, 3 * PI / 2, PI / 2, 0.9

End Sub

Private Sub DrawRest(NOffset As Long, Dur As Single, dot As Integer)

  'Draw Restnotes for 1,2,4,8,16,32,64 Durations (or nearest Integr value)

  Dim i  As Integer

    M_VSS.FillStyle = vbFSSolid
    M_VSS.FillColor = vbBlack
    Select Case Int(Dur)
      Case 1 To 2.999999999999
        M_VSS.Line (NOffset, m_TopLine + m_LinSep * Dur)-Step(m_LinSep, -m_LinSep * IIf(Dur = 1, -0.5, 0.5)), vbBlack, BF
      Case 3.0000000000001 To 3.9999999999999
        For i = 1 To 50
            M_VSS.Line (NOffset, m_TopLine + m_LinSep * 0.5)-Step((m_LinSep * 0.5) - i, m_LinSep * 1), vbBlack
        Next i
        For i = 1 To 50
            M_VSS.Line (NOffset + 25 + i, m_TopLine + m_LinSep * 1.5)-Step(-70, 100), vbBlack
        Next i
        For i = 1 To 5
            M_VSS.Line (NOffset - 25 + i, m_TopLine + m_LinSep * 2)-Step(100, 200), vbBlack
        Next i
        For i = 1 To 25
            M_VSS.Line (NOffset + i, m_TopLine + m_LinSep * 2.5)-Step(-50, 100), vbBlack
        Next i
        M_VSS.Line (NOffset - 25, m_TopLine + m_LinSep * 3)-Step(50, 100), vbBlack

      Case Else
        M_VSS.Line (NOffset, m_TopLine + m_LinSep * IIf(Dur <= 8, 3, 4))-Step(m_LinSep * 0.5, -m_LinSep * IIf(Dur <= 16, IIf(Dur > 8, 3, 2), 4)), vbBlack
        M_VSS.Circle Step(-m_LinSep * 0.5, m_LinSep * 0.35), m_LinSep * 0.25, vbBlack
        M_VSS.Circle Step(m_LinSep * 0.25, -m_LinSep * 0.25), m_LinSep * 0.25, vbBlack, 3 * PI / 2, 2 * PI
        If Dur >= 16 Then
            M_VSS.Circle Step(-m_LinSep * 0.4, m_LinSep * 1.35), m_LinSep * 0.25, vbBlack
            M_VSS.Circle Step(m_LinSep * 0.25, -m_LinSep * 0.25), m_LinSep * 0.25, vbBlack, 3 * PI / 2, 2 * PI
        End If
        If Dur >= 32 Then
            M_VSS.Circle Step(-m_LinSep * 0.4, m_LinSep * 1.25), m_LinSep * 0.25, vbBlack
            M_VSS.Circle Step(m_LinSep * 0.25, -m_LinSep * 0.25), m_LinSep * 0.25, vbBlack, 3 * PI / 2, 2 * PI
        End If
        If Dur = 64 Then
            M_VSS.Circle Step(-m_LinSep * 0.4, m_LinSep * 1.25), m_LinSep * 0.25, vbBlack
            M_VSS.Circle Step(m_LinSep * 0.25, -m_LinSep * 0.25), m_LinSep * 0.25, vbBlack, 3 * PI / 2, 2 * PI
        End If
        '
        '
    End Select
    If dot > 0 Then
        M_VSS.FillStyle = vbFSSolid
        For i = 1 To dot
            M_VSS.Circle (NOffset + m_LinSep * i, m_TopLine + m_LinSep * Dur - m_LinSep / 6), m_LinSep / 6, vbBlack
        Next i
        M_VSS.FillStyle = vbFSTransparent
    End If

    M_VSS.FillStyle = vbFSTransparent

End Sub

Private Sub DrawSharp(X, Y)

  Dim Yoffset As Long

    M_VSS.DrawWidth = m_LineWidth * 2
    Yoffset = Y - m_LinSep / 2
    M_VSS.Line (X, Yoffset + 100)-Step(125, 0), vbBlack             'BotHorz
    M_VSS.Line (X + 20, Yoffset + 40)-Step(125, 0), vbBlack         'TopHorz
    M_VSS.Line (X + 70, Yoffset - 10)-Step(-40, m_LinSep), vbBlack 'LeftVert
    M_VSS.Line (X + 120, Yoffset - 10)-Step(-40, m_LinSep), vbBlack 'Right Vert
    M_VSS.DrawWidth = m_LineWidth

End Sub

Private Sub DrawStave()

    If m_ShowNokiaRange Then
        M_VSS.Line (0, m_TopLine + m_LinSep * 2)-Step(M_VSS.ScaleWidth, -m_LinSep * 3.5), RGB(250, 215, 255), BF
        M_VSS.Line (0, m_TopLine + m_LinSep * 5.5)-Step(M_VSS.ScaleWidth, -m_LinSep * 3.5), RGB(230, 215, 255), BF
        M_VSS.Line (0, m_TopLine + m_LinSep * 9)-Step(M_VSS.ScaleWidth, -m_LinSep * 3.5), RGB(200, 215, 255), BF
    End If
    DrawStaveLine 0, m_TopLine
    DrawStaveLine 0, m_TopLine + m_LinSep
    DrawStaveLine 0, m_TopLine + m_LinSep * 2
    DrawStaveLine 0, m_TopLine + m_LinSep * 3
    DrawStaveLine 0, m_TopLine + m_LinSep * 4
    G_Clef 400, m_TopLine
    M_VSS.Line (0, m_TopLine)-Step(0, m_LinSep * 4), vbBlack
    If m_BassShow Then
        DrawStaveLine 0, m_TopLine + m_LinSep * 6
        DrawStaveLine 0, m_TopLine + m_LinSep * 7
        DrawStaveLine 0, m_TopLine + m_LinSep * 8
        DrawStaveLine 0, m_TopLine + m_LinSep * 9
        DrawStaveLine 0, m_TopLine + m_LinSep * 10
        M_VSS.Line (0, m_TopLine)-Step(0, m_LinSep * 11), vbBlack
        F_Clef 400, m_TopLine + m_LinSep * 7
    End If
    NOffset = 800

End Sub

Private Sub DrawStaveLine(X1, Y, Optional X2 As Long = -1)

    M_VSS.DrawWidth = m_LineWidth
    If X2 = -1 Then
        X2 = M_VSS.ScaleWidth
    End If
    M_VSS.Line (X1, Y)-(X2, Y)

End Sub

Private Sub DrawTune(ByVal Note$, Optional OctOffsetB As Integer = 0)

  Dim sharp As Boolean, Flat As Boolean, dot As Integer, i As Integer
  Dim DBUG As String
  Dim LinSep As Long
  Dim EL As Long
  Dim OMemOffset As Long, OctOffset As Long, NPos As Long, SharpOffset As Long
  Dim spacer As Integer
  Dim PauseLength As Single, LocalLength As Single, BarLineDue As Boolean

    DBUG = Note
    If InStr(DBUG, "P") Then
        '        Stop
    End If

    If Left$(Note, 1) = "L" Then
        GetRemoveChar Note, True
        CurLength = GetRemoveNumerals(Note, True)
        'Exit Sub
    End If
    LocalLength = CurLength
    If Left$(Note, 1) = "O" Then
        GetRemoveChar Note, True
        CurOctave = GetRemoveNumerals(Note, True)
    End If
    If Left$(Note, 1) = "T" Then
        GetRemoveChar Note, True
        CurTempo = GetRemoveNumerals(Note, False)
        Exit Sub '>---> Bottom
    End If
    If Left$(Note, 1) = "P" Then
        If Note <> "P" Then
            PauseLength = GetRemoveNumerals(Note, False)
            LocalLength = PauseLength
        End If
    End If
    If Left$(Note, 1) = "M" Then
        Select Case Mid$(Note, 2)
          Case "N"
            BarlenFactor = 7 / 8
          Case "L"
            BarlenFactor = 1
          Case "S"
            BarlenFactor = 3 / 4
        End Select
        Exit Sub '>---> Bottom
    End If
    dot = CountChar(Note, ".")
    If dot > 0 Then
        Note = Replace(Note, ".", "")
    End If
    If Len(Note) = 0 Then
        Exit Sub '>---> Bottom
    End If

    sharp = InStr(Note, "#")
    If sharp Then
        Note = Replace(Note, "#", "")
    End If
    Flat = InStr(Note, "-")
    If Flat Then
        Note = Replace(Note, "-", "")
    End If
    OctOffset = (4 - CurOctave) * (m_LinSep * 4)

    If CurOctave <> 4 Then 'Octave 4 is ignored as it is the base set
        If CurOctave > 4 Then
            OctOffset = (4 - CurOctave) * (m_LinSep * 4) + m_LinSep / 2 * Abs(4 - CurOctave)
          Else 'NOT O...'NOT CUROCTAVE...
            OctOffset = (4 - CurOctave) * (m_LinSep * 4) - m_LinSep / 2 * Abs(4 - CurOctave)
        End If
    End If
    OMemOffset = (InStr("CDEFGAB", Note) * (m_LinSep * 0.5)) - 1 'OK
    'DrawBarLine LocalLength, dot, NPos, OMemOffset + OctOffset
    NPos = m_MidCpos - OMemOffset + OctOffset
    BarLineDue = TestDrawBarLine(LocalLength, dot, NPos)
    If Note = "P" Then
        DrawRest NOffset, LocalLength, dot
      Else 'NOT NOTE...
        DrawANote NOffset, NPos, LocalLength, dot, sharp, Flat, SharpOffset
    End If
    If Note <> "P" Then
        If NPos < m_TopLine Then
            For EL = m_TopLine To (NPos - m_LinSep / 2) Step -m_LinSep
                DrawStaveLine NOffset - 150, EL, NOffset + 150
            Next EL
        End If

        If NPos >= m_TopLine + m_LinSep * 5 And m_TopLine + m_LinSep * 5.4 > NPos Then
            For EL = m_TopLine + m_LinSep * 5 To NPos Step m_LinSep
                DrawStaveLine NOffset - 150, EL, NOffset + 150
            Next EL
        End If

        If m_TopLine + m_LinSep * 11 < NPos Then
            For EL = m_TopLine + m_LinSep * 11 To NPos Step m_LinSep
                DrawStaveLine NOffset - 150, EL, NOffset + 150
            Next EL
        End If
        If m_BassShow = False Then
            For EL = m_TopLine + m_LinSep * 5 To NPos Step m_LinSep
                DrawStaveLine NOffset - 150, EL, NOffset + 150
            Next EL
        End If
    End If

    spacer = (LocalLength / 64) * 100
    spacer = m_LinSep / spacer
    NOffset = NOffset + (spacer * m_LinSep * 0.5)
    If BarLineDue Then
        DrawBarLine
    End If

End Sub

Private Sub F_Clef(Xpos As Long, Ypos As Long)

    M_VSS.DrawWidth = m_LineWidth + 1
    M_VSS.Circle (Xpos, Ypos + m_LinSep * 2), m_LinSep * 2, vbBlack, PI / 2, 3 * PI / 2, 1
    M_VSS.Circle Step(0, -m_LinSep * 1), m_LinSep * 1, vbBlack, 3 * PI / 2, PI / 2, 1.5
    M_VSS.FillStyle = vbFSSolid
    M_VSS.Circle Step(0, m_LinSep * 0.5), m_LinSep * 0.5, vbBlack
    M_VSS.Circle Step(0, m_LinSep * 2), m_LinSep * 0.5, vbBlack, 3 * PI / 2, 2 * PI, 0.6
    M_VSS.FillStyle = vbFSTransparent
    M_VSS.DrawWidth = m_LineWidth

End Sub

Private Sub G_Clef(Xpos As Long, Ypos As Long)

  'M_VSS.Circle (Noffset, npos), m_LinSep / 3, vbBlack, , , 1

    M_VSS.DrawWidth = m_LineWidth + 1
    M_VSS.Circle (Xpos, Ypos - m_LinSep), m_LinSep, vbBlack, 3 * PI / 2, PI / 2, 1.5
    M_VSS.Line Step(0, -m_LinSep)-Step(0, m_LinSep * 7.5)
    M_VSS.Circle Step(0, -m_LinSep * 3.5), m_LinSep * 2, vbBlack, PI / 2, 3 * PI / 2, 1.5
    M_VSS.Circle Step(0, m_LinSep), m_LinSep * 1.25, vbBlack, 3 * PI / 2, PI / 2, 1.25
    M_VSS.Circle Step(0, -m_LinSep * 0.5), m_LinSep / 1.5, vbBlack, PI / 2, 3 * PI / 2, 1
    M_VSS.Circle Step(0, 0), m_LinSep / 2, vbBlack, -3 * PI / 2, PI / 2, 1
    M_VSS.Circle Step(0, m_LinSep * 2.5), m_LinSep / 2, vbBlack, PI / 2, 3 * PI / 2, 1
    M_VSS.DrawWidth = m_LineWidth

End Sub

Private Function GetRemoveChar(str As String, LTRF As Boolean, Optional CharNum As Long = 1) As String

  'Copyright 2002 Roger Gilchrist
  'take number of characters from Left True or right False of str and remove them from str

    If LTRF Then
        GetRemoveChar = Left$(str, CharNum)
        str = Mid$(str, 1 + CharNum)
      Else 'LTRF = FALSE
        GetRemoveChar = Right$(str, CharNum)
        str = Left$(str, Len(str) - CharNum)
    End If

End Function

Private Function GetRemoveNumerals(T$, LTRF As Boolean, Optional DecimalAware As Boolean = False) As Single

  'Remove numerals (optionally including decimal point) from string and return a count of them

  Dim tmp As String

    If LTRF Then
        Do While IsNumeric(Left$(T$, 1)) Or (Left$(T$, 1) = "." And DecimalAware)
            tmp = tmp & GetRemoveChar(T$, True)
        Loop

      Else 'LTRF = FALSE
        Do While IsNumeric(Right$(T$, 1)) Or (Right$(T$, 1) = "." And DecimalAware) 'numbers or If wanted decimal point
            tmp = GetRemoveChar(T$, False) & tmp
        Loop
    End If
    If Len(tmp) Then
        GetRemoveNumerals = CSng(tmp)
    End If

End Function

Private Function IncreaseNoteSpace(nvalue) As Integer

  Dim i As Long, LNoffset As Long

    IncreaseNoteSpace = nvalue
    'For i = 1 To nvalue
    'LNoffset = Noffset + (nvalue * m_LinSep)
    'BarLine = BarLine + 1
    'If BarLine Mod BeatsPerBar = 0 Then
    '    M_VSS.Line (LNoffset, m_TopLine)-Step(0, m_LinSep * 4), vbRed
    '    M_VSS.Line (LNoffset, m_TopLine + m_LinSep * 7)-Step(0, m_LinSep * 4), vbRed
    'End If
    'Next

End Function

Public Property Get LineSeperationSize() As Variant

    LineSeperationSize = m_LinSep

End Property

Public Property Let LineSeperationSize(ByVal vNewValue As Variant)

    m_LinSep = vNewValue

End Property

Public Property Get LineWidth() As Integer

    LineWidth = m_LineWidth

End Property

Public Property Let LineWidth(ByVal LWidth As Integer)

    m_LineWidth = LWidth

End Property

Public Sub Resize()

    ClearScore
    If Len(LastTune) Then
        AddNoteArray Split(LastTune)
    End If

End Sub

Public Property Get ShowBassStave() As Boolean

    ShowBassStave = m_BassShow

End Property

Public Property Let ShowBassStave(ByVal ShowIt As Boolean)

    m_BassShow = ShowIt

End Property

Public Property Get ShowNokiaRange() As Boolean

    ShowNokiaRange = m_ShowNokiaRange

End Property

Public Property Let ShowNokiaRange(ByVal ShowIt As Boolean)

    m_ShowNokiaRange = ShowIt
    Resize

End Property

Private Function TestDrawBarLine(D As Single, dot As Integer, NPos As Long) As Boolean

    M_VSS.DrawWidth = m_LineWidth + 1
    'If dot > 0 Then
    'D = D / ((3 / 2) ^ dot)
    'End If
    BarLine = BarLine + 1 / D

    If BarLine >= BarlenFactor Then
        BarLine = BarLine - BarlenFactor
        Bars = Bars + 1
        Avg_Barlen = (Avg_Barlen + NOffset) / Bars
        TestDrawBarLine = True

        '        If NOffset + Avg_Barlen > M_VSS.Width Then
        '            Avg_Barlen = 0
        '            Bars = 0
        '            BarLine = 0
        '            m_TopLine = m_TopLine + m_LinSep * IIf(m_BassShow, 16, 10)
        '            m_MidCpos = m_MidCpos + m_LinSep * IIf(m_BassShow, 16, 10)
        '           ' NPos = NPos + m_MidCpos 'move notes to next stave
        '            DrawStave
        '            NOffset = NOffset + m_LinSep
        '
        '        End If
    End If

End Function

Public Property Set VirtualScoreSheet(ByVal PBox As PictureBox)

    Set M_VSS = PBox
    M_VSS.AutoRedraw = True 'This is NOT to be changed; Keyboard will not display if you change this
    M_VSS.BorderStyle = 1   'This can be changed through the wrapper Property
    M_VSS.DrawWidth = 2
    ClearScore

End Property

':) Ulli's VB Code Formatter V2.13.6 (9/10/2002 11:45:58 AM) 31 + 598 = 629 Lines
