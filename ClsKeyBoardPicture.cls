VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsKeyBoardPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Copyright 2002 Roger Gilchrist
'This class grew out of playing around with Nokia Ringtone Player by Ovidiu Daniel Diaconescu
'It seemed logical that a keyboard rather than text was the way to input at least some of the data
' to write tunes, so I wrote this.
'All you really need is a PictureBox.

'(Put it in a Frame and the Keyboard will make use of the Frame caption but this is not essential.)
'The keyboard is fitted to the frame shape so a short and wide one is best
'Place in a Declare section
'       Public KB As New ClsKeyBoardPicture
'Place in program's start up point
'       KB.AssignControl PictureBox1
'Place the class's matching Sub in the following Subs of the PictureBox1
'PictureBox1_MouseDown
''              KB.MouseDown                    draw pressed key, play a note
'PictureBox1_MouseMove
''              KB.MouseMove                    Send messages to Frame(if it exists)
'PictureBox1_MouseUp
''              KB.MouseUp                     'Unpress button
''These use the private routines:
''Private MouseRead   private routine provides the data for the other Mouse events
''Private MouseDuration returns a duration base on button and shift

'You now have a playable keyboard
''
'To link to other controls call the relevant Public Properties (see Class outputs below)

'Set various colours in the control     Default
''KB.BackColor                          VbBlack
''KB.ColorBlackKey                      VbBlack
''KB.ColorKeyEdge                       Grey       Key Edges are drawn to give non-responsive areas to allow proper note detection
''KB.UseCKeyColourTags                  True       Turn on the following two colour uses
''KB.WholeKeyTag                        False      Tag is bottom 1/10 of key ; True Whole Key is tag colours
''KB.ColorCKey                          VbYellow   Mark any C key
''KB.ColorMidCKey                       VBRed      Mark out Middle C Key
''KB.ColorPressed                       Grey       Key colour changes to this while a note is playing

''Kb.KeyBoardLook                                   This routine allows you to select stored colour schemes
''                                                  rather than setting each element b(feel free to add to the set
''                                                  Remember to update the KBLook Enum
''
''KB.BorderStyle                                   'Wrapper for PictureBox1.BorderStyle 0 Flat 1 raised
''

''KB.OffSetEnds                         True        Draw keyBoard with small Left and Right borders of BackColor

''KB.PauseKeyOn                         False       Hide Pause key. If you only want to play with keyboard a pause is not useful
'                                                   True a Pause Key allows mouse input of pauses if you are sending note text to textboxes
''KB.ColorPause                         Purple      The Pause Key is drawn one key to the right of the rest of the keyboard and distinctly different
'
''KB.UseFormCaption                     False       Allows keyboard to use the Form's caption to display messages

'Class outputs      These allow you to see and write the classes values to TextBoxes

''KB.DurationVal     Length of Note
''KB.Octave          Octave number
''KB.OctaveMember     [#]<A-G>
''KB.TextOutPutStyle                     Basica    Set the format of NoteTextOutPut
''KB.NoteTextOutPut                                    Emit a Printable Code String (Basica or Nokia) to textbox
''KB.Tempo                                120      Used by the internal PlayKBNote .Call this Property separately if you need to emit it to TextBox

' Other Public members DeBug tools
''DEBUG ONLY  (DO NOT USE EXCEPT FOR DEBUG PURPOSES)
''KB.MouseX    ' show X position in Picturebox1
''KeyNumber     count of White keys reveals m_KeyNumber used as part of internal calculation
''KB.Frequency ' Show Frequency that would be used (in case you get the keys out of sync with the array)
''             ' Also makes a nice addition to the Frame Message

'Private members
''KeyBoardDraw                 Draw whole board whenever Colour or format or size changes
''Key                       Draw Key and its surrounding blackkeys
''GetRemoveRight            move Right of string to another str
''BasicaInvertSharp               reverse order of <A-G># to #<A-G>
''KeepInBoundsSimple        simple guard routine
''Playable                  a properly formed note exists
''PlayKBNote                play the note; calls to
''playbeep                  play sound
''playsleep                 play pause
''FrequencyFromNoteName     used by PlayKBNote to extract Frequency from Array
''GetDurationValue          convert string duration and dots to durationvalue for beep,sleep
'' See IMPORTANT notice below for more


Option Explicit
Private M_KB As PictureBox               'Internal name of PictureBox
Attribute M_KB.VB_VarHelpID = -1
Private M_KBFrame As Frame              'Optional Internal Name of Frame
Private M_KBForm As Form                 'Holds PictureBox.Parent
Private m_ShowOnFormCaption As Boolean   'Allow use of Form.Caption Area Default False
Private m_ContainerIsFrame As Boolean   'Test if Frame exists befor  referencing it

Private KeyWidth As Integer             'Key Layout Descriptors
Private KeyHeight As Integer
Private BKeyWidth As Integer
Private BKeyHeight As Integer
Private KeyTabtop As Integer
Private KeyTabHeight As Integer
Private Const Def_KeyTabTopDiv As Integer = 8   'Divisors for ColourTag affects bottom 10% of keys
Private Const Def_KeyTabHeightDiv As Integer = 9
Private ReSizing  As Boolean
Private m_IsBlack As Boolean            'Detect which coloured Key the mouse is over
Private M_IsWhite As Boolean            'based on detecting the colour with M_KB.Point(X,Y)
Private M_IsPause As Boolean            'NB these colours must be different
Private m_WhitePressed As Boolean       'M_IsXXX AND Button>0 sets these
Private m_LeftBlackPressed As Boolean
Private m_RightBlackPressed As Boolean

Private m_DurationSettings(31) As Integer

Private Enum ShiftDescriptor
    KNoShift
    KAlt
    KCtrl
    KShift
    KAltCtrl
    KShiftCtrl
    KShiftAlt
    KShiftAltCtrl
End Enum

Rem Mark Off
'Stops Code formatter complaining about these
#If False Then 'Enforce Case For Enums (does not compile but fools IDE)
Dim KNoShift
Dim KAlt
Dim KCtrl
Dim KShift
Dim KAltCtrl
Dim KShiftCtrl
Dim KShiftAlt
Dim KShiftAltCtrl
#End If  'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1
Rem Mark On

Public Enum ButtonShift
    ButtonL
    ButtonR
    ButtonM
    AltButtonL
    AltButtonR
    AltButtonM
    CtrlButtonL
    CtrlButtonR
    CtrlButtonM
    ShiftButtonL
    ShiftButtonR
    ShiftButtonM
    AltCtrlButtonL
    AltCtrlButtonR
    AltCtrlButtonM
    ShiftCtrlButtonL
    ShiftCtrlButtonR
    ShiftCtrlButtonM
    ShiftAltButtonL
    ShiftAltButtonR
    ShiftAltButtonM
    ShiftAltCtrlButtonL
    ShiftAltCtrlButtonR
    ShiftAltCtrlButtonM
End Enum
Rem Mark Off
'Stops Code formatter complaining about these
#If False Then 'Enforce Case For Enums (does not compile but fools IDE)
Dim ButtonL
Dim ButtonR
Dim ButtonM
Dim AltButtonL
Dim AltButtonR
Dim AltButtonM
Dim CtrlButtonL
Dim CtrlButtonR
Dim CtrlButtonM
Dim ShiftButtonL
Dim ShiftButtonR
Dim ShiftButtonM
Dim AltCtrlButtonL
Dim AltCtrlButtonR
Dim AltCtrlButtonM
Dim ShiftCtrlButtonL
Dim ShiftCtrlButtonR
Dim ShiftCtrlButtonM
Dim ShiftAltButtonL
Dim ShiftAltButtonR
Dim ShiftAltButtonM
Dim ShiftAltCtrlButtonL
Dim ShiftAltCtrlButtonR
Dim ShiftAltCtrlButtonM
#End If  'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1
Rem Mark On

Private m_Octave As Integer             'Store current octave
Private m_OctaveMember As Integer       'Store Current note name
Private PrevPressedOctave As Integer    'Store PreviousOctave while reading Pause key
Private m_KeyNumber As Long             'White key count; part of the determining which key you pressed(Not Directly meaningful)
Private m_Duration As Integer           'Store a duration value
Private m_UseColorTags As Boolean       'True =  colour tagged C keys
Private m_MouseX As Single              'Current X value of mouse cursor
Private m_Tempo As Integer              'Current Tempo Def 120
Private m_BasicaMFactor As Single       'Carries value of  MN=7/8(Default) ML=1 and MS=3/4 for translating Basica to Nokia
Private m_DottedNote As Boolean         'Add a dot to notation
Private m_BasicaClean As String         'Hold a cleaned version of Basica code
Private m_BasicaExpanded As String      'Hold an expanded version of Basica code
Private m_NokiaClean As String          'Hold a cleaned version of Nokia Composer
Private m_RTTTLClean As String          'Hold a cleaned version of RTTTL Nokia code
Private m_BasicaAsNokia As String       '
''Private m_PlayingNote As String       'DEBUG use only
'Private colour values                  'These colours MUST BE different to each other
Private m_C_Colour As Long              'as they play a role in determining which key was clicked
Private m_MidC_Colour As Long           'Just use RGB() and very small differences if you want a
Private m_PressedColour As Long         'Mono-chromatic look
Private m_EdgeColour As Long
Private m_BackColour As Long
Private m_BlackKeyColour As Long
Private m_WhiteKeyColour As Long
Private m_PauseKeyColour As Long
Private m_WholeKeyTag As Boolean
Private m_ShowNokiaOctave As Boolean
Private m_ColorHiNokiaOct As Long
Private m_ColorMidNokiaOct As Long
Private m_ColorLoNokiaOct As Long
Private m_NokiaDisableColour As Long
Private m_OffSetEnds As Boolean     'True = small border of BackColor on Left and right of keyboard
Private ButtonDur As String
Private KWLet As Variant '           'Stores A-G string values
Private KBLet As Variant             'Stores mathing sharps values

Private m_PauseKeyOn As Boolean      'True Show Pause Key False hide it(Keybard stretches to cover space created)

Private KeysDown As New Collection          'keeps track of pressed keys and releases them as needed (response to comment on version 1)
Private Const KeyCount As Integer = 60   'Number of WhiteKeys
Private Const PauseKeyNo As Integer = 62 'Send this value to key if using PauseKey
Private m_abort As Boolean              'allows you to interupt playing before tune is finished
Private m_NokiaAutoRepair As Boolean  'True warn not is out of range at each occurance False list wrong notes after playing
Private Type NoteValue            'Structure of NoteValueA array member
    Name As String
    Value As Single
End Type
Private NoteValueA() As NoteValue 'Array of Notes and Frequencies
Private m_LongestNoteValue As Single
Public Enum TextTypeOutPut ' Enum for NoteTextOutPut
    Basica
    Nokia
    RTTTL
End Enum
Private m_UseTxtOutPut As TextTypeOutPut 'Current TexpOutPut Style

Rem Mark Off
'Stops Code formatter complaining about these
#If False Then 'Enforce Case For Enums (does not compile but fools IDE)
Dim Basica
Dim Nokia
Dim RTTTL
#End If  'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1
Rem Mark On

Public Enum KBLook
    Antique
    Classical
    Default
    Random
End Enum

Rem Mark Off
'Stops Code formatter complaining about these
#If False Then 'Enforce Case For Enums (does not compile but fools IDE)
Dim Antique
Dim Classical
Dim Default
Dim Random
#End If  'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1
Rem Mark On

Private Enum CMEnum
    LargestMember
    LargestValue
    SmallestMember
    SmallestValue
    SmallestNotZeroMember
    SmallestNotZeroValue
End Enum
Rem Mark Off
'Stops Code formatter complaining about these
#If False Then 'Enforce Case For Enums (does not compile but fools IDE)
Dim LargestMember
Dim LargestValue
Dim SmallestMember
Dim SmallestValue
Dim SmallestNotZeroMember
Dim SmallestNotZeroValue
#End If  'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1
Rem Mark On
Private Enum RTTTLPartEnum
    RTTitle
    RTLen
    RTOct
    RTBeat
    RTNote
End Enum
Rem Mark Off
'Stops Code formatter complaining about these
#If False Then 'Enforce Case For Enums (does not compile but fools IDE)
Dim RTTitle
Dim RTLen
Dim RTOct
Dim RTBeat
Dim RTNote
#End If  'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1
Rem Mark On

Private m_RTTTLtitle As String
'declare the beep and sleep functions private
'for use only inside the class
Private Declare Function Beep Lib "kernel32" (ByVal dwFreq As Long, ByVal dwDuration As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'------------------IMPORTANT------------------------------------------------------
'As some people noticed the previous version did not support .9x versions of Windows
'a Quick search in PSC for 'PC Speaker' found me the stuff listed below
'I don't have access to a .9x system to test it so let me know if it doesn't work
'Thanks to Bob Richards for the code and permission to use it.
'I have copied and pasted his comments to help you use this.
'vvvvvvvvvvvvvvVERY IMPORTANTvvvvvvvvvvvvvvvvvvvvvvvvvvvvv
'You will need to download and install a DLL as described below.
'^^^^^^^^^^^^^^VERY IMPORTANT^^^^^^^^^^^^^^^^^^^^^^^^^^^^^
''I renamed the Bob's routines by adding BR to the start of their names
'so that Ulli's VB Code Formatter (V2.13.6 at time of coding)
'<<<DOWNLOAD AT PSC NOW!!! AND USE IT WHENEVER POSSIBLE!!!!!:)>>>
'would sort them next to each other for ease of reading code.
'Every thing from here (except some comments) to the bottom of
'Declarations section is part of Bob's code.
'If you don't want to support Win.9x then delete the support code
'(BRGetPlatform, BRWin9xBeep, BRPcSpeakerBeep)and from here to the
'bottom of Declarations section and make indicated change in PlayBeep routine)
'                                                                  Roger Gilchrist
'------------------IMPORTANT------------------------------------------------------

'Internal Speaker Beep Routine
'by Bob Richards
'adapted from Jorge Loubet's code for Win9x
'For Win9xMe systems, the file WIN95IO.DLL must be copied to the Windows/System folder.
'WIN95IO.DLL is available from http://www.softcircuits.com
''Routines used to add .9x support (slightly modified from Bob's versions see each for where)
''Sub BRPcSpeakerBeep Wrapper for either Beep method
''Sub BRWin9xBeep     .9x beep routine
''Function BRGetPlatform 'used in BRPcSpeakerBeep to select which method to call

Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Enum OsType
    Nt2000
    Win9xMe
    OsUnknown
End Enum
'I added this RG
Rem Mark Off
'Stops Code formatter complaining about these
#If False Then 'Enforce Case For Enums (does not compile but fools IDE)
Nt2000
Win9xMe
OsUnknown
#End If  'Barry Garvin VBPJ 101 Tech Tips 11 March 2001 p1
Rem Mark On

'This line is for NT2000 platforms
'Public Declare Function NtBeep Lib "kernel32" Alias "Beep" (ByVal FreqHz As Long, ByVal DurationMs As Long) As Long

'following lines are for Win9xMe platforms
'For these systems, the file WIN95IO.DLL must be copied
'to the Windows/System folder.
'WIN95IO.DLL is available from http://www.softcircuits.com
Private Declare Sub vbOut Lib "WIN95IO.DLL" (ByVal nPort As Integer, ByVal nData As Integer)
Private Declare Function vbInp Lib "WIN95IO.DLL" (ByVal nPort As Integer) As Integer

Public Sub Abort()

  'Copyright 2002 Roger Gilchrist
  'allows you to set the Abort value from
  'your program's stop playing button

    m_abort = True

End Sub

Public Sub About()

  Dim msg As String

    msg = "This Class is still too experimental to have stable documentation." & vbNewLine & _
          "Copyright 2002 Roger Gilchrist"
    MsgBox msg, vbInformation, "ClsKeyBoardPicture"

End Sub

Private Function ArrayRemoveBlankMembers(a As Variant, Optional Sep As String = "^") As Variant

  'Copyright 2002 Roger Gilchrist
  'remove blank members from any array
  'Sep allows you to change the seperator to one you know won't occur naturally in the data
  'I choose '^' as the default because it is not part of any strings used by this class
  'It doesn't matter what character you use as it will never get out of this routine.
  'For example the default " " for split is part of the array members for some sections of
  'this class so it would not be a good option for me to use.

  Dim TmpStr As String, i As Long

    For i = LBound(a) To UBound(a)
        If Len(a(i)) Then
            TmpStr = TmpStr & a(i) & Sep
        End If
    Next i
    TmpStr = Left$(TmpStr, Len(TmpStr) - 1)

    ArrayRemoveBlankMembers = Split(TmpStr, Sep)

End Function

Public Property Set AssignControl(ByVal PBox As PictureBox)

  'Copyright 2002 Roger Gilchrist
  'Set PBox and set up the keyboard to fit the PictureBox

    Set M_KB = PBox
    M_KB.AutoRedraw = True 'This is NOT to be changed; Keyboard will not display if you change this

    M_KB.BorderStyle = 1   'This can be changed through the wrapper Property

    'This is a non-essential pretty; If the PictureBox is in a frame
    'then set it up to display messages (See MouseMove for mechanics of this)
    'If you don't want a Frame you can set Property UseFormCaption=True
    'and the messages will appear on the from caption instead
    Set M_KBForm = M_KB.Parent
    If TypeOf M_KB.Container Is Frame Then
        Set M_KBFrame = M_KB.Container
        m_ContainerIsFrame = True
    End If
    KeyBoardLook Default 'setting the look automatical draws the keyboard

End Property

Public Property Let BackColor(ByVal Colr As Long)

    m_BackColour = Colr
    KeyBoardDraw

End Property

Public Property Get BackColor() As Long

    BackColor = m_BackColour

End Property

Private Function Basica2Nokia(ByVal n$) As String

  'Copyright 2002 Roger Gilchrist
  'Basica       "P<Duration>"
  'Basica Note Format "L<Duration> O<Octave> [#]<Note>[...]"

  'Nokia format <Duration>-
  'Nokia Note Format <Duration>[...][#]<Note><Octave>

  Dim DurationValue As Single, DotValue As Long, Oct As Long
  Dim TmpN As String

    If Right$(n$, 1) = "P" Then
        'n = Mid$(n$, 2)
        DurationValue = GetRemoveNumerals(n, True) ' * m_BasicaMFactor
        DotValue = RemoveDots(n)
        Basica2Nokia = DurationValue & String$(DotValue, ".") & "P"
      Else 'NOT LEFT$(N$,...'NOT RIGHT$(N$,...
        'L16O4C
        n$ = Replace(n$, "L", "")
        DurationValue = GetRemoveNumerals(n$, True) '* m_BasicaMFactor
        n$ = Replace(n$, "O", "")
        Oct = GetRemoveNumerals(n$, False) - 2  '- 4  'get the Octave value and '+ 4'
        DotValue = RemoveDots(n$)
        If Right$(n$, 1) = "#" Then
            n = BasicaInvertSharp(n)
        End If
        Basica2Nokia = Int(DurationValue) & String$(DotValue, ".") & n & CInt(Oct)
    End If

End Function

Public Function Basica2RTTTL(n$) As String

  Dim TmpVal As Integer, Tmpval2 As Integer, DotValue As Integer

    TmpVal = GetRemoveNumerals(n, False)
    If InStr(n, "#") Then
        Tmpval2 = GetRemoveNumerals(n, True)
        DotValue = RemoveDots(n$)
        n = BasicaInvertSharp(n)
        n = Tmpval2 & n & String$(DotValue, ".")
    End If
    If TmpVal > 0 Then
        Basica2RTTTL = n & TmpVal + 1 'octave shift
      Else 'NOT TMPVAL...
        Basica2RTTTL = n
    End If

    'd=duration o=octave b=beats per min n=note
    'Title:d= ,o= ,b= : [d]n[#][.][o],[d]p,

    'm_BasicaExpanded

End Function

Private Function BasicaAlphaNotes(str As String, Char As String) As String

  'Copyright 2002 Roger Gilchrist
  'Deal with the standard 'CDEFGAB' and possible '#+-'

    If InStr("#+-", Left$(str, 1)) > 0 And Len(str) > 0 Then 'has sharp or flat
        BasicaAlphaNotes = BasicaAlphaNoteShift(Left$(str, 1), Char, False)
        str = Mid$(str, 2) 'trim from main str
      Else 'NOT INSTR("#+-",...
        BasicaAlphaNotes = Char
    End If

End Function

Private Function BasicaAlphaNoteShift(Symb$, n$, Flat2Sharp As Boolean) As String

  'Copyright 2002 Roger Gilchrist
  'This only hits if there is a half-note symbol so there is no checking for other sybmols
  'Basica allows '#' or '+' for sharp half-notes and '-' for flat half-notes Nokia only uses '#'
  'For ease of use this routine does not tell you if the note is illegal
  'it just deletes the half-note symbol
  'Flat2Sharp =False Delete illegals #- convert + to #
  'Flat2Sharp =True Delete illegals and convert Legal Flats to <Note-1>#

    If Symb = "-" Then 'convert flats to previous KEY#
        If Flat2Sharp Then
            If InStr("DEGAB", n$) Then ' 'legal flats
                'This would make any flat into note-1#

                If n$ = "A" Then
                    n$ = "G"
                  Else 'NOT N$...
                    n$ = Chr$(Asc(n$) - 1)
                End If
                Symb = "#" 'set to Nokia's only half-note mark
              Else 'NOT INSTR("DEGAB",...
                Symb = "" ' delete illegal flats
            End If
          Else 'FLAT2SHARP = FALSE
            If InStr("DEGAB", n$) Then ' 'legal flats
                Symb = "-"
              Else 'NOT INSTR("DEGAB",...
                Symb = "" ' delete illegal flats
            End If
        End If
      ElseIf Symb = "+" Then 'reset '+" to "#"'NOT SYMB...
        Symb = "#" 'set to Nokia's only half-note mark
    End If

    If InStr("CDFGA", n$) > 0 And Len(Symb) = 1 Then 'legal sharps
        BasicaAlphaNoteShift = n$ & Symb
      Else 'NOT INSTR("CDFGA",...
        BasicaAlphaNoteShift = n$ 'Delete # for notes which don't have it
    End If

End Function

Private Function BasicaClean(tsong As String) As String

  'Copyright 2002 Roger Gilchrist
  'Phase 1 of BasicaTranslate; separates the string into discrete code commands
  'This routine transforms most of the old Basic/Basica Play strings into a space delimited version
  'but disposes of some parts which are no longer supported:

  Dim CurChar As String
  Dim PrevChar As String
  Dim NextChar As String
  Dim i As Integer
  Dim CurLValue As Integer
  Dim TmpCurLength As Integer
  Dim TmpCurVal As Integer
  Dim TmpLenVal As Long
  Dim TmpDotVal As Long
  Dim TmpChar As String 'use to allow extration of value from 'O#' for '>' and '<'
  Dim TmpChar2 As String 'used by n<NoteNUMBER> note translator
  Dim PrevOctave As Integer
  Dim RestoreOct  As Boolean
  Dim OneNoteLen As Single
  Dim Tmpa  As Variant, TmpLValue As Integer, j As Integer
  Dim dotCount As Long, showme As String, showme2 As String
  Dim JunkVal As Integer 'Value returned is not needed

    On Error Resume Next
        PrevOctave = 4 'Default value used by '>' and '<' to rewrite to "O#"
        CurLValue = 1
        tsong = Replace(UCase$(tsong), " ", "") 'Force to Ucase( simplifies decoding) and junk all spaces they are not needed

        BasicaInitialDefaults tsong
        Do While Len(tsong)
            CurChar = Left$(tsong, 1)
            tsong = Mid$(tsong, 2)
            NextChar = Left$(tsong, 1)
            If CurChar = "F" Then
                'Stop
            End If
            Select Case CurChar

              Case "A", "B", "C", "D", "E", "F", "G"
                'Standard Notes: check for '+#-' and convert flats to prev note sharps

                TmpChar = BasicaAlphaNotes(tsong, CurChar) 'this takes care of any Flat or Sharp symbols
                TmpLenVal = 0
                TmpDotVal = 0
                NextChar = Left$(tsong, 1)
                If IsNumeric(NextChar) Then
                    OneNoteLen = GetRemoveNumerals(tsong, True)  'OneNoteLen is used as trigger later so don't combine with next lone
                    TmpLenVal = OneNoteLen
                    NextChar = Left$(tsong, 1)
                End If

                Do While NextChar = "."
                    tsong = Mid$(tsong, 2)
                    NextChar = Left$(tsong, 1)
                    TmpDotVal = TmpDotVal + 1
                Loop

                If TmpLenVal > 0 Then ' add single note length value
                    BasicaClean = BasicaClean & " L" & TmpLenVal
                End If

                BasicaClean = BasicaClean & " " & TmpChar & String$(TmpDotVal, ".")

                If OneNoteLen > 0 Or TmpDotVal > 0 Then 'Restore general length if needed
                    BasicaClean = BasicaClean & " L" & CurLValue
                End If

                If RestoreOct Then
                    BasicaClean = BasicaClean & " " & "O" & PrevOctave
                    RestoreOct = False
                End If

              Case "N" 'Numbered Notes
                TmpChar = NoteValueA(Val(Mid$(BasicaGetRemoveNumberWithLimits(tsong, CurChar, 0, 84, JunkVal), 2))).Name
                TmpChar2 = GetRemoveNumerals(TmpChar, False)
                BasicaClean = BasicaClean & " O" & TmpChar2 & " " & TmpChar
              Case "L" 'Length
                BasicaClean = BasicaClean & " " & BasicaGetRemoveNumberWithLimits(tsong, CurChar, 1, 64, CurLValue)
              Case "P" ' Pause
                BasicaClean = BasicaClean & " " & BasicaGetRemoveNumberWithLimits(tsong, CurChar, 1, 64, JunkVal)
              Case "O" 'Octave
                BasicaClean = BasicaClean & " " & BasicaGetRemoveNumberWithLimits(tsong, CurChar, 0, 6, PrevOctave)
              Case "T" 'Tempo
                BasicaClean = BasicaClean & " " & BasicaGetRemoveNumberWithLimits(tsong, CurChar, 25, 900, JunkVal)
              Case "M"
                If InStr("FB", Left$(tsong, 1)) Then
                    'throw away 'M" and next char as it is not not supported
                    'and you don't want the 'F' or 'B' becoming notes
                    GetRemoveChar tsong, True
                  Else 'NOT INSTR("FB",...
                    'if someone comes up with an asynchronous way of using API Beep then
                    'delete this 'If Else End If' structure except for the next line
                    BasicaClean = BasicaClean & " M" & GetRemoveChar(tsong, True)
                End If
              Case "."
                BasicaClean = BasicaClean & "."
                tsong = Mid$(tsong, 2)
                NextChar = Left$(tsong, 1)
              Case " "
                'Skip spaces
              Case "<", ">"
                BasicaClean = BasicaClean & " " & "O" & CStr(PrevOctave + IIf(CurChar = ">", 1, -1))
                RestoreOct = True
            End Select
            PrevChar = CurChar
        Loop
        Tmpa = Split(BasicaClean)
        CurLValue = 1
        Tmpa = BasicaCleanDuplicates(Tmpa)
        BasicaClean = Join(Tmpa)
    On Error GoTo 0
    'a

End Function

Private Function BasicaCleanDuplicates(Tmpa As Variant) As Variant

  'Copyright 2002 Roger Gilchrist
  'Remove duplicate and unneeded calls to Basica 'L' and 'O' commands generated by automatic rewriting routines

  Dim i As Long

    '   Delete first of pairs of L and O commands only the last one matters
    For i = LBound(Tmpa) To UBound(Tmpa) - 1
        If Left$(Tmpa(i), 1) = "L" And Left$(Tmpa(i + 1), 1) = "L" Then
            Tmpa(i) = "" ' delete first of successive pairs of command
        End If
        If Left$(Tmpa(i), 1) = "O" And Left$(Tmpa(i + 1), 1) = "O" Then
            Tmpa(i) = ""
        End If
    Next i
    'Delete meaningless last code elements
    '( do after previous test in case last two are meaningless)
    If Left$(Tmpa(UBound(Tmpa)), 1) = "L" Then
        Tmpa(UBound(Tmpa)) = ""
    End If
    If Left$(Tmpa(UBound(Tmpa)), 1) = "O" Then
        Tmpa(UBound(Tmpa)) = ""
    End If
    'Basica does not need repetative calls to L or O
    'So delete succesive duplicates
    BasicaDuplicateDeleter Tmpa, "L"
    BasicaDuplicateDeleter Tmpa, "O"

    BasicaCleanDuplicates = ArrayRemoveBlankMembers(Tmpa)

End Function

Private Sub BasicaDuplicateDeleter(Tmpa As Variant, IntStr As String)

  'Copyright 2002 Roger Gilchrist
  'called from BasicaCleanDuplicates
  'Cleans repetative calls to L or O
  'Search for IntStr starting an array member
  'and delete successive calls to the same value IntStr value

  Dim i As Long, j As Long
  Dim Test As String

    For i = LBound(Tmpa) To UBound(Tmpa)
        If Left$(Tmpa(i), 1) = IntStr Then
            Test = Tmpa(i)
            For j = i + 1 To UBound(Tmpa)
                If Left$(Tmpa(j), 1) = IntStr Then
                    If Tmpa(j) = Test Then
                        Tmpa(j) = ""
                      Else 'NOT TMPA(J)...
                        Exit For '>---> Next
                    End If
                End If
            Next j
        End If
    Next i

End Sub

'Private Function BasicaConvert2CB(N) As String
'
'    BasicaConvert2CB = NoteValueA(CInt(Replace(N, "", ""))).Name
'
'End Function

Private Function BasicaGetFirstNotePos(tsong As String) As Long

  Dim FindNote As Long, FirstNote As Long, i As Long

    BasicaGetFirstNotePos = Len(tsong) 'Get the first notes position
    For i = 0 To 7
        FindNote = InStr(tsong, Chr$(Asc("A") + i))
        If FindNote > 0 Then
            If FindNote < BasicaGetFirstNotePos Then
                BasicaGetFirstNotePos = FindNote
            End If
        End If
    Next i

End Function

Private Function BasicaGetRemoveNumberWithLimits(Song As String, Char As String, Min As Integer, MAx As Integer, Value As Integer) As String

  'Copyright 2002 Roger Gilchrist
  'deal with code letters O(ctave),N(ote),L(ength),P(ause),T(empo)
  'which take a numeric argument
  'Min and Max are the values I got from Basic 2nd Edition(Revised January 1983) Version 1.10 IBM

  Dim Num As String

    Num = GetRemoveNumerals(Song, True)
    Value = Val(Num)
    KeepInBounds Min, Value, MAx
    BasicaGetRemoveNumberWithLimits = Char & Value

End Function

Private Sub BasicaInitialDefaults(tsong As String)

  'Copyright 2002 Roger Gilchrist
  'this guards against user entering codes without a Length, Tempo or octave values before a Note is played

  Dim FindNote As Long, FirstNote As Long, TmpStr As String, TmpVal As Long

    'Length guard
    If InStr(tsong, "O") = 0 Then 'none present so set a default octave
        tsong = "O4 " & tsong
      Else 'NOT INSTR(TSONG,...
        FirstNote = BasicaGetFirstNotePos(tsong) 'Restest in case previsous test changed things
        If InStr(tsong, "O") > FirstNote Then
            TmpStr = Mid$(tsong, InStr(tsong, "O") + 1)
            TmpVal = GetRemoveNumerals(TmpStr, True)
            tsong = "O" & TmpVal & " " & tsong
        End If
    End If

    If InStr(tsong, "L") = 0 Then 'none present so set a default note length
        tsong = "L8 " & tsong
      Else 'NOT INSTR(TSONG,...
        'Exists but may not be before first note so check
        FirstNote = BasicaGetFirstNotePos(tsong) 'Restest in case previsous test changed things
        If InStr(tsong, "L") > FirstNote Then ' if a note is before the Length then copy the first Length to start of string
            TmpStr = Mid$(tsong, InStr(tsong, "L") + 1)
            TmpVal = GetRemoveNumerals(TmpStr, True)
            tsong = "L" & TmpVal & " " & tsong
        End If
    End If

    'Tempo guard same general structure as Lenght
    If InStr(tsong, "T") = 0 Then ' safety set a default Tempo
        tsong = "T" & m_Tempo & tsong ' get the tempo from the Nokia combobox
      Else 'NOT INSTR(TSONG,...
        FirstNote = BasicaGetFirstNotePos(tsong) 'Restest in case previsous test changed things
        If InStr(tsong, "T") > FirstNote Then
            TmpStr = Mid$(tsong, InStr(tsong, "T") + 1)
            TmpVal = GetRemoveNumerals(TmpStr, True)
            tsong = "T" & TmpVal & " " & tsong
        End If
    End If

End Sub

Private Function BasicaInvertSharp(str As String) As String

  'Reverse the order of <A-G># for switch between Nokia and Basica

    If InStr(str, "#") Then
        BasicaInvertSharp = Mid$(str, 2) & Left$(str, 1)
      Else 'NOT INSTR(STR,...
        BasicaInvertSharp = str
    End If

End Function

Public Function BasicaNoteShift(ByVal T$, Dir As Integer) As String

  'Copyright 2002 Roger Gilchrist
  'Find a notevalue and increase or decrease its value
  'UpDated to more Complex shifter

    T$ = UCase$(T)
    BasicaNoteShifter T$, Dir
    BasicaNoteShift = T$

End Function

Private Sub BasicaNoteShifter(Note$, Dir As Integer)

  'Copyright 2002 Roger Gilchrist
  'Move notes 1 note up(dir=1) or down(dir=-1) scale
  'Up Natural becomes sharp,Sharp becomes Flat of next note(No sound change), Flat becomes Natural
  'Down Natural becomes Flat, Sharp becomes Natural, Flat becomes Sharp of previous note(No sound change)
  'If note is at Octave boundry, Octave is also changed for just that note

  Dim TmpChar As String, TmpStr As String, Junk As String
  Dim CurOctave As Integer, DoOctaveShift As Integer

    Do
        DoOctaveShift = 0
        TmpChar = GetRemoveChar(Note, True)
        If TmpChar = "O" Then
            CurOctave = GetRemoveNumerals(Note, True)
            TmpChar = "O" & CurOctave
        End If
        If InStr("CDEFGAB", TmpChar) Then
            If TmpChar = "C" And Dir = -1 Then
                DoOctaveShift = -1
            End If
            If TmpChar = "B" And Dir = 1 Then
                DoOctaveShift = 1
            End If

            Select Case Left$(Note, 1)
              Case "#"
                If Dir = -1 Then
                    Junk = GetRemoveChar(Note, True) 'just remove sharp
                  Else 'NOT DIR...
                    Junk = GetRemoveChar(Note, True) 'switch to flat of next note
                    Note = "-" & Note
                    TmpChar = RotateNote(TmpChar, Dir)
                End If
              Case "-"

                If Dir = -1 Then
                    Junk = GetRemoveChar(Note, True) 'switch to sharp of next note
                    Note = "#" & Note
                    TmpChar = RotateNote(TmpChar, Dir)
                  Else 'NOT DIR...
                    Junk = GetRemoveChar(Note, True) ''just remove the flat
                End If

              Case Else

                If Dir = 1 Then
                    If InStr("CDFGA", TmpChar$) Then
                        TmpChar = TmpChar & "#"
                      Else 'NOT INSTR("CDFGA",...
                        TmpChar = RotateNote(TmpChar, Dir)
                    End If
                End If
                If Dir = -1 Then
                    If InStr("DEGAB", TmpChar$) Then
                        TmpChar = TmpChar & "-"
                      Else 'NOT INSTR("DEGAB",...
                        TmpChar = RotateNote(TmpChar, Dir)
                    End If
                End If
            End Select
        End If
        If DoOctaveShift <> 0 Then
            TmpChar = "O" & CurOctave + DoOctaveShift & TmpChar & "O" & CurOctave
        End If
        TmpStr = TmpStr & TmpChar
    Loop While Len(Note)
    Note = TmpStr

End Sub

Public Function BasicaOctaveShift(ByVal T$, Dir As Integer) As String

  'Copyright 2002 Roger Gilchrist
  'Apologies to Ulli's Code formatter
  'Find an octave code and increase or decrease its value

  Dim i As Long, TmpVal As String, Newvalue As Integer

    For i = 1 To Len(T)
        If Mid$(T, i, 1) = "O" Then
            i = i + 1  'Sorry Ulli, I know it's naughty but this is safe':( Modifies active For-Variable
            TmpVal = ""
            Do While IsNumeric(Mid$(T, i, 1))
                TmpVal = TmpVal & Mid$(T, i, 1)
                i = i + 1   'Sorry Ulli, I know it's naughty but this is safe':( Modifies active For-Variable
            Loop
            Newvalue = Val(TmpVal) + Dir
            KeepInBounds 0, Newvalue, 6
            BasicaOctaveShift = BasicaOctaveShift & "O" & Newvalue & " "
          Else 'NOT MID$(T,...
            BasicaOctaveShift = BasicaOctaveShift & Mid$(T, i, 1)
        End If
    Next i

End Function

Public Function BasicaRandomLength() As Integer

  'keep Length to basic note lengths rather than fully random

  Dim rlen As Integer

    Select Case Int(Rnd * 8)
      Case 0
        rlen = 1
      Case 1
        rlen = 2
      Case 2
        rlen = 4
      Case 3
        rlen = 8
      Case 4
        rlen = 16
      Case 5
        rlen = 32
      Case 6
        rlen = 64
      Case Else
        rlen = Int(Rnd * 64 + 1)
    End Select
    BasicaRandomLength = rlen

End Function

Public Function BasicaRandomTune() As String

  'Copyright 2002 Roger Gilchrist
  'Based on 'Table 6-1 The Probability Distribution for the Random Selection of Notes'
  '1001 Things To Do With Y+our IBM PC' by Mark R. Sawusch and Tan A. Summers
  'This uses that table to produce slightly better than completely random music
  'Uses the Basica "N<NoteNumber>" code to simplify getting notes ('N' notes contain their own octave value)
  'Tempo is between 90 and 130
  'Length is kept to basic note lengths rather than fully random

  Dim i As Long, NRnd As Integer, Dir As Integer, CurLen As Integer
  Dim PrevN As Integer, SpeedShift As Single, PrevLen As Integer, InitSpeed As Single

    'Initialise Initial values
    Dir = IIf(Rnd > 0.5, 1, -1) '50% chance of initial direction to move
    SpeedShift = Rnd            'Probabaility of changing Length value
    InitSpeed = SpeedShift      'Store Def speed
    CurLen = BasicaRandomLength
    NRnd = 24 + Int(Rnd * 12)
    BasicaRandomTune = "ML T" & Int(Rnd * 40) + 90 & " L" & CurLen
    For i = 1 To 30
        BasicaRandomTune = BasicaRandomTune & "N" & NRnd
        'Change or maintain SpeedShift
        If PrevN = NRnd Or CurLen > 16 Or CurLen = PrevLen Then
            SpeedShift = SpeedShift - 0.1
          Else 'NOT PREVN...
            SpeedShift = SpeedShift + 0.1
        End If
        KeepInBounds 0, SpeedShift, 1
        If SpeedShift = 1 Or SpeedShift = 0 Then
            SpeedShift = InitSpeed
        End If

        If Rnd > SpeedShift Then 'change note length
            PrevLen = CurLen
            CurLen = BasicaRandomLength
            BasicaRandomTune = BasicaRandomTune & "L" & CurLen
        End If

        If Rnd > 0.5 Then ' 50% chance of changing direction at all
            Dir = IIf(Rnd > 0.5, 1, -1) '50% chance of reversing
        End If
        PrevN = NRnd 'Store prev note value
        Select Case Int(Rnd * 100) 'percentage decides how far to move to next note
          Case Is <= 4
            NRnd = NRnd ' same note
          Case Is > 4, Is <= 18
            NRnd = NRnd + Dir
          Case Is > 18, Is <= 30
            NRnd = NRnd + (Dir * 2)
          Case Is > 30, Is <= 40
            NRnd = NRnd + (Dir * 3)
          Case Is > 40, Is <= 48
            NRnd = NRnd + (Dir * 4)
          Case Is > 48, Is <= 60
            NRnd = NRnd + (Dir * 5)
          Case Is > 60, Is <= 62
            NRnd = NRnd + (Dir * 6)
          Case Is > 62, Is <= 74
            NRnd = NRnd + (Dir * 7)
          Case Is > 74, Is <= 80
            NRnd = NRnd + (Dir * 8)
          Case Is > 80, Is <= 86
            NRnd = NRnd + (Dir * 9)
          Case Is > 86, Is <= 90
            NRnd = NRnd + (Dir * 10)
          Case Is > 90, Is <= 94
            NRnd = NRnd + (Dir * 11)
          Case Is > 94
            NRnd = NRnd + (Dir * 12)
          Case Is > 99
            NRnd = NRnd + (Dir * Int(Rnd * 12)) '1% chance of totally random move
        End Select
        KeepInBounds 1, NRnd, 84

    Next i
    BasicaRandomTune = BasicaClean(BasicaRandomTune)

End Function

Public Sub BasicaRead(str As String)

  'Copyright 2002 Roger Gilchrist
  'convert simplifed Basica to Nokia Code format
  'Basica       "P<Duration>"
  'Basica Note Format "L<Duration> O<Octave> [#]<Note>[...]"

  Dim Notes As Variant, i As Integer
  Dim Dur As Integer, n As String, Oct As String
  Dim tran As String

    'ClearScore
    If Len(Trim$(str)) = 0 Then
        Exit Sub '>---> Bottom
    End If
    Notes = Split(BasicaTranslate(str))

    m_NokiaClean = ""
    m_RTTTLClean = ""
    RTTTLWrite RTBeat, CStr(m_Tempo)
    m_abort = False
    m_BasicaMFactor = 7 / 8
    For i = LBound(Notes) To UBound(Notes)
        If InStr("MT", Left$(Notes(i), 1)) Then
            If Left$(Notes(i), 1) = "T" Then
                m_Tempo = GetRemoveNumerals(CStr(Notes(i)), False)
              Else 'NOT LEFT$(NOTES(I),...
                Select Case GetRemoveChar(CStr(Notes(i)), False)
                  Case "N"
                    m_BasicaMFactor = 7 / 8

                  Case "L"
                    m_BasicaMFactor = 1

                  Case "S"
                    m_BasicaMFactor = 3 / 4

                End Select
            End If
          Else 'NOT INSTR("MT",...
            'DrawTune Notes(i), 2 ', 1 ', 2
            tran = Basica2Nokia(CStr(Notes(i)))

            ShowMsg "Player Piano Mode " & Format$(Percent(i, UBound(Notes)), "###") & "% " & Notes(i) & "=>" & tran
            m_NokiaClean = m_NokiaClean & " " & tran
            RTTTLWrite RTNote, Basica2RTTTL(CStr(Notes(i)))
            PlayKBNote CStr(Notes(i)), True, True
            '            PlayKBNote CStr(tran), True, True
            If m_abort Then
                Exit For '>---> Next
            End If
        End If
    Next i
    ShowMsg ""
    RTTTLClean
    NokiaClean m_NokiaClean

End Sub

Private Function BasicaTranslate(ByVal tsong As String) As String

  'Copyright 2002 Roger Gilchrist
  'converts Baisca string into Nokia style code

  Dim t1 As Variant, t2 As Variant, ConvertBasicToNokiB As String 'DEBUG TEST STUFF
  Dim i As Long ' For..Next counter
  Dim Octave() As Integer 'Store an octave value for each note
  Dim CurOctave As Integer 'Hold current octave value for Octave()
  Dim TmpOctave As Integer 'Hold CurOvtave while '<' and '>' code is dealt with
  Dim LenNote() As Integer 'Store an Length value for each note
  Dim CurLengthN As Integer 'Hold Current Length value for LenNote()
  Dim CurMEffect As Integer 'Hold Current Value of 'M'  MN=7/8,ML=1,MS=3/4'Defualt 1
  Dim N_NoteValue As Integer 'Hold value of notes using the 'N#' codeing
  Dim Tmpa As Variant  'Temporary Array used by secnd phase of translation. Each member is one command
  Dim TmpStr As String
  Dim ResetOneTimeOctave As Integer, ResetOctaveActivated As Boolean

    'Phase 1 separate the string into discrete code commands
    m_BasicaClean = BasicaClean(tsong)
    'NB m_cleanBasic can be read by this Routine

    'Nokia format
    'PHASE 2 convert the discrete strings to their Nokia code equivalent

    'Rewrite sharp, flat and pause to Nokia format
    Tmpa = Split(Trim$(m_BasicaClean))
    For i = LBound(Tmpa) To UBound(Tmpa)
        If Left$(Tmpa(i), 1) = "P" Then
            TmpStr = Mid$(Tmpa(i), 2)
            Tmpa(i) = TmpStr & "P"
        End If
        If Right$(Tmpa(i), 1) = "#" Then
            Tmpa(i) = StrReverse(Tmpa(i))
        End If
        If Right$(Tmpa(i), 1) = "-" Then
            Tmpa(i) = StrReverse(BasicaAlphaNoteShift("-", Left$(Tmpa(i), 1), True))
        End If
    Next i

    'Rewrite  O(ctave) Data
    ReDim Octave(UBound(Tmpa)) As Integer
    CurOctave = 4
    For i = LBound(Tmpa) To UBound(Tmpa)
        If Left$(Tmpa(i), 1) = "O" Then
            Select Case Mid$(Tmpa(i), 2)
              Case "+1"
                CurOctave = CurOctave + 1
                ResetOneTimeOctave = 1
                ResetOctaveActivated = True
              Case "-1"
                CurOctave = CurOctave - 1
                ResetOneTimeOctave = -1
                ResetOctaveActivated = True
              Case Else
                CurOctave = CInt(Mid$(Tmpa(i), 2))
            End Select
            KeepInBounds 0, CurOctave, 10 '7'TESTING
            Tmpa(i) = "" ' delete it
        End If

        If Not ResetOctaveActivated Then ' dont do it immediately it has to apply to next member
            CurOctave = CurOctave + ResetOneTimeOctave
        End If
        Octave(i) = CurOctave
        ResetOctaveActivated = False
        ResetOneTimeOctave = 0
    Next i

    '    For i = LBound(tmpA) To UBound(tmpA)
    '        If Left$(tmpA(i), 1) = "M" Then
    '            tmpA(i) = "" ' delete it
    '        End If
    '    Next i

    ReDim LenNote(UBound(Tmpa)) As Integer
    CurLengthN = 64 'Defualt value
    CurMEffect = 1 'Defualt value
    For i = LBound(Tmpa) To UBound(Tmpa)
        If Left$(Tmpa(i), 1) = "L" Then
            CurLengthN = CInt(Mid$(Tmpa(i), 2))
            Tmpa(i) = "" ' delete it
            KeepInBounds 1, CurLengthN, 64
        End If
        If Left$(Tmpa(i), 1) = "M" Then
            Select Case Mid$(Tmpa(i), 2)
              Case "N"
                CurMEffect = 7 / 8
              Case "L"
                CurMEffect = 1
              Case "S"
                CurMEffect = 3 / 4
              Case "F", "B" 'no longer supported just deletes them
            End Select
            'tmpA(i) = "" ' delete it
        End If

        LenNote(i) = CurLengthN * CurMEffect
    Next i

    For i = LBound(Tmpa) To UBound(Tmpa) 'convert N## note to standard note
        If Left$(Tmpa(i), 1) = "N" Then
            N_NoteValue = CLng(Mid$(Tmpa(i), 2))
            KeepInBounds 1, N_NoteValue, UBound(NoteValueA)
            Tmpa(i) = NoteValueA(N_NoteValue).Name
            If InStr(Tmpa(i), "#") Then
                Tmpa(i) = "#" & Replace(Tmpa(i), "#", "")
            End If
            Tmpa(i) = Tmpa(i) & "*"  'add a blocker so that the octave is not added
        End If
    Next i

    For i = LBound(Tmpa) To UBound(Tmpa)
        If Len(Tmpa(i)) Then
            If Right$(Tmpa(i), 1) = "*" Then 'it was in N# form and already has an octave atacheded
                Tmpa(i) = Left$(Tmpa(i), Len(Tmpa(i)) - 1)
              Else 'NOT RIGHT$(TMPA(I),...
                If Not (InstrArray(Tmpa(i), "P", "M", "T")) Then
                    Tmpa(i) = LenNote(i) & Tmpa(i) & Octave(i)
                End If
            End If
        End If
    Next i

    For i = LBound(Tmpa) To UBound(Tmpa)
        If Len(Tmpa(i)) Then
            BasicaTranslate = BasicaTranslate & " " & Tmpa(i)
        End If
    Next i
    BasicaTranslate = Trim$(BasicaTranslate)
    m_BasicaExpanded = BasicaTranslate
    'DEBUG STUFF send BASICA MISSION IMPOSSIBLE  AND TEST AGAINST THE WORKING NOKIA CODE
    '    ConvertBasicToNokiB = "16g2   8- 16g2 8-  16f2 16-  16#f2 16-  16g2 8-  16g2 8- 16#a2 16- 16c3 16- 16g2 8- 16g2 8- 16f2 16- 16#f2 16- 16g2 8- 16g2 8- 16#a2 16- 16c3 16- 16#a2 16g2 2d2 32- 16#a2 16g2 2#c2 32- 16#a2 16g2 2c2 16- 16#a1   16c2"
    '    ConvertBasicToNokiB = Replace(Trim$(UCase$(ConvertBasicToNokiB)), "  ", " ")
    '    If ConvertBasicToNokiB <> BasicaTranslate Then
    '        Stop
    '    End If

End Function

Public Property Get BorderStyle() As Variant

    BorderStyle = M_KB.BorderStyle

End Property

Public Property Let BorderStyle(ByVal vNewValue As Variant)

  'wrapper Property for the PictureBox

    KeepInBoundsSimple 0, vNewValue, 1
    M_KB.BorderStyle = vNewValue

End Property

Private Function BRGetPlatform() As OsType

  'for speaker beep function, only the platform type is relevant
  'Internal Speaker Beep Routine
  'by Bob Richards
  'adapted from Jorge Loubet's code for Win9x
  'MODIFICATION  Roger Gilchrist None here
  'but Bob's code also contains a GetVersion function
  'he used for a MsgBox which I didn't want

  Dim osinfo As OSVERSIONINFO
  Dim retvalue As Integer

    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)

    Select Case osinfo.dwPlatformId
      Case 1
        BRGetPlatform = Win9xMe
      Case 2
        BRGetPlatform = Nt2000
      Case Else
        BRGetPlatform = OsUnknown
    End Select

End Function

Private Sub BRPcSpeakerBeep(ByVal FreqHz As Single, ByVal LengthMs As Single)

  'This is where the beep method depends on the operating system
  'Internal Speaker Beep Routine
  'by Bob Richards
  'adapted from Jorge Loubet's code for Win9x
  'added code for NT2000
  '
  'This program tests for the operating system and executes
  'a beep routine appropriate for the system.
  '
  'For Win9x systems, the file WIN95IO.DLL must be copied
  'to the Windows/System folder.
  'WIN95IO.DLL is available from http://www.softcircuits.com

  'MODIFICATION  Roger Gilchrist
  'Bob Richards renamed the API Beep to NtBeep but I figured that might worry new coders (XP is part of the NT family)
  'So instead I used the explicit VBA.Beep to allow this routine to cope with getting the OsUnknown Beep not the API Beep.
  'Also added a MsgBox although I think Bob only included the Case as an extreme future safety feature,
  'I don't think your program would even run long enough to reach this Sub if the possibility occurred.
  'Also changed programming style see commented version which are Bob's. (just a matter of style, I think)
  'Also added the m_LongestNoteValue test (about 30 seconds which should be more than anyone can stand of PC Speaker tones)
  'and MsgBox to stop Beep locking computer while playing enormously long beeps
  'To test this try a Basica note such as  "T120 L1 O4 C........."

    If LengthMs < m_LongestNoteValue Then
        Select Case BRGetPlatform
          Case Win9xMe
            Call BRWin9xBeep(CInt(FreqHz), LengthMs)
          Case Nt2000
            'Call NtBeep(CLng(FreqHz), CLng(LengthMs))

            Beep FreqHz, LengthMs
          Case OsUnknown
            VBA.Beep    'use the default beep routine, probably the sound card
            MsgBox "BRPcSpeakerBeep routine does not know what operating system you are using!", vbCritical, "BRPcSpeakerBeep Error"
        End Select
      Else 'NOT LENGTHMS...
        MsgBox "Beep Duration = " & LengthMs & " Current Limit = " & m_LongestNoteValue & "." & vbNewLine & _
               "This Duration value would play a note for about " & GenerateTime$(CInt(LengthMs / 2000)) & "." & vbNewLine & _
               "If you want that change the Limit by setting LongestNote." & vbNewLine & _
               "Warning very large values lock the computer in both IDE and compiled programs." & vbNewLine & _
               "In WInXP you can use Ctrl-Alt-Del to launch Task Manager and close it but in Pre NT systems you will be stuck.", vbCritical, "Very, Very Long Duration"
    End If

End Sub

'following routine largely by Jorge Loubet
Private Sub BRWin9xBeep(ByVal Freq As Integer, ByVal Length As Single)

  'Internal Speaker Beep Routine
  'by Bob Richards
  'adapted from Jorge Loubet's code for Win9x
  'added code for NT2000
  '
  'This program tests for the operating system and executes
  'a beep routine appropriate for the system.
  '
  'For Win9x systems, the file WIN95IO.DLL must be copied
  'to the Windows/System folder.
  'WIN95IO.DLL is available from http://www.softcircuits.com

  Dim LoByte As Integer
  Dim HiByte As Integer
  Dim Clicks As Integer
  Dim SpkrOn As Integer
  Dim SpkrOff As Integer
  Dim TimeEnd As Single
  Static ErrorShown As Boolean 'Added this so you don't get locked into a tune's worth of msgboxes Roger Gilchrist

    'MODIFICATION  Roger Gilchrist next two lines

    On Error GoTo Oops

    TimeEnd = Timer + Length / 1000

    'Ports 66, 67, and 97 control timer and speaker
    '
    'Divide clock frequency by sound frequency
    'to get number of "clicks" clock must produce.
    Clicks = CInt(1193280 / Freq)
    LoByte = Clicks And &HFF
    HiByte = Clicks \ 256
    'Tell timer that data is coming
    vbOut 67, 182
    'Send count to timer
    vbOut 66, LoByte
    vbOut 66, HiByte
    'Turn speaker on by setting bits 0 and 1 of PPI chip.
    SpkrOn = vbInp(97) Or &H3
    vbOut 97, SpkrOn

    'Leave speaker on (while timer runs)
    Do While Timer < TimeEnd
        'Let processor do other tasks
        DoEvents
    Loop
    'Turn speaker off.
    SpkrOff = vbInp(97) And &HFC
    vbOut 97, SpkrOff
    'MODIFICATION  Roger Gilchrist rest of Sub

Exit Sub

Oops:
    If Not ErrorShown Then
        MsgBox "This should work but for Win9x systems, the file WIN95IO.DLL must be copied to the Windows/System folder." & vbNewLine & _
               "Bob Richards says WIN95IO.DLL is available from http://www.softcircuits.com" & vbNewLine & _
               "If you have it and have installed it this MsgBox should never appear." & vbNewLine & _
               "Let me know if that does not work, I don't have access to a system to test it on.", _
               vbInformation, _
               "You are trying to use my class on a Win 9x system"
        ErrorShown = True
    End If

End Sub

Private Sub Class_Initialize()

  'Copyright 2002 Roger Gilchrist
  'Note there are three different implementations of this Collection in this demo
  'each class has a different one tailored to its needs. You can not swap and change them.

  Dim TmpArrayValue As Variant, TmpArrayName As Variant, i As Integer

    m_RTTTLtitle = "Untitled"
    Randomize Timer
    'My default values for MouseDuration
    m_DurationSettings(ButtonL) = 2
    m_DurationSettings(ButtonR) = 4
    m_DurationSettings(ButtonM) = 8
    m_DurationSettings(ShiftButtonL) = 16
    m_DurationSettings(ShiftButtonR) = 32
    m_DurationSettings(ShiftButtonM) = 64
    m_DurationSettings(CtrlButtonL) = 1
    m_DurationSettings(CtrlButtonR) = 1
    m_DurationSettings(CtrlButtonM) = 1

    ButtonDur = "Duration: L=" & Fractional(4 / GetButtonValue(ButtonL)) & " R=" & Fractional(4 / GetButtonValue(ButtonR)) & " M=" & Fractional(4 / GetButtonValue(ButtonM)) & _
                " Shift+L=" & Fractional(4 / GetButtonValue(ShiftButtonL)) & " Shift+R=" & Fractional(4 / GetButtonValue(ShiftButtonR)) & " Shift+M=" & Fractional(4 / GetButtonValue(ShiftButtonM)) & _
                " Ctrl+Any=" & Fractional(4 / GetButtonValue(CtrlButtonL))

    m_LongestNoteValue = 30000
    m_ShowNokiaOctave = False
    ColorHiNokiaOct = RGB(250, 225, 255)
    ColorMidNokiaOct = RGB(220, 215, 255)
    ColorLoNokiaOct = RGB(180, 215, 255)
    m_NokiaDisableColour = RGB(50, 50, 50)
    'Default values
    m_Tempo = 120
    KWLet = Array("C", "D", "E", "F", "G", "A", "B", "P") 'Arrays used to name keys
    KBLet = Array("#C", "#D", "#F", "#F", "#G", "#A", "")
    m_UseColorTags = True          'Show C keys with colour tagging
    KeyBoardLook Default
    m_WholeKeyTag = False          'C key colour tags can be whole key or just lower tip
    m_OffSetEnds = True            ' ends of key board can be against left & right edges of PBox
    'or slightly inset (Looks better; helps MouseMove message to stay upto date)

    m_PauseKeyOn = False

    m_UseTxtOutPut = Basica
    m_ShowOnFormCaption = False

    'this collection is based on a colFR collection copyright of Barry Dune
    'The collection (and my array version of it) represent
    'the frequencies of piano keys and their notation in Nokia code
    'except for the first value "0", 32767 which is only there to support the Basica note code 'N0'
    'You could construct a similar collection for other musical notations
    'but remember Basica and (I guess) Nokia code are piano based.
    'I have laid the TmpArrayXXX out one octave to a line
    'to make them a bit easier to play with.
    '(see 'Note on NoteCollection.txt' for more details.

    'Initialize the returned key values

    TmpArrayValue = Array(0, _
                    32.703, 34.648, 36.708, 38.891, 41.203, 43.654, 46.249, 48.999, 51.913, 55, 58.27, 61.735, _
                    65.406, 69.296, 73.416, 77.782, 82.407, 87.307, 92.499, 97.999, 103.826, 110, 116.541, 123.471, _
                    130.813, 138.591, 146.832, 155.564, 164.814, 174.614, 184.997, 195.998, 207.652, 220, 233.082, 246.942, _
                    261.626, 277.183, 293.665, 311.127, 329.628, 349.228, 369.994, 391.995, 415.305, 440, 466.164, 493.883, _
                    523.251, 554.365, 587.33, 622.254, 659.255, 698.457, 739.989, 783.991, 830.609, 880, 932.328, 987.767, _
                    1046.502, 1108.731, 1174.659, 1244.508, 1318.51, 1396.913, 1479.978, 1567.982, 1661.219, 1760, 1864.655, 1975.533, _
                    2093.005, 2217.461, 2349.318, 2489.016, 2637.021, 2793.826, 2959.956, 3135.964, 3322.438, 3520, 3729.31, 3951.066, _
                    4186.009, 4434.922, 4698.637, 4978.032, 5274.042, 5587.652, 5919.912, 6271.928, 6644.876, 7040, 7458.62, 7902.133, _
                    8372.019, 8869.845, 9397.273, 9956.064, 10548.083, 11175.305, 11839.823, 12543.855, 13289.752)

    TmpArrayName = Array("0", _
                   "C2", "#C2", "D2", "#D2", "E2", "F2", "#F2", "G2", "#G2", "A2", "#A2", "B2", _
                   "C3", "#C3", "D3", "#D3", "E3", "F3", "#F3", "G3", "#G3", "A3", "#A3", "B3", _
                   "C4", "#C4", "D4", "#D4", "E4", "F4", "#F4", "G4", "#G4", "A4", "#A4", "B4", _
                   "C5", "#C5", "D5", "#D5", "E5", "F5", "#F5", "G5", "#G5", "A5", "#A5", "B5", _
                   "C6", "#C6", "D6", "#D6", "E6", "F6", "#F6", "G6", "#G6", "A6", "#A6", "B6", _
                   "C7", "#C7", "D7", "#D7", "E7", "F7", "#F7", "G7", "#G7", "A7", "#A7", "B7", _
                   "C8", "#C8", "D8", "#D8", "E8", "F8", "#F8", "G8", "#G8", "A8", "#A8", "B8", _
                   "C9", "#C9", "D9", "#D9", "E9", "F9", "#F9", "G9", "#G9", "A9", "#A9", "B9", _
                   "C10", "#C10", "D10", "#D10", "E10", "F10", "#F10", "G10", "#G10")

    ReDim NoteValueA(UBound(TmpArrayValue)) As NoteValue
    For i = LBound(TmpArrayValue) To UBound(TmpArrayValue)
        NoteValueA(i).Name = TmpArrayName(i)
        NoteValueA(i).Value = TmpArrayValue(i)
    Next i

End Sub

Public Function CleanTextOutPut(Mode As TextTypeOutPut) As String

    Select Case Mode
      Case Basica
        CleanTextOutPut = Trim$(m_BasicaClean)
      Case Nokia
        CleanTextOutPut = Trim$(m_NokiaClean)
      Case RTTTL

        CleanTextOutPut = Trim$(m_RTTTLClean)
    End Select

End Function

Public Property Let ColorBlackKey(ByVal Colr As Long)

    m_BlackKeyColour = Colr
    KeyBoardDraw

End Property

Public Property Get ColorBlackKey() As Long

    ColorBlackKey = m_BlackKeyColour

End Property

Public Property Let ColorCKey(ByVal Colr As Long)

    m_C_Colour = Colr
    KeyBoardDraw

End Property

Public Property Get ColorCKey() As Long

    ColorCKey = m_C_Colour

End Property

Public Property Let ColorHiNokiaOct(ByVal HiNokiaOct As Long)

    m_ColorHiNokiaOct = HiNokiaOct

End Property

Public Property Get ColorHiNokiaOct() As Long

    ColorHiNokiaOct = m_ColorHiNokiaOct

End Property

Public Property Let ColorKeyEdge(ByVal Colr As Long)

  'Copyright 2002 Roger Gilchrist
  'Set edge Colour; Any Appearance change should call KeyBoardDraw

    m_EdgeColour = Colr
    KeyBoardDraw

End Property

Public Property Get ColorKeyEdge() As Long

    ColorKeyEdge = m_EdgeColour

End Property

Public Property Let ColorLoNokiaOct(ByVal LoNokiaOct As Long)

    m_ColorLoNokiaOct = LoNokiaOct

End Property

Public Property Get ColorLoNokiaOct() As Long

    ColorLoNokiaOct = m_ColorLoNokiaOct

End Property

Public Property Get ColorMidCKey() As Long

    ColorMidCKey = m_MidC_Colour

End Property

Public Property Let ColorMidCKey(ByVal Colr As Long)

  'Copyright 2002 Roger Gilchrist
  'Set MidC Colour; Any Appearance change should call KeyBoardDraw

    m_MidC_Colour = Colr
    KeyBoardDraw

End Property

Public Property Get ColorMidNokiaOct() As Long

    ColorMidNokiaOct = m_ColorMidNokiaOct

End Property

Public Property Let ColorMidNokiaOct(ByVal MidNokiaOct As Long)

    m_ColorMidNokiaOct = MidNokiaOct

End Property

Public Property Let ColorPause(ByVal Colr As Long)

    m_PauseKeyColour = Colr
    KeyBoardDraw

End Property

Public Property Get ColorPause() As Long

    ColorPause = m_PauseKeyColour

End Property

Public Property Get ColorPressed() As Long

    ColorPressed = m_PressedColour

End Property

Public Property Let ColorPressed(ByVal Colr As Long)

  'Copyright 2002 Roger Gilchrist
  'Set edge Colour; Any Appearance change should call KeyBoardDraw

    m_PressedColour = Colr
    KeyBoardDraw

End Property

Public Property Get ColorWhiteKey() As Long

    ColorWhiteKey = m_WhiteKeyColour

End Property

Public Property Let ColorWhiteKey(ByVal Colr As Long)

  'Copyright 2002 Roger Gilchrist
  'Set White Key Colour; Any Appearance change should call KeyBoardDraw

    m_WhiteKeyColour = Colr
    KeyBoardDraw

End Property

Public Function ColoursSafe() As Boolean

  'This routine was originally designed to return false if you used the same colour twice
  'But I changed it so that it just varies the colour in a repeatable fashion
  'Colous at start of the array take precedence over later colours so
  'Black keys are always the colour you assign
  'White keys only change if you have set them to equal Black keys and so on.
  '
  ' if you prefer the Boolean T|F version comment out the 'ReTest' parts and the value changing line
  'and uncomment the 'ColoursSafe = False' line there is a message box which will tell you something went wrong
  ' and stop the keyboard drawing but I think this way is safer (Even if you build an error into the KeyBoardLook routine
  ' it will be consistently fixed here and you and your end user's will never know.

  Dim TestColoursArray As Variant, i As Integer, j As Integer

    TestColoursArray = Array(m_BlackKeyColour, m_WhiteKeyColour, m_PressedColour, m_C_Colour, m_MidC_Colour, m_EdgeColour, _
                       m_PauseKeyColour, m_BackColour, ColorHiNokiaOct, ColorMidNokiaOct, ColorLoNokiaOct, m_NokiaDisableColour)
ReTest:
    For i = LBound(TestColoursArray) To UBound(TestColoursArray)
        For j = LBound(TestColoursArray) To UBound(TestColoursArray)
            If i <> j Then
                If TestColoursArray(i) = TestColoursArray(j) Then
                    Select Case TestColoursArray(j)
                      Case vbWhite, vbBlack
                        TestColoursArray(j) = RGB(127, 127, 127)
                      Case Else
                        TestColoursArray(j) = TestColoursArray(j) - IIf(TestColoursArray(j) > 50, -50, 50)
                    End Select
                    GoTo ReTest
                End If
            End If
        Next j
    Next i

    m_BlackKeyColour = TestColoursArray(0)
    m_WhiteKeyColour = TestColoursArray(1)
    m_PressedColour = TestColoursArray(2)
    m_C_Colour = TestColoursArray(3)
    m_MidC_Colour = TestColoursArray(4)
    m_EdgeColour = TestColoursArray(5)
    m_PauseKeyColour = TestColoursArray(6)
    m_BackColour = TestColoursArray(7)
    ColorHiNokiaOct = TestColoursArray(8)
    ColorMidNokiaOct = TestColoursArray(9)
    ColorLoNokiaOct = TestColoursArray(10)
    m_NokiaDisableColour = TestColoursArray(11)
    ColoursSafe = True

End Function

Private Function ConvertTo7s(Nval As Integer) As Integer

    Select Case Nval
      Case 0, 1
        ConvertTo7s = 0 'C 'C#D-
      Case 2, 3
        ConvertTo7s = 1 'D 'D#,E-
      Case 4
        ConvertTo7s = 2 'E
      Case 5, 6
        ConvertTo7s = 3 'F 'F#G-
      Case 7, 8
        ConvertTo7s = 4 'G'G#A-
      Case 9, 10
        ConvertTo7s = 5 'A 'A#B-
      Case 11
        ConvertTo7s = 6 'B
    End Select

End Function

Private Function CountChar(a, c$) As Long

  Dim i As Integer

    For i = 1 To Len(a)
        If Mid$(a, i, 1) = c$ Then
            CountChar = CountChar + 1
        End If
    Next i

End Function

Private Function CounterMember(a, Mode As CMEnum) As Long

  Dim i As Long, tVal As Long, MemberNo As Long

    Select Case Mode
      Case SmallestMember, SmallestValue ' get high limit for use by low search
        tVal = CounterMember(a, LargestValue)
    End Select
    For i = LBound(a) To UBound(a)
        Select Case Mode
          Case LargestMember, LargestValue
            If tVal < a(i) Then
                tVal = a(i)
                MemberNo = i
            End If

          Case SmallestMember, SmallestValue
            If tVal > a(i) Then
                tVal = a(i)
                MemberNo = i
            End If
          Case SmallestNotZeroMember, SmallestNotZeroValue
            If a(i) > 0 Then
                If tVal > a(i) Then
                    tVal = a(i)
                    MemberNo = i
                End If
            End If
        End Select
    Next i

    Select Case Mode
      Case LargestMember, SmallestMember, SmallestNotZeroMember
        CounterMember = MemberNo
      Case LargestValue, SmallestValue, SmallestNotZeroValue
        CounterMember = tVal
    End Select

End Function

Public Property Get DurationVal() As Integer

  'Return Mouse button based Duration to program

    DurationVal = m_Duration

End Property

Public Function Fractional(Dec As Variant, Optional Limiter As LimitType = DoubleLimit) As String

  'Copyright 2002 Roger Gilchrist
  'Probably been done before but it's cute and useful
  'given a decimal fraction 0.125 will return a fractional string 1/4
  'a whole number (5.0) returns (5)

  'Because of internal math the only real differencce in LimitType is
  'between Integer and any other type
  '(For Integer the Irration trap is sprung for all others it is not)

  'WARNING 1
  'Very large nominator values in mixed numbers can overwhelm the function
  'WARNING 2
  'irrational numbers like Pi are returned marked "*APPROX* "Value" *APPROX*" if using Integer LimitType
  'WARNING 3
  'repeated fractions like .333 return at first wholenumber denominator
  ' unless you send them at maximum VB length(just enter .3333333333333333333333333333333333 and let VB truncate it it to 0.333333333333333)
  '

  Dim Numerator As Variant, Denominator As Variant, IrrationalHitCounter As Double
  Dim WholeValue As Variant
  Dim TmpVal As Variant, Tmpval2 As Variant
  Dim limit As Double
  Const IntLimit As Integer = 32767 - 1
  Const LngLimit As Long = 2147483647 - 1
  Const DblLimit As Double = 1.79769313486232E+307 - 1
  Const SngLimit As Single = 3.402823E+38 - 1

    'these are all 1 less than maximum possible value

    Select Case Limiter
      Case IntegerLimit
        limit = IntLimit
      Case LongLimit
        limit = LngLimit
      Case SingleLimit
        limit = SngLimit
      Case DoubleLimit
        limit = DblLimit
    End Select
    limit = limit * 0.8 ' set limit low enough to trap irrational numbers
    If Dec = "1.#INF" Then 'Handle an error trapped calculation in the argument
        MsgBox "Division by Zero Error", vbCritical, "Fractional"
        Fractional = "Division by Zero Error"
        Exit Function '>---> Bottom
    End If

    If Dec < 1 Then         'Fraction only
        If Dec = 0 Then         'not caused by calculation in the argument but simple sending zero
            Fractional = 0
            Exit Function '>---> Bottom
        End If

        Dec = CSng(Dec)
        Numerator = Dec
      Else 'Mixed number'NOT DEC...
        TmpVal = (Dec - CSng(Dec))
        TmpVal = Dec - TmpVal
        If TmpVal > LngLimit Then      ''used to calculate fractional
            MsgBox "Too large to calculate.", vbCritical, "Fractional"
            'part of mixed numbers
            Fractional = Dec
            Exit Function '>---> Bottom
        End If
        Tmpval2 = CLng(TmpVal)
        If Tmpval2 - TmpVal <= 0.5 Then ' decide how to round down
            WholeValue = Fix(TmpVal)
          Else 'NOT TMPVAL2...
            WholeValue = Int(TmpVal)
        End If
        Dec = CSng(Dec - WholeValue) 'get fractional part
        If Dec = 0 Then              'if no fractional part
            Fractional = WholeValue
            Exit Function '>---> Bottom
        End If
        Numerator = Dec
    End If
    Denominator = 1

    Do Until Numerator = CLng(Numerator) 'Loop to get whole number fraction
        Numerator = Numerator + Dec
        Denominator = Denominator + 1
        IrrationalHitCounter = IrrationalHitCounter + 1
        If IrrationalHitCounter >= limit Then  'Exit if too many hits made
            Exit Do '>---> Loop
        End If

    Loop

    If Denominator = Numerator Then 'cope with 23/23 style input
        Fractional = "1"
      Else 'NOT DENOMINATOR...
        Fractional = Numerator & "/" & Denominator
    End If

    If WholeValue > 0 And Numerator > 0 Then 'Add wholeValue if it exists
        Fractional = WholeValue & " " & Fractional
    End If
    If IrrationalHitCounter >= limit Then 'Add ApProx message if needed
        Fractional = "*APPROX* " & Fractional & " *APPROX*"
    End If

End Function

Public Function Frequency() As Single

  'Copyright 2002 Roger Gilchrist
  'Mostly for debug purposes

    Frequency = FrequencyFromNoteName(OctaveMember & (m_Octave + 2))

End Function

Private Function FrequencyFromNoteName(ByVal n$) As Single

  Dim i As Long

    For i = 1 To UBound(NoteValueA)
        If NoteValueA(i).Name = n$ Then
            FrequencyFromNoteName = NoteValueA(i).Value
            Exit Function '>---> Bottom
        End If
    Next i

End Function

Private Function GenerateTime(tim) As String

  'Copyright 1998 Roger Gilchrist
  'Create a : delimited time string from tim

  Dim hr As Integer, mn As Integer
  Dim T As String

    hr = Int((tim / 60) / 60)
    If hr > 0 Then
        T$ = Format$(hr, "00") & " hour" & IIf(hr > 1, "s ", " ")
    End If
    tim = tim - hr * Int(86400 / 24)
    mn = Int(tim / 60)
    If mn > 0 Then
        T$ = T$ & Format$(mn, "00") & " minute" & IIf(hr > 1, "s ", " ")
    End If
    tim = tim - mn * Int(86400 / 24 / 60)
    If tim > 0 Then
        T$ = T$ & Format$(Int(tim), "00") & " second" & IIf(tim > 1, "s ", " ")
    End If
    GenerateTime$ = T$

End Function

Public Function GetButtonValue(ButDesc As ButtonShift) As Integer

    GetButtonValue = m_DurationSettings(ButDesc)

End Function

Private Function GetDurationValue(T$) As Single

  Dim DotValue As Long, Gama As Single

    GetDurationValue = GetRemoveNumerals(T$, True, True)
    If GetDurationValue = 0 Then
        GetDurationValue = 0.01
    End If
    DotValue = RemoveDots(T$)
    If DotValue Then
        GetDurationValue = GetDurationValue / (3 / 2) ^ DotValue
    End If
    Gama = 240000 / GetDurationValue
    GetDurationValue = Gama / m_Tempo

End Function

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

Private Function InstrArray(Test, ParamArray ltr()) As Boolean

  Dim i As Integer

    For i = LBound(ltr) To UBound(ltr)
        If InStr(Test, ltr(i)) Then
            InstrArray = True
            Exit Function '>---> Bottom
        End If
    Next i

End Function

Private Sub KeepInBounds(Min As Variant, Val As Variant, MAx As Variant)

  'Copyright 2002 Roger Gilchrist
  'Simple Version

    If Val > MAx Then
        Val = MAx
    End If
    If Val < Min Then
        Val = Min
    End If

End Sub

Private Sub KeepInBoundsSimple(Min As Variant, Val As Variant, MAx As Variant)

  'Copyright 2002 Roger Gilchrist
  'basic version of this sort of thing see ClsMusic_Nokia for complex version

    If Val < Min Then
        Val = Min
      ElseIf Val > MAx Then 'NOT VAL...
        Val = MAx
    End If

End Sub

Private Sub Key(Keyno As Long)

  'Copyright 2002 Roger Gilchrist
  'Draw one key and surrounding black keys if necessary
  'Show Pressed colour if argument is true
  'NOTE because I developed this for a one note at a time player there is no allowance for chords
  'If you need it you could operate this by storing arrays of pressed states for each key LeftB, White and RightB
  'and use the arrays to trigger the IIf colour statements

  Dim OctaveMember As Integer 'Variables required in more than one place in Sub so set once call often
  Dim OctaveNo As Integer
  Dim KeyX1 As Integer
  Dim BOffSet As Integer
  Dim TopOffset As Integer  ' set top of keyboard just far enough from PictureBox top to allow
  Dim KeyColour As Long

    If Not ReSizing Then
        ' cause the Off-Keyboard MouseMove message to appear
        If Keyno > KeyCount Then
            If (m_PauseKeyOn = False) Or (Keyno <> PauseKeyNo) Then 'Spaces around Pause key or Pause key not shown
                m_IsBlack = False
                M_IsWhite = False
                M_IsPause = False
                Exit Sub '>---> Bottom
            End If
        End If
        If KeyIsDown Then
            KeyStack = Keyno
        End If
    End If

    TopOffset = 9
    OctaveMember = Keyno Mod 7 'Detect any C and decide whether to draw Black Keys
    OctaveNo = Int(Keyno / 7) 'detect Middle C
    KeyX1 = (KeyWidth * Keyno) - (KeyWidth) + IIf(m_OffSetEnds, 40, 0) 'initial start for drawing key
    BOffSet = (KeyWidth / 2)  'offset value for the Black keys
    If ReSizing Then
        KeyColour = IIf(Keyno = PauseKeyNo, m_PauseKeyColour, m_WhiteKeyColour)
      Else 'RESIZING = FALSE
        KeyColour = IIf(m_WhitePressed, m_PressedColour, IIf(Keyno = PauseKeyNo, m_PauseKeyColour, m_WhiteKeyColour))
    End If
    If ShowNokiaSafeOctaves Then
        If Keyno < 21 Or Keyno > 41 And Keyno <> PauseKeyNo Then
            KeyColour = m_NokiaDisableColour
        End If
    End If
    M_KB.Line (KeyX1, TopOffset)-Step(KeyWidth, KeyHeight + TopOffset), KeyColour, BF
    If m_UseColorTags Then ' Colour tag system
        If Not (ShowNokiaSafeOctaves And KeyColour = m_NokiaDisableColour) Then
            If Not m_WhitePressed Then 'don't tag if drawing a pressed key
                'You could expand this to give a separate color to each key(OctaveMember is a 0 to 6 range
                If OctaveMember = 0 Then 'If its a C then --------------------------------------------------- and if it is Middle C then
                    M_KB.Line (KeyX1, KeyHeight - KeyTabtop + TopOffset)-Step(KeyWidth, KeyTabHeight), IIf(OctaveNo = 4, m_MidC_Colour, m_C_Colour), BF
                End If
            End If
        End If
    End If
    M_KB.Line (KeyX1, TopOffset)-Step(KeyWidth, KeyHeight + TopOffset), m_PressedColour, B 'Outline key
    'This draw routine hits the same black keys many times because they need to be redrawn every time they or the underlying white key is pressed
    If Keyno < 62 Then ' if it is a playable key
        KeyX1 = KeyX1 + BOffSet  'Right Black key
        Select Case OctaveMember
          Case 2, 6 'These don't have Sharps
          Case Else
            If Keyno > 0 And Keyno < KeyCount Then 'Not first or last key D0 and G10 which would normally have a Flat and a Sharp key
                'You could expand this to give a separate color to each key(OctaveMember is a 0 to 6 range
                If ReSizing Then
                    KeyColour = m_BlackKeyColour
                  Else 'RESIZING = FALSE
                    KeyColour = IIf(m_RightBlackPressed, m_PressedColour, m_BlackKeyColour)
                End If
                If ShowNokiaSafeOctaves Then
                    If Keyno < 21 Or Keyno > 41 And Keyno <> PauseKeyNo Then
                        KeyColour = m_NokiaDisableColour
                    End If
                End If

                M_KB.Line (KeyX1 + KeyWidth / 10, TopOffset)-Step(BKeyWidth, BKeyHeight), KeyColour, BF
                M_KB.Line (KeyX1 + KeyWidth / 10, TopOffset)-Step(BKeyWidth, BKeyHeight), m_PressedColour, B
            End If
        End Select
        KeyX1 = KeyX1 - (BOffSet * 2) 'Left Black key 'added one BoffSet so remove two
        Select Case OctaveMember
          Case 0, 3 'These don't have flats
          Case Else
            If Keyno > 1 Then 'And ((KeyNo < KeyCount - 1) And m_PauseKeyOn) Then
                'You could expand this to give a separate color to each key(OctaveMember is a 0 to 6 range
                If m_PauseKeyOn And Keyno = 62 Then

                  Else 'NOT M_PAUSEKEYON...
                    If ReSizing Then
                        KeyColour = m_BlackKeyColour
                      Else 'RESIZING = FALSE
                        KeyColour = IIf(m_LeftBlackPressed, m_PressedColour, m_BlackKeyColour)
                    End If
                    If ShowNokiaSafeOctaves Then
                        If Keyno < 21 Or Keyno > 41 And Keyno <> PauseKeyNo Then
                            KeyColour = m_NokiaDisableColour
                        End If
                    End If

                    M_KB.Line (KeyX1 + KeyWidth / 10, TopOffset)-Step(BKeyWidth, BKeyHeight), KeyColour, BF
                    M_KB.Line (KeyX1 + KeyWidth / 10, TopOffset)-Step(BKeyWidth, BKeyHeight), m_PressedColour, B
                End If
            End If
        End Select
        If m_ContainerIsFrame Then
            M_KBFrame.ForeColor = m_WhiteKeyColour 'does nothing in XPAware Programs
        End If

    End If
    If ShowNokiaSafeOctaves Then ' puts the colour bar indicator across the top of keyboards
        M_KB.Line (KeyWidth * 20 + IIf(m_OffSetEnds, 40, 0), TopOffset)-Step(KeyWidth * 7, KeyHeight * 0.17), m_ColorLoNokiaOct, BF
        M_KB.Line (KeyWidth * 27 + IIf(m_OffSetEnds, 40, 0), TopOffset)-Step(KeyWidth * 7, KeyHeight * 0.17), m_ColorMidNokiaOct, BF
        M_KB.Line (KeyWidth * 34 + IIf(m_OffSetEnds, 40, 0), TopOffset)-Step(KeyWidth * 7, KeyHeight * 0.17), m_ColorHiNokiaOct, BF
    End If
    M_KB.Refresh

End Sub

Private Sub KeyBoardDraw()

  'Copyright 2002 Roger Gilchrist
  'Draw whole keyboard

  Dim i As Long

    'test that you have initializsed the class first
    If M_KB Is Nothing Then
        Exit Sub '>---> Bottom
    End If
    If Not ColoursSafe Then
        MsgBox "The Colours you have set are not acceptable. Please reset them"
        'this will never hit in current programming
      Else 'NOT NOT...
        'Set basic measurements; Resets if size has been changed
        M_KB.BackColor = m_BackColour
        'White Keys
        KeyWidth = M_KB.Width / IIf(m_PauseKeyOn, KeyCount + 2, KeyCount) + IIf(m_OffSetEnds, -3, 0)
        KeyHeight = M_KB.Height * 0.9
        'BlackKeys
        BKeyWidth = KeyWidth * 0.7
        BKeyHeight = KeyHeight * 0.7

        'C Key Tags; Change these through the WholeKeyTag or KeyTag property

        If m_WholeKeyTag Then
            KeyTabtop = KeyHeight
            KeyTabHeight = KeyHeight
          Else 'M_WHOLEKEYTAG = FALSE
            KeyTabtop = KeyHeight / Def_KeyTabTopDiv
            KeyTabHeight = KeyHeight / Def_KeyTabHeightDiv
        End If
        'draw all standard keys
        For i = 1 To KeyCount
            If i <> 61 Then
                Key i
            End If
        Next i
        'draw pause key if needed
        If m_PauseKeyOn Then
            Key PauseKeyNo
        End If
        M_KB.Refresh
    End If

End Sub

Public Sub KeyBoardLook(Look As KBLook)

  'Copyright 2002 Roger Gilchrist
  'Be sure that the Combo box that calls here is filled in the same order as the KBLook Enum members.
  'Antique, Classic, Default and Random  styles are original to the code anything else someone else did
  'colours should not be the same but the Routine ColoursSafe makes sure that any duplicate or unassigned
  'colours are set to an unused value if you don't set it explicitly (that's why Random style is safe to use)
  'updated: On fast systems you won't notice (in Win95 486 it was quite strange and ugly)
  '           the old way of calling the Properties to reset colours also fired the KeyBoardDraw 10 times
  '           setting the Private variable directly then calling Resize redraws just once

    Select Case Look
      Case Default
        m_BlackKeyColour = vbBlack         '
        m_WhiteKeyColour = vbWhite         '
        m_PressedColour = 15132390         'Default Key Down Colour
        m_C_Colour = vbYellow            'Default C key
        m_MidC_Colour = vbRed            'Default Middle C
        m_EdgeColour = m_PressedColour 'Key Edges ; Key edges are to allow a little lee way for the class
        'to work out which black key is being clicked
        m_PauseKeyColour = RGB(160, 160, 255) 'Pause colour (set it even if you are not using it(no longer necessary as ColorsSafe will deal with it)

        m_BackColour = RGB(1, 1, 1)        'PictureBox colour where not covered by keys
        m_UseColorTags = True
        m_WholeKeyTag = False             'C key colour tags can be whole key or just lower tip
      Case Antique
        m_BlackKeyColour = RGB(0, 0, 0)
        m_WhiteKeyColour = RGB(255, 255, 230)
        m_PressedColour = RGB(155, 155, 130)
        m_C_Colour = RGB(255, 450, 50)
        m_MidC_Colour = RGB(255, 240, 100)
        m_EdgeColour = RGB(145, 145, 130)
        m_PauseKeyColour = RGB(251, 251, 230)
        m_BackColour = RGB(1, 1, 1)
        m_UseColorTags = True
        m_WholeKeyTag = False
      Case Classical
        m_BlackKeyColour = RGB(30, 30, 30)
        m_WhiteKeyColour = RGB(255, 255, 250)
        m_PressedColour = RGB(210, 210, 185)
        m_C_Colour = RGB(255, 255, 235)
        m_MidC_Colour = RGB(255, 255, 215)
        m_EdgeColour = RGB(210, 210, 180)
        m_PauseKeyColour = RGB(251, 251, 250)
        m_BackColour = RGB(0, 0, 0)
        m_UseColorTags = True
        m_WholeKeyTag = False
      Case Random
        Randomize Timer
        m_BlackKeyColour = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
        m_WhiteKeyColour = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
        m_PressedColour = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
        m_C_Colour = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
        m_MidC_Colour = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
        m_EdgeColour = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
        m_PauseKeyColour = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
        m_BackColour = RGB(Int(Rnd * 255), Int(Rnd * 255), Int(Rnd * 255))
        m_UseColorTags = Rnd > 0.5
        m_WholeKeyTag = Rnd > 0.5

      Case Else
        KeyBoardLook Default 'recursive safety net
    End Select
    Resize 'cause redraw to occur

End Sub

Public Sub KeyBoardLookPrint()

  'Copyright 2002 Roger Gilchrist
  'dump KeyBoardLook settings to disk
  'Allows you to create a keyboard look by saving a RAndom look and editing it into
  'Sub KeyBoardLook as a new Case

  Dim Fnum As Integer, str As String

    str = vbNewLine & "Just cut and paste into a new Case in Sub KeyBoardLook," & vbNewLine & "Create a new KBLook Enum member to name it." & vbNewLine & vbNewLine & _
          "Case DummyEnum_Replace_This_With_Your_New_KBLook_Enum" & vbNewLine & _
          "     m_BlackKeyColour = " & m_BlackKeyColour & vbNewLine & _
          "     m_WhiteKeyColour = " & m_WhiteKeyColour & vbNewLine & _
          "     m_PressedColour = " & m_PressedColour & vbNewLine & _
          "     m_C_Colour = " & m_C_Colour & vbNewLine & _
          "     m_MidC_Colour = " & m_MidC_Colour & vbNewLine & _
          "     m_EdgeColour = " & m_EdgeColour & vbNewLine & _
          "     m_PauseKeyColour = " & m_PauseKeyColour & vbNewLine & _
          "     m_BackColour = " & m_BackColour & vbNewLine & _
          "     m_UseColorTags = " & m_UseColorTags & vbNewLine & _
          "     m_WholeKeyTag = " & m_WholeKeyTag & vbNewLine

    Fnum = FreeFile
    Open App.Path & "\" & "KeyBoardLooks.txt" For Append As Fnum
    Print #1, str
    Close #Fnum

End Sub

Private Sub KeyColour(X As Single, Y As Single)

  Dim ColourTest As Long 'get colour at mouse pointer

    ColourTest = M_KB.Point(X, Y)
    m_IsBlack = (ColourTest = m_BlackKeyColour)
    If m_IsBlack Then 'complexity of these If..Else..End If speeds up reading of key state meaning that multiple clicks are not lost
        M_IsWhite = False
        M_IsPause = False
      Else 'M_ISBLACK = FALSE
        M_IsWhite = ((ColourTest = m_WhiteKeyColour) Or (ColourTest = m_C_Colour) Or (ColourTest = m_MidC_Colour))
    End If
    If m_IsBlack Or M_IsWhite Then
        M_IsPause = False
      Else 'NOT M_ISBLACK...
        M_IsPause = (ColourTest = m_PauseKeyColour) And m_PauseKeyOn
    End If

End Sub

Private Function KeyIsDown() As Boolean

    KeyIsDown = m_IsBlack Or M_IsWhite ' Or M_IsPause

End Function

Private Function KeyNoFromNote(ByVal n$) As Long

  'copyright 2002 Roger Gilchrist
  'TEST ME should cope with Basica and Nokia
  'Seems to work OK

  Dim Junk As Single, Oct As Long, NNam As String

    Junk = GetRemoveNumerals(n$, True) 'don't need this but have to remove the characters
    Oct = CLng(GetRemoveNumerals(n$, False))
    NNam = n$
    m_IsBlack = InStr(NNam, "#") Or InStr(NNam, "-")
    M_IsWhite = InStr(NNam, "#") = 0 And InStr(NNam, "-") = 0
    m_WhitePressed = M_IsWhite
    m_RightBlackPressed = m_IsBlack
    NNam = Replace(n$, "#", "")
    KeyNoFromNote = Oct * 7 + InStr("CDEFGAB", NNam) - 1
    If KeyNoFromNote <> m_KeyNumber Then
        'Stop
    End If

End Function

Public Property Get KeyStack() As Long

  'KeysDown acts as a stack; this pops a value off the stack
  ' and is used to call key to draw unpressed keys
  'called in MouseUp and PlayKBNote

    If KeysDown.Count Then
        KeyStack = CLng(KeysDown.Item(KeysDown.Count))
        KeysDown.Remove KeysDown.Count
    End If

End Property

Public Property Let KeyStack(ByVal Keyno As Long)

  'KeysDown acts as a stack; this pushes a value on to the stack
  'called in Key

    KeysDown.Add Keyno

End Property

Public Property Get LongestNote() As Single

    LongestNote = m_LongestNoteValue

End Property

Public Property Let LongestNote(ByVal LongNote As Single)

    m_LongestNoteValue = LongNote

End Property
Private Function MemberOfArray(tVal As Variant, pMembers) As Integer
Dim i As Long
    For i = LBound(pMembers) To UBound(pMembers)
If tVal = pMembers(i) Then
            MemberOfArray = i
            Exit Function '>---> Bottom
        End If
    Next i
MemberOfArray = -1
End Function

Private Function MemberOfParamArray(tVal As Variant, ParamArray pMembers()) As Boolean

  Dim wal As Variant

    For Each wal In pMembers
        If wal = tVal Then
            MemberOfParamArray = True
            Exit Function '>---> Bottom
        End If
    Next wal

End Function

Private Function MisKeyFix(WrongKey As String) As String

  'assumes you have miskeyed by one key

    Select Case WrongKey
      Case "Q", "Z"
        MisKeyFix = "A"
      Case "W"
        MisKeyFix = IIf(Rnd > 0.3, "A", "E")
      Case "S"
        MisKeyFix = IIf(Rnd > 0.5, "A", "D")
      Case "X"
        MisKeyFix = IIf(Rnd > 0.3, "C", "D")
      Case "R"
        MisKeyFix = IIf(Rnd > 0.3, "E", "F")
      Case "T"
        MisKeyFix = IIf(Rnd > 0.4, "G", "F")
      Case "V"
        MisKeyFix = IIf(Rnd > 0.5, "C", "B")
      Case "H"
        MisKeyFix = IIf(Rnd > 0.3, "G", "B")
      Case "N"
        MisKeyFix = IIf(Rnd > 0.3, "B", "G")
      Case "Y"
        MisKeyFix = "G"
      Case Else
        'if that has not worked assign a legal key at random
        MisKeyFix = RandomNoteName
    End Select

End Function

Public Sub MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'Copyright 2002 Roger Gilchrist

    If MouseRead(Button, Shift, X, Y) Then  'Make sure it is up-to-date then
        m_MouseX = X
        'KeyColour X, Y

        If Playable Then 'Moved from MouseRead as only used here
            m_WhitePressed = M_IsWhite And Button > 0
            m_OctaveMember = (m_KeyNumber) Mod 7 ' FIXED this was outside the 'If Playable Then' and caused move to reset
            '  On Error Resume Next
            If m_IsBlack Then
                If X <= (KeyWidth * m_KeyNumber) - (KeyWidth / 2) Then
                    m_OctaveMember = m_OctaveMember - 1 '
                    m_LeftBlackPressed = Button > 0
                  Else 'NOT X...
                    m_OctaveMember = m_OctaveMember
                    m_RightBlackPressed = Button > 0
                End If

            End If
            If M_IsPause Then
                m_Octave = PrevPressedOctave
              Else 'M_ISPAUSE = FALSE
                m_Octave = Int(m_KeyNumber / 7)
                If m_WhitePressed Or m_LeftBlackPressed Or m_RightBlackPressed Then
                    PrevPressedOctave = m_Octave
                End If
            End If
            Key m_KeyNumber
            PlayKBNote DurationVal & OctaveMember & IIf(m_DottedNote, ".", "") & (Octave)
            'DEBUG LINE
            '  Key KeyStack 'If you are using breaks in IDE to step thru then uncomment this or
            'you will only get one press per key other wise leave it to MouseUp
        End If
    End If

End Sub

Private Function MouseDuration(Button As Integer, Shift As Integer) As Integer

  Dim ShiftValue As ShiftDescriptor

    If Button = 0 Then
        MouseDuration = m_Duration 'don't reset if just moving
        Exit Function '>---> Bottom
    End If
    'SetUnassigned
    m_DottedNote = (Shift And vbAltMask) 'T|F based on vbAltMask
    If m_DottedNote Then 'if has vbAltMask then remove to get rest of Shift value
        Shift = Shift Xor vbAltMask
    End If
    ShiftValue = ShiftReader(Shift)
    Select Case Button 'Base values
      Case vbLeftButton
        Select Case ShiftValue
          Case KNoShift
            MouseDuration = m_DurationSettings(ButtonL) '2
          Case KAlt
            '            MouseDuration = m_DurationSettings(AltButtonL) 'Unassigned
          Case KCtrl
            MouseDuration = m_DurationSettings(CtrlButtonL) '1
          Case KShift
            MouseDuration = m_DurationSettings(ShiftButtonL) ' 16
            '          Case KAltCtrl
            '            MouseDuration = m_DurationSettings(AltCtrlButtonL) 'Unassigned
            '          Case KShiftCtrl
            '            MouseDuration = m_DurationSettings(ShiftCtrlButtonL) 'Unassigned
            '          Case KShiftAlt
            '            MouseDuration = m_DurationSettings(ShiftAltButtonL) 'Unassigned
            '          Case KShiftAltCtrl
            '            MouseDuration = m_DurationSettings(ShiftAltCtrlButtonL) 'Unassigned
        End Select

      Case vbRightButton
        Select Case ShiftValue
          Case KNoShift
            MouseDuration = m_DurationSettings(ButtonR) '4
            '          Case KAlt
            '            MouseDuration = m_DurationSettings(AltButtonR) 'Unassigned
          Case KCtrl
            MouseDuration = m_DurationSettings(CtrlButtonR) '1
          Case KShift
            MouseDuration = m_DurationSettings(ShiftButtonR) '32
            '          Case KAltCtrl
            '            MouseDuration = m_DurationSettings(AltCtrlButtonR) 'Unassigned
            '          Case KShiftCtrl
            '            MouseDuration = m_DurationSettings(ShiftCtrlButtonR) 'Unassigned
            '          Case KShiftAlt
            '            MouseDuration = m_DurationSettings(ShiftAltButtonR) 'Unassigned
            '          Case KShiftAltCtrl
            '            MouseDuration = m_DurationSettings(ShiftAltCtrlButtonR) 'Unassigned
        End Select

      Case vbMiddleButton
        Select Case ShiftValue
          Case KNoShift
            MouseDuration = m_DurationSettings(ButtonM) '8
            '          Case KAlt
            '            MouseDuration = m_DurationSettings(AltButtonM) 'Unassigned
          Case KCtrl
            MouseDuration = m_DurationSettings(CtrlButtonM) '1
          Case KShift
            MouseDuration = m_DurationSettings(ShiftButtonM) '64
            '          Case KAltCtrl
            '            MouseDuration = m_DurationSettings(AltCtrlButtonM) 'Unassigned
            '          Case KShiftCtrl
            '            MouseDuration = m_DurationSettings(ShiftCtrlButtonM) 'Unassigned
            '          Case KShiftAlt
            '            MouseDuration = m_DurationSettings(ShiftAltButtonM) 'Unassigned
            '          Case KShiftAltCtrl
            '            MouseDuration = m_DurationSettings(ShiftAltCtrlButtonM) 'Unassigned
        End Select
      Case Else
        Exit Function '>---> Bottom
    End Select

End Function

Public Sub MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'Copyright 2002 Roger Gilchrist
  'Read position then Display it on the frame(if it exists)or the Form if That is allowed
  'This works best if there is a small border of non-keyboard PictureBox around the KeyBoard so that the Second message appears
  'whenever the mouse is off the keys (also appears if cursor is over the keyEdges).

  Dim msg As String
  Dim MsgOct As Integer, MsgMemberNo As Integer, MsgMember As String

    MouseRead Button, Shift, X, Y
    If Playable Then  'And Button = 0 Then 'it is a note but just moving over it
        MsgMemberNo = (m_KeyNumber) Mod 7 ' FIXED this was outside the 'If Playable Then' and caused move to reset
        MsgOct = Int(m_KeyNumber / 7)
        If m_IsBlack Then
            If X <= (KeyWidth * m_KeyNumber) - (KeyWidth / 2) Then
                MsgMemberNo = MsgMemberNo - 1
            End If
        End If
        '
        '        If Not M_IsPause Then
        '            MsgOct = Int(m_KeyNumber / 7)
        '        End If

        If m_IsBlack Then
            MsgMember = KBLet(MsgMemberNo)
          ElseIf M_IsWhite Then 'M_ISBLACK = FALSE
            MsgMember = KWLet(MsgMemberNo)
          Else 'M_ISWHITE = FALSE
            MsgMember = "P"
        End If

        If M_IsPause Then
            msg = "Keyboard (Pause Key " & ButtonDur
          Else 'NOT OCTAVEMEMBER...'NOT MSGMEMBER...'M_ISPAUSE = FALSE
            msg = "Keyboard (Octave: " & MsgOct & "  Note:" & MsgMember & "  Freq: " & FrequencyFromNoteName(MsgMember & (MsgOct + 2)) & " )"
        End If
      Else 'PLAYABLE = FALSE
        msg = "Keyboard (Mouse Button Note " & ButtonDur
    End If

    ShowMsg msg

End Sub

Private Function MouseRead(Button As Integer, Shift As Integer, X As Single, Y As Single) As Boolean

  'Copyright 2002 Roger Gilchrist
  'gather information on keyboard position and mouse button states;
  'sets various variables and Properties for rest of class

    MouseRead = True
    If ShowNokiaSafeOctaves Then 'disable keys outside Nokia range
        MouseRead = Not (m_KeyNumber < 21 Or m_KeyNumber > 41 And m_KeyNumber <> PauseKeyNo)
    End If

    If m_PauseKeyOn Then
        MouseRead = Not (m_KeyNumber > KeyCount And m_KeyNumber <> PauseKeyNo)
    End If
    If MouseRead = False Then
        Exit Function '>---> Bottom
    End If
    KeyColour X, Y
    m_Duration = MouseDuration(Button, Shift)
    m_KeyNumber = Int((X + KeyWidth) / KeyWidth)

End Function

Public Sub MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

  'Copyright 2002 Roger Gilchrist
  'UnPress key
  'does not work if you are stepping through the code in IDE as IDE eats the event
  'not a problem as long as you don't set break points that hit between MouseDown and MouseUp

    If Playable Or M_KB.Point(X, Y) = m_PressedColour Then
        m_LeftBlackPressed = False
        m_RightBlackPressed = False
        m_WhitePressed = False
        If KeysDown.Count Then ' release key if it is pressed
            Key KeyStack 'pop and draw key
        End If
    End If

End Sub

Public Property Get MouseX() As Single

  'Copyright 2002 Roger Gilchrist
  'Mostly for Debug purposes

    MouseX = m_MouseX

End Property

Private Function Nokia2Basica(ByVal n$, Optional SharpInvert As Boolean = False) As String

  'Basica       "P<Duration>"
  'Basica Note Format "L<Duration> O<Octave> <Note>[#|-][...]"

  'Nokia format <Duration>P
  'Nokia Note Format <Duration>[...][#]<Note><Octave>

  Dim DurationValue As Single, DotValue As Long, Oct As Long
  Dim TmpN As String

    If Right$(n$, 1) = "P" Then
        n = Left$(n$, Len(n) - 1)
        DurationValue = GetRemoveNumerals(n, True)
        DotValue = RemoveDots(n)
        Nokia2Basica = "P" & DurationValue
      Else 'NOT LEFT$(N$,...'NOT RIGHT$(N$,...
        'L16O4C
        DurationValue = GetRemoveNumerals(n$, True)
        Oct = GetRemoveNumerals(n$, False) + 2    'get the Octave value and '+ 2' to match basica format
        DotValue = RemoveDots(n$)
        If SharpInvert Then
            n = BasicaInvertSharp(n)
        End If
        Nokia2Basica = "O" & CInt(Oct) & " L" & Int(DurationValue) & n & String$(DotValue, ".")
    End If

End Function

Private Function Nokia2RTTTL(n As String) As String

  Dim TmpVal As Integer, Tmpval2 As Integer, DotValue As Integer

    TmpVal = GetRemoveNumerals(n, False)
    If InStr(n, "#") Then
        Tmpval2 = GetRemoveNumerals(n, True)
        DotValue = RemoveDots(n$)
        n = BasicaInvertSharp(n)
        n = Tmpval2 & n & String$(DotValue, ".")
    End If

    If TmpVal > 0 Then
        Nokia2RTTTL = n & TmpVal + 3 'octave shift
      Else 'NOT TMPVAL...
        Nokia2RTTTL = n
    End If

End Function

Private Sub NokiaAutoFixCharacter()

  Dim Tmpa As Variant, TmpStr As String, TmpOct As Single, i As Long, j As Long

    Tmpa = Split(m_NokiaClean)
    For i = LBound(Tmpa) To UBound(Tmpa)
        TmpStr = Tmpa(i)
        For j = 1 To Len(TmpStr)
            If InStr("1234567890-#PABCDEFG", Mid$(TmpStr, j, 1)) = 0 Then
                TmpStr = Replace(TmpStr, Mid$(TmpStr, j, 1), MisKeyFix(Mid$(TmpStr, j, 1)))
            End If
        Next j
        Tmpa(i) = TmpStr
    Next i
    m_NokiaClean = Join(Tmpa)

End Sub

Private Sub NokiaAutoFixOctave()

  Dim Tmpa As Variant, TmpStr As String, TmpOct As Single, i As Long

    Tmpa = Split(m_NokiaClean)
    For i = LBound(Tmpa) To UBound(Tmpa)
        If InStr(Tmpa(i), "P") = 0 Then
            TmpStr = Tmpa(i)
            TmpOct = GetRemoveNumerals(TmpStr, False)
            If Right$(TmpStr, 1) = "-" Then
                GetRemoveChar TmpStr, False
                TmpOct = -TmpOct
            End If
            KeepInBounds 1, TmpOct, 3
            Tmpa(i) = TmpStr & TmpOct

        End If
    Next i
    m_NokiaClean = Join(Tmpa)

End Sub

Private Sub NokiaAutoFixSharp()

  Dim Tmpa As Variant, TmpStr As String, TmpOct As Single, i As Long

    Tmpa = Split(m_NokiaClean)
    For i = LBound(Tmpa) To UBound(Tmpa)
        TmpStr = Tmpa(i)
        If InStr(TmpStr, "#") > 0 Then
            If InStr("CDFGA", Mid$(TmpStr, InStr(TmpStr, "#") + 1, 1)) = 0 Then 'legal sharps
                TmpStr = Replace(TmpStr, "#", "")
                Tmpa(i) = TmpStr
            End If
        End If
    Next i
    m_NokiaClean = Join(Tmpa)

End Sub

Public Property Let NokiaAutoRepair(ByVal Immediate As Boolean)

    m_NokiaAutoRepair = Immediate

End Property

Public Property Get NokiaAutoRepair() As Boolean

    NokiaAutoRepair = m_NokiaAutoRepair

End Property

Private Sub NokiaClean(str As String)

    m_NokiaClean = NormaliseStr(str)
    NokiaStandardizePause

End Sub

Private Function nokiaFix() As Boolean

  Dim Tmpa As Variant
  Dim i As Long
  Dim BadNotes As String
  Dim BadNoteCount As Integer
  Dim CharError As String
  Dim CharFix As String
  Dim GoodChar As String
  Dim GoodOctave As Integer
  Dim DoFix As Boolean

    Tmpa = Split(m_NokiaClean)

    For i = LBound(Tmpa) To UBound(Tmpa)
        If NokiaSafeCharacter(Tmpa(i), i, CharError) Then
            'do nothing
          Else 'NOT NOKIASAFECHARACTER(TMPA(I),...
            Select Case CharError
              Case "Q", "Z"
                GoodChar = "A"
              Case "Y"
                GoodChar = "G"
              Case "W"
                GoodChar = "Rnd .7 E or A"
              Case "S"
                GoodChar = "Rnd .5 A or D"
              Case "X"
                GoodChar = "Rnd .7 C or D"
              Case "R"
                GoodChar = "Rnd .7 E or F"
              Case "T"
                GoodChar = "Rnd .6 G or F"
              Case "V"
                GoodChar = "Rnd .5 C or B"
              Case "H"
                GoodChar = "Rnd .7 B or G"
              Case "N"
                GoodChar = "Rnd .7 B or G"
              Case Else
                GoodChar = "Random legal"
            End Select

            BadNotes = BadNotes & Format$(i + 1, "@@@") & vbTab & Tmpa(i) & vbTab & "Character '" & CharError & "'" & vbTab & GoodChar & vbNewLine
            BadNoteCount = BadNoteCount + 1
        End If

        If NokiaSafeOctave(Tmpa(i), i, GoodOctave) Then
            'do nothing
          Else 'NOT NOKIASAFEOCTAVE(TMPA(I),...
            BadNotes = BadNotes & Format$(i + 1, "@@@") & vbTab & Tmpa(i) & vbTab & IIf(GoodOctave = 1, "Low", "High") & " Octave" & vbTab & "Set to " & GoodOctave & vbNewLine
            BadNoteCount = BadNoteCount + 1
        End If
        If NokiaSafeSharp(Tmpa(i), i) Then
            'do nothing
          Else 'NOT NOKIASAFESHARP(TMPA(I),...
            BadNotes = BadNotes & Format$(i + 1, "@@@") & vbTab & Tmpa(i) & vbTab & "Illegal Sharp" & vbTab & "Delete #" & vbNewLine
            BadNoteCount = BadNoteCount + 1
        End If
    Next i

    If BadNoteCount > 0 Then
        If Not m_NokiaAutoRepair Then
            If MsgBox("OK to Auto Fix and play (repairs show after playing), Cancel to do it yourself (don't play)." & vbNewLine & _
               "The following" & IIf(BadNoteCount > 1, " " & BadNoteCount, "") & " note" & IIf(BadNoteCount > 1, "s were", " is") & " unacceptable:" & vbNewLine & _
               "#" & vbTab & "Note" & vbTab & "Reason" & vbTab & vbTab & "Repair" & vbNewLine & _
               BadNotes, vbExclamation + vbOKCancel, "NOKIA CODE ERRORS(list may be truncated)") = vbOK Then
                DoFix = True
            End If
          Else 'NOT NOT...
            DoFix = True
        End If
      Else 'NOT BADNOTECOUNT...
        nokiaFix = True
    End If

    If DoFix Then
        NokiaAutoFixOctave
        NokiaAutoFixCharacter
        NokiaAutoFixSharp
        nokiaFix = True
    End If

End Function

Public Function NokiaOctaveShift(ByVal T$, Dir As Integer) As String

  Dim i As Long, TmpVal As String, Tmpa As Variant, TmpStr As String
  Dim LOctave As Integer

    Tmpa = Split(T$)
    For i = LBound(Tmpa) To UBound(Tmpa)
        If Not InStr(Tmpa(i), "P") Then
            TmpStr = Tmpa(i)
            LOctave = GetRemoveNumerals(TmpStr, False)
            Tmpa(i) = TmpStr & (LOctave + Dir)
        End If
    Next i
    NokiaOctaveShift = Join(Tmpa)

End Function

Public Sub NokiaRead(str As String)

  'Copyright 2002 Roger Gilchrist
  'Play Nokia code

  Dim Notes As Variant, i As Long
  Dim tran As String

    'Nokia format <Duration>-  (alt code <Duration>P)
    'Nokia Note Format <Duration>[...][#]<Note><Octave>
    NokiaClean str
    If Not nokiaFix Then
        Exit Sub '>---> Bottom
    End If
    m_BasicaClean = ""
    m_RTTTLClean = ""
    Notes = Split(m_NokiaClean)
    m_abort = False
    m_BasicaClean = "MN T" & m_Tempo
    RTTTLWrite RTBeat, CStr(m_Tempo)

    'ClearScore
    For i = LBound(Notes) To UBound(Notes)
        tran = Nokia2Basica(CStr(Notes(i)), True)

        RTTTLWrite RTNote, Nokia2RTTTL(CStr(Notes(i)))

        ShowMsg "Player Piano Mode " & Format$(Percent(i, UBound(Notes)), "###") & "% " & Notes(i)
        m_BasicaClean = m_BasicaClean & " " & tran
        PlayKBNote Notes(i), False, True
        'End If
        If m_abort Then
            Exit For '>---> Next
        End If
    Next i

    ShowMsg ""
    m_BasicaClean = BasicaClean(m_BasicaClean)
    RTTTLClean

End Sub

Private Function NokiaSafeCharacter(ByVal n, notenumber As Long, CharError As String) As Boolean

  Dim i As Integer

    For i = 1 To Len(n)
        If InStr("1234567890-#.PABCDEFG", Mid$(n, i, 1)) = 0 Then
            CharError = Mid$(n, i, 1)
            NokiaSafeCharacter = False
            Exit For '>---> Next
          Else 'NOT INSTR("1234567890-#.PABCDEFG",...
            NokiaSafeCharacter = True
        End If
    Next i

End Function

Private Function NokiaSafeOctave(ByVal n, notenumber As Long, GoodOctave As Integer) As Boolean

  Dim TmpOct As Integer, TmpStr As String

    If InStr(n, "P") Then
        NokiaSafeOctave = True
        Exit Function '>---> Bottom
    End If
    TmpStr = n
    TmpOct = GetRemoveNumerals(TmpStr, False) ' + 4   'get the Octave value and '+ 4' for nokia range
    If Right$(TmpStr, 1) = "-" Then
        TmpStr = GetRemoveChar(TmpStr, False)
        TmpOct = -TmpOct
    End If

    If TmpOct < 1 Or TmpOct > 3 Then
        NokiaSafeOctave = False
        GoodOctave = TmpOct
        KeepInBounds 1, GoodOctave, 3
      Else 'NOT TMPOCT...
        NokiaSafeOctave = True
    End If

End Function

Private Function NokiaSafeSharp(ByVal n, notenumber As Long) As Boolean

  Dim i As Integer

    If InStr(n, "#") = 0 Then
        NokiaSafeSharp = True
      Else 'NOT INSTR(N,...
        If InStr("CDFGA", Mid$(n, InStr(n, "#") + 1, 1)) Then 'legal sharps
            NokiaSafeSharp = True
          Else 'NOT INSTR("CDFGA",...
            NokiaSafeSharp = False
        End If
    End If

End Function

Private Sub NokiaStandardizePause()

  Dim Tmpa As Variant, i As Long, TmpStr As String

    Tmpa = Split(m_NokiaClean)
    For i = LBound(Tmpa) To UBound(Tmpa)
        If Right$(Tmpa(i), 1) = "-" Then
            TmpStr = Tmpa(i)
            TmpStr = Replace(TmpStr, "-", "P")
            Tmpa(i) = TmpStr
        End If
    Next i
    m_NokiaClean = Join(Tmpa)

End Sub

Private Function NormaliseStr(str As String) As String

  'Copyright 2002 Roger Gilchrist
  'strip string of Tab NewLine CR LF and extra spaces

    str = Replace(str, vbTab, " ")
    str = Replace(str, vbNewLine, " ")
    str = Replace(str, vbCr, " ")
    str = Replace(str, vbLf, " ")
    str = Replace(str, ",", "")
    str = Trim$(UCase$(Replace(str, "  ", " ")))
    Do While InStr(str, "  ")
        str = Replace(str, "  ", " ")
    Loop
    NormaliseStr = str

End Function

Public Function NoteTextOutPut(Mode As TextTypeOutPut) As String

  'Copyright 2002 Roger Gilchrist

  Dim RTLD As Integer, RTLO As Integer, RTLB As Integer, DBUG As String

    If Playable Then

        Select Case Mode
          Case Basica
            If OctaveMember = "P" Then
                'Basica       "P<Duration>"
                NoteTextOutPut = "P" & DurationVal
              Else 'NOT OCTAVEMEMBER...
                'Basica Note Format "L<Duration> O<Octave> [#]<Note>[...]"
                NoteTextOutPut = "L" & DurationVal & "O" & Octave & BasicaInvertSharp(OctaveMember) & IIf(m_DottedNote, ".", "")
                ' NoteTextOutPut = "L" & Ldur & "O" & Loct & BasicaInvertSharp(Lnote) & IIf(Ldot, ".", "")
            End If
          Case Nokia

            If OctaveMember = "P" Then
                'Nokia format <Duration>-  (alt code <Duration>P) this class only uses -
                NoteTextOutPut = DurationVal & "-"
              Else 'NOT OCTAVEMEMBER...
                'Nokia Note Format <Duration>[...][#]<Note><Octave>
                NoteTextOutPut = DurationVal & OctaveMember & IIf(m_DottedNote, ".", "") & (Octave - 2)
                'NoteTextOutPut = Ldur & Lnote & IIf(Ldot, ".", "") & (Loct - 2)
            End If
          Case RTTTL
            'Title:d= ,o= ,b= : [d]n[#][.][o] ,

            If OctaveMember = "P" Then
                '[d]p
                If DurationVal = RTLD Then
                    NoteTextOutPut = "p"
                  Else 'NOT DURATIONVAL...
                    NoteTextOutPut = DurationVal & "p"
                End If
              Else 'NOT OCTAVEMEMBER...
                ' [d]n[#][.][o] ,
                NoteTextOutPut = LCase$(OctaveMember) & IIf(m_DottedNote, ".", "")
                If DurationVal <> RTLD Then
                    NoteTextOutPut = DurationVal & NoteTextOutPut
                End If
                If Octave <> RTLO Then
                    NoteTextOutPut = NoteTextOutPut & Octave
                End If
            End If
            'this is not quite the right spot for this but for now this is where I've put it
            RTLD = RTTTLGetDefDuration
            RTLO = RTTTLGetDefOctave
            RTLB = RTTTLGetDefBeat
            RTTTLWrite RTTitle, RTTTLGetDefTitle
            RTTTLWrite RTLen, RTTTLGetDefDuration
            RTTTLWrite RTOct, RTTTLGetDefOctave
            RTTTLWrite RTBeat, RTTTLGetDefBeat
            RTTTLWrite RTNote, NoteTextOutPut

        End Select

    End If

End Function

Public Property Get Octave() As Integer

  'Copyright 2002 Roger Gilchrist
  'ReadOnly public interface

    Octave = m_Octave

End Property

Public Property Get OctaveMember() As String

  'Copyright 2002 Roger Gilchrist
  'ReadOnly public interface

    If Playable Then
        If m_OctaveMember > -1 Then
            If m_IsBlack Then
                OctaveMember = KBLet(m_OctaveMember)
              ElseIf M_IsWhite Then 'M_ISBLACK = FALSE
                OctaveMember = KWLet(m_OctaveMember)
              Else 'M_ISWHITE = FALSE
                OctaveMember = "P"
            End If
        End If
    End If

End Property

Public Property Get OffSetEnds() As Variant

    OffSetEnds = m_OffSetEnds

End Property

Public Property Let OffSetEnds(ByVal UseOffSet As Variant)

  'Copyright 2002 Roger Gilchrist
  'can be set to false but recommend leaving True it makes the MouseMove messageing work better

    m_OffSetEnds = UseOffSet

End Property

Public Property Get PauseKeyOn() As Boolean

    PauseKeyOn = m_PauseKeyOn

End Property

Public Property Let PauseKeyOn(ByVal OnOff As Boolean)

    m_PauseKeyOn = OnOff

End Property

Private Function Percent(a, b) As Single

  'calculate percentage given two numbers

  Dim c As Single

    If a > b Then
        a = b
    End If
    If b = 0 Then
        Percent = 0
        Exit Function '>---> Bottom
    End If
    c = Int(a / b * 100)
    If c > 100 Then
        c = 100
    End If
    Percent = c

End Function

Public Property Get Playable() As Boolean

  'Copyright 2002 Roger Gilchrist
  'ReadOnly The cursor is over a key so a note can be selected

    Playable = (m_IsBlack Or M_IsWhite Or M_IsPause)

End Property

Private Sub playbeep(F As Single, V As Long)

    KeepInBoundsSimple 36, F, 32768
    'Beep F, V
    'If you don't want to support Win.9x then uncomment the previous line and
    'delete the next line and all its support code (BRGetPlatform, BRWin9xBeep, BRPcSpeakerBeep and bottom of Declarations section)
    BRPcSpeakerBeep F, V
    DoEvents
    'because API Sleep/Beep are asynchronous there is
    'a chance to read the Abort command just after it

End Sub

Private Sub PlayKBNote(ByVal KeyName As String, Optional BasicaMode As Boolean = True, Optional PlayerPianoMode As Boolean = False)

  'Copyright 2002 Roger Gilchrist
  'updated to make player piano calls

  Dim NoteFreq As Single
  Dim DurationValue As Long
  Dim LOctave As Integer
  Dim FakeKey As Long
  Dim DBUG As String

    DBUG = KeyName

    DurationValue = GetDurationValue(KeyName)
    If InStr(KeyName, "P") Then
        playsleep DurationValue
      Else 'NOT INSTR(KEYNAME,...
        If PlayerPianoMode Then
            FakeKey = KeyNoFromNote(KeyName) + IIf(BasicaMode, 0, 14)
        End If
        LOctave = GetRemoveNumerals(KeyName, False) + IIf(BasicaMode, 2, 4) 'get the Octave value and '+ 4'
        NoteFreq = FrequencyFromNoteName(KeyName & LOctave)
        If PlayerPianoMode Then
            Key FakeKey
        End If
        playbeep NoteFreq, DurationValue    ', 0)    'play the note
        If PlayerPianoMode Then
            m_WhitePressed = False
            m_RightBlackPressed = False
            Key KeyStack 'pop and redraw
        End If
    End If

End Sub

'plays a sleep (pause) durating V
Private Sub playsleep(D As Long)

    Sleep D
    DoEvents
    'because API Sleep/Beep are asynchronous there is
    'a chance to read the Abort command just after it

End Sub

Private Function RandomNoteName() As String

    RandomNoteName = Chr$(Asc("A") + Int(Rnd * 7))

End Function

Public Sub ReassignButtonValues(ParamArray BValues())

    MsgBox "Not implemented. Don't call yet", vbInformation, "Sorry About This"
    'Not implemented Don't call yet
    'Bvalue will be an array or numbers including -1 to allow default values to be assigned

Exit Sub

    '  Dim I As Integer
    '
    '    For I = LBound(BValues) To UBound(BValues)
    '        If Len(BValues(I)) Then
    '            m_DurationSettings(I) = BValues(I)
    '        End If
    '    Next I

End Sub

Private Function RemoveDots(T$) As Long

  'Remove dots from string and return a count of them
  'simplified version
  'Must be called after GetRemoveNumerals as it will damage decimal numbers

  Dim tmp As String

    tmp$ = T$
    T$ = Replace(T, ".", "")
    RemoveDots = Len(tmp$) - Len(T$)

End Function

Public Sub Resize()

  'Copyright 2002 Roger Gilchrist

    ReSizing = True
    KeyBoardDraw
    ReSizing = False

End Sub

Private Function RotateNote(n$, Dir As Integer) As String

    If InStr("ABCDEFG", n$) Then
        RotateNote = Chr$(Asc(n$) + Dir)
        If RotateNote = "@" Then ' if tmpval is beyond legal note names then switch to other end
            RotateNote = "G"
        End If
        If RotateNote = "H" Then
            RotateNote = "A"
        End If
      Else 'NOT INSTR("ABCDEFG",...
        RotateNote = n$
    End If

End Function

Public Function RTTL2Composer(RTTL$) As String

  Dim Title As String
  Dim DefLen As Integer
  Dim DefOct As Integer
  Dim DefTempo As Integer
  Dim i As Long
  Dim Tmpa As Variant
  Dim TmpStr As String

    Title = Left$(RTTL$, InStr(RTTL, ":"))
    RTTL = Mid$(RTTL, Len(Title) + 1)
    RTTL = UCase$(RTTL)
    TmpStr = Left$(RTTL$, InStr(RTTL, ",") - 1)
    RTTL = Mid$(RTTL, Len(TmpStr) + 2)
    DefLen = GetRemoveNumerals(TmpStr, False)
    TmpStr = Left$(RTTL$, InStr(RTTL, ",") - 1)
    RTTL = Mid$(RTTL, Len(TmpStr) + 2)
    DefOct = GetRemoveNumerals(TmpStr, False)
    TmpStr = Left$(RTTL$, InStr(RTTL, ":") - 1)
    RTTL = Mid$(RTTL, Len(TmpStr) + 2)
    m_Tempo = GetRemoveNumerals(TmpStr, False)

    Tmpa = Split(RTTL, ",")
    Tmpa = ArrayRemoveBlankMembers(Tmpa)
    For i = LBound(Tmpa) To UBound(Tmpa)
        If Left$(Tmpa(i), 1) = "P" Then
            If Val(Right$(Tmpa(i), 1)) = 0 Then
                Tmpa(i) = Tmpa(i) & DefLen
            End If
          Else 'NOT LEFT$(TMPA(I),...
            If Val(Left$(Tmpa(i), 1)) = 0 Then
                Tmpa(i) = DefLen & Tmpa(i)
            End If
            If Val(Right$(Tmpa(i), 1)) = 0 Then
                Tmpa(i) = Tmpa(i) & DefOct
            End If
        End If
    Next i
    RTTL2Composer = UCase$(Join(Tmpa))
    'Stop

End Function

Public Function RTTTL2Basica(RTTL$) As String

  Dim Title As String
  Dim DefLen As Integer
  Dim DefOct As Integer
  Dim DefTempo As Integer
  Dim i As Long
  Dim Tmpa As Variant
  Dim TmpStr As String
  Dim TmpStr2 As String

    m_RTTTLClean = RTTL
    RTTL$ = RTTTLClean

    Title = Left$(RTTL$, InStr(RTTL, ":"))

    RTTL = Mid$(RTTL, Len(Title) + 1)
    RTTL = UCase$(RTTL)
    TmpStr = Left$(RTTL$, InStr(RTTL, ",") - 1)
    RTTL = Mid$(RTTL, Len(TmpStr) + 2)
    DefLen = GetRemoveNumerals(TmpStr, False)
    TmpStr = Left$(RTTL$, InStr(RTTL, ",") - 1)
    RTTL = Mid$(RTTL, Len(TmpStr) + 2)
    DefOct = GetRemoveNumerals(TmpStr, False) ' - 2
    TmpStr = Left$(RTTL$, InStr(RTTL, ":") - 1)
    RTTL = Mid$(RTTL, Len(TmpStr) + 2)
    m_Tempo = GetRemoveNumerals(TmpStr, False)

    Tmpa = Split(RTTL, ",")
    Tmpa = ArrayRemoveBlankMembers(Tmpa)
    For i = LBound(Tmpa) To UBound(Tmpa)
        If Right$(Tmpa(i), 1) = "P" Then
            If Val(Left$(Tmpa(i), 1)) = 0 Then
                Tmpa(i) = DefLen & Tmpa(i)
            End If
          Else 'NOT RIGHT$(TMPA(I),...

            If Val(Left$(Tmpa(i), 1)) = 0 Then
                Tmpa(i) = DefLen & Tmpa(i)
              Else 'NOT VAL(LEFT$(TMPA(I),...
                TmpStr = Tmpa(i)
                TmpStr2 = GetRemoveNumerals(TmpStr, True)
                Tmpa(i) = TmpStr2 & TmpStr
            End If

            If Val(Right$(Tmpa(i), 1)) = 0 Then
                Tmpa(i) = Tmpa(i) & DefOct - 3
              Else 'NOT VAL(RIGHT$(TMPA(I),...
                TmpStr = Tmpa(i)
                TmpStr2 = GetRemoveNumerals(TmpStr, False) - 3
                Tmpa(i) = TmpStr & TmpStr2
            End If
        End If
    Next i
    For i = LBound(Tmpa) To UBound(Tmpa)
        Tmpa(i) = Nokia2Basica(Tmpa(i))
    Next i

    TmpStr = Join(Tmpa)

    RTTTL2Basica = "T" & m_Tempo & " " & TmpStr
    'Stop

End Function

Public Function RTTTLClean() As String

  Dim TmpStr As String, LDur As Integer, LOct As Integer, Ltempo As Integer, i As Long
  Dim Title As String, TitlePos As Long, Junk As Integer
  Dim Tmpa As Variant, Counter(64) As Integer, TmpNum As Integer
  Dim DBUG As String

    DBUG = m_RTTTLClean
    TitlePos = InStr(m_RTTTLClean, ":")
    If TitlePos Then

        Title = Left$(m_RTTTLClean, TitlePos)
        If InStr(Title, "=") Then 'blocks generated '"b=" & m_Tempo' becoming title
            Title = m_RTTTLtitle
        End If
      Else 'TITLEPOS = FALSE
        Title = m_RTTTLtitle
    End If

    LDur = RTTTLGetDefDuration
    LOct = RTTTLGetDefOctave
    Ltempo = RTTTLGetDefBeat

    TmpStr = UCase$(m_RTTTLClean)
    TmpStr = Mid$(TmpStr, InStr(TmpStr, "B=") + Len(Ltempo) + 4)
    Tmpa = Split(TmpStr, ",")
    If LDur = 0 Then
        'get most common duration value
        For i = LBound(Tmpa) To UBound(Tmpa)
            TmpStr = Tmpa(i)
            If IsNumeric(Left$(TmpStr, 1)) Then
                TmpNum = GetRemoveNumerals(TmpStr, True)
                Counter(TmpNum) = Counter(TmpNum) + 1
            End If
        Next i
        LDur = CounterMember(Counter, LargestMember)
    End If
    For i = LBound(Tmpa) To UBound(Tmpa)
        Counter(i) = 0
    Next i
    If LOct = 0 Then
        'get most common octave
        For i = LBound(Tmpa) To UBound(Tmpa)
            TmpStr = Tmpa(i)
            If IsNumeric(Right$(TmpStr, 1)) Then
                TmpNum = GetRemoveNumerals(TmpStr, False)
                Counter(TmpNum) = Counter(TmpNum) + 1
            End If
        Next i
        LOct = CounterMember(Counter, LargestMember)

    End If

    For i = LBound(Tmpa) To UBound(Tmpa) 'delete duplicates of Loct and LDur
        TmpStr = Tmpa(i)
        If Val(Left$(TmpStr, Len(CStr(LDur)))) = LDur Then
            Junk = GetRemoveNumerals(TmpStr, True)
            Tmpa(i) = TmpStr
        End If
        TmpStr = Tmpa(i)
        If Val(Right$(TmpStr, Len(CStr(LOct)))) = LOct Then
            Junk = GetRemoveNumerals(TmpStr, False)
            Tmpa(i) = TmpStr
        End If
    Next i
    RTTTLClean = StrConv(Title, vbProperCase) & ":d=" & LDur & ",o=" & LOct & ",b=" & Ltempo & ":" & LCase$(Join(Tmpa, ","))
    RTTTLClean = Replace(RTTTLClean, "::", ":")
    If Right$(RTTTLClean, 1) = "," Then
        RTTTLClean = Left$(RTTTLClean, Len(RTTTLClean) - 1)
    End If
    m_RTTTLClean = RTTTLClean

End Function

Private Function RTTTLGetDefBeat() As Integer

  Dim TmpStr As String, pos As Long

    TmpStr = UCase$(m_RTTTLClean)
    pos = InStr(TmpStr, "B=")
    If pos Then
        TmpStr = Mid$(TmpStr, pos + 2)
        RTTTLGetDefBeat = GetRemoveNumerals(TmpStr, True)
      Else 'POS = FALSE
        RTTTLGetDefBeat = m_Tempo
    End If

End Function

Private Function RTTTLGetDefDuration() As Integer

  Dim TmpStr As String, pos As Long

    TmpStr = UCase$(m_RTTTLClean)
    pos = InStr(TmpStr, "D=")
    If pos Then
        TmpStr = Mid$(TmpStr, pos + 2)
        RTTTLGetDefDuration = GetRemoveNumerals(TmpStr, True)
      Else 'POS = FALSE
        RTTTLGetDefDuration = 0
    End If

End Function

Private Function RTTTLGetDefOctave() As Integer

  Dim TmpStr As String, pos As Long

    TmpStr = UCase$(m_RTTTLClean)
    pos = InStr(TmpStr, "O=")
    If pos Then
        TmpStr = Mid$(TmpStr, pos + 2)
        RTTTLGetDefOctave = GetRemoveNumerals(TmpStr, True)
      Else 'POS = FALSE
        RTTTLGetDefOctave = 0
    End If

End Function

'
'Private Sub SetUnassigned()
'
'  'Called from MouseDuration
'  'Unassigned Durations set to current system value
'  'Stops unassigned Shifts from having any effect
'  'If you assign a value comment out in this routine
'  'NOT FULLY IMPLMENTED YET the values are set in Class_Initialize only
'  'm_DurationSettings(ButtonL)= m_Duration
'
'    'm_DurationSettings(AltButtonL) = m_Duration
'    'm_DurationSettings(CtrlButtonL)= m_Duration
'    'm_DurationSettings(ShiftButtonL)= m_Duration
'    m_DurationSettings(AltCtrlButtonL) = m_Duration
'    m_DurationSettings(ShiftCtrlButtonL) = m_Duration
'    m_DurationSettings(ShiftAltButtonL) = m_Duration
'    m_DurationSettings(ShiftAltCtrlButtonL) = m_Duration
'
'    'm_DurationSettings(ButtonR)= m_Duration
'    m_DurationSettings(AltButtonR) = m_Duration
'    'm_DurationSettings(CtrlButtonR)= m_Duration
'    'm_DurationSettings(ShiftButtonR)= m_Duration
'    m_DurationSettings(AltCtrlButtonR) = m_Duration
'    m_DurationSettings(ShiftCtrlButtonR) = m_Duration
'    m_DurationSettings(ShiftAltButtonR) = m_Duration
'    m_DurationSettings(ShiftAltCtrlButtonR) = m_Duration
'
'    'm_DurationSettings(ButtonM)= m_Duration
'    m_DurationSettings(AltButtonM) = m_Duration
'    'm_DurationSettings(CtrlButtonM)= m_Duration
'    'm_DurationSettings(ShiftButtonM)= m_Duration
'    m_DurationSettings(AltCtrlButtonM) = m_Duration
'    m_DurationSettings(ShiftCtrlButtonM) = m_Duration
'    m_DurationSettings(ShiftAltButtonM) = m_Duration
'    m_DurationSettings(ShiftAltCtrlButtonM) = m_Duration
'
'End Sub
Private Function RTTTLGetDefTitle() As String

  Dim TmpStr As String, pos As Long, pos2 As Long

    TmpStr = UCase$(m_RTTTLClean)
    pos = InStr(TmpStr, ":")
    pos2 = InStr(pos + 1, TmpStr, ":")
    If pos2 > pos Then ' there were two : so in proper format
        RTTTLGetDefTitle = StrConv(Left$(TmpStr, pos - 1), vbProperCase)
      Else 'NOT POS2...
        RTTTLGetDefTitle = m_RTTTLtitle
    End If

End Function

Public Sub RTTTLRead(str As String)

    BasicaRead RTTTL2Basica(str)

End Sub

Public Property Get RTTTLString() As String

    RTTTLString = m_RTTTLClean

End Property

Public Property Let RTTTLString(ByVal StrIn As String)

    m_RTTTLClean = StrIn

End Property

Public Property Get RTTTLTitle() As String

    RTTTLTitle = m_RTTTLtitle

End Property

Public Property Let RTTTLTitle(ByVal vNewValue As String)

    m_RTTTLtitle = vNewValue

End Property

Public Function RTTTLValid(str As String) As Boolean

  Dim Tmpm_RTTTLClean As String

    Tmpm_RTTTLClean = str
    m_RTTTLClean = str
    RTTTLClean
    'On Error Resume Next
    If InStr(UCase$(str), UCase$(RTTTLGetDefTitle)) Then
        If InStr(str, RTTTLGetDefDuration) Then
            If InStr(str, RTTTLGetDefOctave) Then
                If InStr(str, RTTTLGetDefBeat) Then
                    RTTTLValid = True
                End If
            End If
        End If
    End If
    If Not RTTTLValid Then
        m_RTTTLClean = Tmpm_RTTTLClean
    End If

End Function

Private Function RTTTLWrite(WriteWhat As RTTTLPartEnum, str As String) As String

  'Title:d=#,o=#,b=#:n

  Dim pos As Long
  Dim pos2 As Long
  Dim defval As Integer
  Dim TmpStr As String

    Select Case WriteWhat
      Case RTTitle
        pos = InStr(m_RTTTLClean, ":d")
        If pos Then
            m_RTTTLClean = str & ":" & Mid$(m_RTTTLClean, pos + 1)
          Else 'POS = FALSE
            m_RTTTLClean = str & ":" & m_RTTTLClean
        End If
      Case RTLen
        pos = InStr(m_RTTTLClean, "d=")
        If pos Then
            defval = RTTTLGetDefDuration
            If defval <> CInt(str) Then
                m_RTTTLClean = Replace(m_RTTTLClean, "d=" & defval, "d=" & str)
            End If
          Else 'POS = FALSE
            pos = InStr(m_RTTTLClean, ":") 'end of title
            TmpStr = "d=" & str & ","
            m_RTTTLClean = Left$(m_RTTTLClean, pos) & TmpStr & Mid$(m_RTTTLClean, pos + 1)
        End If
      Case RTOct
        pos = InStr(m_RTTTLClean, "o=")
        If pos Then
            defval = RTTTLGetDefDuration
            If defval <> CInt(str) Then
                m_RTTTLClean = Replace(m_RTTTLClean, "o=" & defval, "o=" & str)
            End If
          Else 'POS = FALSE
            pos = InStr(m_RTTTLClean, ",") 'end of d
            TmpStr = "o=" & str & ","
            m_RTTTLClean = Left$(m_RTTTLClean, pos) & TmpStr & Mid$(m_RTTTLClean, pos + 1)
        End If

      Case RTBeat
        pos = InStr(m_RTTTLClean, "b=")
        If pos Then
            defval = RTTTLGetDefBeat
            If defval <> CInt(str) Then
                m_RTTTLClean = Replace(m_RTTTLClean, "b=" & defval, "b=" & str)
            End If
          Else 'POS = FALSE
            pos = InStr(m_RTTTLClean, ",") 'end of d
            pos = InStr(pos + 1, m_RTTTLClean, ",") 'end of o
            TmpStr = "b=" & str & ":"
            m_RTTTLClean = Left$(m_RTTTLClean, pos) & TmpStr & Mid$(m_RTTTLClean, pos + 2)
        End If
      Case RTNote
        m_RTTTLClean = m_RTTTLClean & str & ","
    End Select
    m_RTTTLClean = Replace(m_RTTTLClean, "::", ":")
    RTTTLWrite = m_RTTTLClean

End Function

Private Function ShiftReader(Shift As Integer) As ShiftDescriptor

  'MSDN Library Visual Studio 6 (Modified)
  'I know you could just read them directly but this makes code readable
  'by using Enums

  Dim ShiftDown As Boolean, AltDown As Boolean, CtrlDown As Boolean

    ShiftDown = (Shift And vbShiftMask) > 0
    AltDown = (Shift And vbAltMask) > 0
    CtrlDown = (Shift And vbCtrlMask) > 0
    If ShiftDown And CtrlDown And AltDown Then
        ShiftReader = KShiftAltCtrl
      ElseIf ShiftDown And AltDown Then 'NOT SHIFTDOWN...
        ShiftReader = KShiftAlt
      ElseIf ShiftDown And CtrlDown Then 'NOT SHIFTDOWN...
        ShiftReader = KShiftCtrl
      ElseIf CtrlDown And AltDown Then 'NOT SHIFTDOWN...
        ShiftReader = KAltCtrl
      ElseIf ShiftDown Then 'NOT CTRLDOWN...
        ShiftReader = KShift
      ElseIf CtrlDown Then 'SHIFTDOWN = FALSE
        ShiftReader = KCtrl
      ElseIf AltDown Then 'CTRLDOWN = FALSE
        ShiftReader = KAlt
      ElseIf Shift = 0 Then 'ALTDOWN = FALSE
        ShiftReader = KNoShift
    End If

End Function

Private Sub ShowMsg(msg As String)

  'Preserve original Caption and Previous message

  Static OriginalMsg As String
  Static Prevmsg As String

    If msg <> Prevmsg Then
        If msg = "" Then 'reinstate original caption if msg is blank
            msg = OriginalMsg
        End If
        If m_ContainerIsFrame Then
            If Prevmsg = "" Then
                OriginalMsg = M_KBFrame.Caption
            End If
            M_KBFrame.Caption = msg
          Else 'M_CONTAINERISFRAME = FALSE
            If m_ShowOnFormCaption Then
                If Prevmsg = "" Then
                    OriginalMsg = M_KBFrame.Caption
                End If
                M_KBForm.Caption = msg
            End If
        End If
    End If
    Prevmsg = msg

End Sub

Public Property Get ShowNokiaSafeOctaves() As Boolean

    ShowNokiaSafeOctaves = m_ShowNokiaOctave

End Property

Public Property Let ShowNokiaSafeOctaves(ByVal ShowNokiaOctave As Boolean)

    m_ShowNokiaOctave = ShowNokiaOctave

End Property

Public Property Let Tempo(ByVal vNewValue As Integer)

    m_Tempo = vNewValue

End Property

Public Property Get Tempo() As Integer

    Tempo = m_Tempo

End Property

Public Property Get UseCKeyColourTags() As Boolean

    UseCKeyColourTags = m_UseColorTags

End Property

Public Property Let UseCKeyColourTags(ByVal UseCTag As Boolean)

  'Copyright 2002 Roger Gilchrist
  'Allow/Prevent Colour tagging; Any Appearance change should call KeyBoardDraw

    m_UseColorTags = UseCTag
    KeyBoardDraw

End Property

Public Property Let UseFormCaption(ByVal vNewValue As Boolean)

    m_ShowOnFormCaption = vNewValue

End Property

Public Property Get UseFormCaption() As Boolean

    UseFormCaption = m_ShowOnFormCaption

End Property

Public Property Get WholeKeyTag() As Boolean

    WholeKeyTag = m_WholeKeyTag

End Property

Public Property Let WholeKeyTag(ByVal WholeKey As Boolean)

  'Copyright 2002 Roger Gilchrist
  'Set TagKeyStyle; Any Appearance change should call KeyBoardDraw

    m_WholeKeyTag = WholeKey
    KeyBoardDraw

End Property

':) Ulli's VB Code Formatter V2.13.6 (9/10/2002 11:46:53 AM) 379 + 3431 = 3810 Lines
Public Function BasicaRealTranspose(str As String, TransposeFrom As String, TransposeTo As String) As String
'I have to check this out with a musical friend to check that I have the right names
'but meanwhile
'Based on
'1001 Things To Do With Your IBM PC' by Mark R. Sawusch and Tan A. Summers  1984 TAB Books Inc) p 148 for KeyOfXX stuff
'This should be a musically valid version of what the BasicaNoteShift routine does
'at present it assumes the tune is in Key of C once I know more I may modify it to allow other base keys.
Dim FromKey As Variant, ToKey As Variant
Dim FromPos As Integer
Dim Tmpa As Variant, i As Long, TmpStr As String
Dim KeyofC As Variant, KeyOfD As Variant, KeyOfGFlat As Variant, KeyOfBFlat As Variant
Dim KeyofEFlat As Variant, KeyOfGSharp As Variant, KeyOfCSharp As Variant, KeyOfB As Variant

Dim KeyofE As Variant, KeyOfA As Variant, KeyOfD2 As Variant, KeyOfGFlat2 As Variant

KeyofC = Array("C", "C#", "D-", "D", "D#", "E-", "E", "F", "F#", "G-", "G", "G#", "A-", "A", "A#", "B-", "B", "C#", "D")
KeyOfD = Array("D", "D#", "E", "E", "F", "G-", "G", "G", "G#", "A", "A", "A#", "B", "B", "C", "D-", "D", "D", "D#", "E", "E", "F")
KeyOfGFlat = Array("G-", "G", "G", "G#", "A", "A", "A#", "B", "B", "C", "D", "D#", "E-", "E", "F", "F", "F#", "G", "G#", "A-", "A", "A#")
KeyOfBFlat = Array("B-", "B", "C", "C", "C#", "D#", "E", "E", "F", "F#", "G-", "G", "A-", "A", "A", "A#", "B", "B", "C", "C#", "D-", "D")
KeyofEFlat = Array("E-", "E", "E", "F", "F#", "G-", "G", "A-", "A", "A", "A#", "B", "B", "C", "C#", "D-", "D", "E", "F", "F", "F#", "G", "G")
KeyOfGSharp = Array("G#", "A", "A#", "B-", "B", "C", "C", "C#", "D", "D", "D#", "F", "F#", "G-", "G", "G#", "A-", "A", "B-", "B", "B", "C")
KeyOfCSharp = Array("C#", "D-", "D", "D#", "E-", "E", "F#", "G", "G", "G#", "A", "A", "A#", "B", "C", "C", "C#", "D", "D", "D#", "E", "E", "F")
KeyOfB = Array("B", "C", "C#", "D-", "D", "D#", "E-", "E", "F", "F#", "G#", "A", "A", "A#", "B", "B", "C", "D-", "D", "D", "D#", "E")
KeyofE = Array("E", "F", "F#", "G-", "G", "A-", "A", "A", "A#", "B", "B", "C", "D-", "D", "D", "D#", "E", "E", "F", "F#", "G-", "G")
KeyOfA = Array("A", "A#", "B-", "B", "C", "C", "C#", "D", "D#", "E-", "E", "F", "F", "F#", "G", "G", "G#", "A#", "B", "B", "C", "C#", "D-")
KeyOfD2 = Array("D", "E-", "E", "E", "F", "F#", "G-", "G", "G#", "A-", , "A", "B-", "B", "B", "C", "C#", "D-", "D", "E-", "E", "E", "F", "F#")
KeyOfGFlat2 = Array("G-", "G", "G#", "A-", "A", "B", "C", "C", "C#", "D", "D", "D#", "F", "F", "F", "F#", "G", "G", "G#", "A", "A", "A#")
Select Case TransposeFrom
Case "C": FromKey = KeyofC
End Select
Select Case TransposeTo
Case "D": ToKey = KeyOfD
Case "G-": ToKey = KeyOfGFlat
Case "B-": ToKey = KeyOfBFlat
Case "E-": ToKey = KeyofEFlat
Case "G#": ToKey = KeyOfGSharp
Case "C#": ToKey = KeyOfCSharp
Case "B": ToKey = KeyOfB
Case "E": ToKey = KeyofE
Case "A": ToKey = KeyOfA
Case "D2": ToKey = KeyOfD2
Case "G-2": ToKey = KeyOfGFlat2
End Select

Tmpa = Split(BasicaClean(str))
For i = LBound(Tmpa) To UBound(Tmpa)
TmpStr = Left(Tmpa(i), 1)
If InStr("TMLO", TmpStr) = 0 Then 'not a non-note command
FromPos = MemberOfArray(Tmpa(i), FromKey)
If FromPos > -1 Then
    Tmpa(i) = ToKey(FromPos)
End If
End If

Next
BasicaRealTranspose = Join(Tmpa)
End Function

