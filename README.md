<div align="center">

## Play


</div>

### Description

A "Play" Command for Visual Basic. This is the equivalent of the QBasic

PLAY command that enabled you to play notes through the PC speaker. This

version allows you to take advantage of the Sound card and therefore has

many advantages. It uses the MIDI interface to send individual or

multiple notes to the sound card.
 
### More Info
 
It requires the notes that you wish to play - See the code for the

exact syntax.

The notes available range from A-G

If you end the program while it is playing, it may not function

correctly until you next restart windows (or log off and log on again).


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Aidan](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/aidan.md)
**Level**          |Unknown
**User Rating**    |3.0 (15 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[VB function enhancement](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/vb-function-enhancement__1-25.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/aidan-play__1-1473/archive/master.zip)

### API Declarations

```
Private Declare Sub SleepAPI Lib "kernel32" Alias "Sleep"_
	(ByVal dwMilliseconds As Long)
Private Declare Function midiOutOpen Lib "winmm.dll"_
	(lphMidiOut As Long, ByVal uDeviceID As Long, ByVal_
dwCallback As Long, ByVal dwInstance As Long,_
ByVal dwflags As Long) As Long
Private Declare Function midiOutShortMsg Lib "winmm.dll"_
	 (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Private Declare Function midiOutClose Lib "winmm.dll"_
	(ByVal hMidiOut As Long) As Long
Private Declare Function midiOutReset Lib "winmm.dll"_
	(ByVal hMidiOut As Long) As Long
```


### Source Code

```
' Place a Textbox (Text1) and a Command Button (Command1)
' on the Form
' The following code should be placed in Form1:
' This code gives the Visual Basic equivalent of the QBasic
' PLAY command. A few extra options have been added (such as
' playing several notes simultaneously).
' I have found it difficult to stop the notes after playing them
' You can try this out by using the MN switch.
' If anyone knows how to do this, please E-Mail me at
' aidanx@yahoo.com
' Constants
Private Const Style_Normal = 0
Private Const Style_Staccato = 1
Private Const Style_Legato = 2
Private Const Style_Sustained = 3
Private Const PlayState_Disable = 0
Private Const PlayState_Enable = 1
Private Const PlayState_Auto = -1
' Types
Private Type Note
  Pitch As Long
  Length As Integer
  Volume As Long
  Style As Integer
End Type
' Variables
Private MIDIDevice As Long
Private Sub Command1_Click()
  ' The notes in the text box are played when the button
  ' is pressed
  Play Text1.Text
End Sub
Private Sub Form_Load()
  Text1.Text = "cdecdefgafga"
  Play "MDO3cdefgabO4c"
End Sub
Public Sub Play(Notes As String)
  ' Plays a note(s) using MIDI
  ' E.g. Play "T96O3L4cd.efgabO4cL8defgabO5c"
  ' Note Letter - Plays Note (C is lowest in an octave, B is highest)
  ' L + NoteLength (4 = Crotchet, 2 = Minim, etc.,
	' 0 = Play Simultaneously)
  ' N + Note Number (37 = Mid C, 38 = C#, etc.)
  ' O + Octave No. (3 = Middle - i.e. O3C = Mid. C)
  ' P + Length (Pause of Length - See "L" -
	' Without a Number = Current Note Length)
  ' T + Tempo (Crotchet Beats per Minute)
  ' V + VolumeConstant (F = Forte, O = Mezzo-Forte,
	' I = MezzoPiano, P = Piano)
  ' M + Music Style Constant (S = Staccato, N = Normal,
	' L = Legato, D = Sustained)
      ' Only the Sustained style appears to function
	' correctly as the time taken to stop a midi note
	' is not negligible
  ' If ommitted, uses last set option
  Dim CurrentNote As Long, PauseNoteLength
  Dim i As Long, LenStr As String
  Dim Note(6) As Integer, Sharp As Integer
  Dim NoteCaps As String, NoteASCII As Integer, _
PlayLength As Double
  Dim PlayNote() As Note
  Static NotFirstRun As Boolean ' Set to True if
	' it is not the first time Play has been called
  Static Octave As Integer, Tempo As Integer, _
CurrentNoteLength As Integer, CurrentVolume As Integer, _
MusicStyle As Integer
  ' Enable MIDI
  If Not EnablePlay(PlayState_Enable) Then Exit Sub
  If Not NotFirstRun Then
    NotFirstRun = True
    Octave = 3
    Tempo = 120
    CurrentVolume = 96
    CurrentNoteLength = 4
    MusicStyle = Style_Sustained
  End If
  ' Notes
    Note(0) = 9   ' A
    Note(1) = 11  ' B
    Note(2) = 0   ' C
    Note(3) = 2   ' D
    Note(4) = 4   ' E
    Note(5) = 5   ' F
    Note(6) = 7   ' G
  ' End Notes
  NoteCaps = UCase$(Notes)
  CurrentNote = -1
  i = 0
  Do Until i = Len(NoteCaps)
    i = i + 1
    NoteASCII = Asc(Mid$(NoteCaps, i, 1))
    If Chr$(NoteASCII) = "N" Then
      ' Play Note by Number
      LenStr = ""
      Do Until i = Len(NoteCaps) Or _
Val(Mid$(NoteCaps, i + 1, 1)) = 0
        LenStr = LenStr + Mid$(NoteCaps, i + 1, 1)
        i = i + 1
      Loop
      If LenStr <> "" Then
        CurrentNote = CurrentNote + 1
        ReDim Preserve PlayNote(CurrentNote)
        If Val(LenStr) <> 0 Then
          PlayNote(CurrentNote).Pitch = Val(LenStr) + 23
        Else
          PlayNote(CurrentNote).Pitch = -1
        End If
        PlayNote(CurrentNote).Length = CurrentNoteLength
        PlayNote(CurrentNote).Volume = CurrentVolume
        PlayNote(CurrentNote).Style = MusicStyle
      End If
    End If
    If NoteASCII >= 0 Then
      If Chr$(NoteASCII) = "T" Then
        ' Set Tempo
        LenStr = ""
        Do Until i = Len(NoteCaps) Or _
Val(Mid$(NoteCaps, i + 1, 1)) = 0
          LenStr = LenStr + Mid$(NoteCaps, i + 1, 1)
          i = i + 1
        Loop
        If LenStr <> "" Then
          Tempo = Val(LenStr)
        End If
      End If
      If Chr$(NoteASCII) = "." And CurrentNote >= 0 Then
        ' Make last note length 3/2 times as long
        PlayNote(CurrentNote).Length = _
PlayNote(CurrentNote).Length / 1.5
      End If
      If Chr$(NoteASCII) = "P" Then
        ' Pause
        LenStr = ""
        Do Until i = Len(NoteCaps) Or _
Val(Mid$(NoteCaps, i + 1, 1)) = 0
          LenStr = LenStr + Mid$(NoteCaps, i + 1, 1)
          i = i + 1
        Loop
        NoteASCII = -1
        If LenStr <> "" Then
          PauseNoteLength = Val(LenStr)
        Else
          PauseNoteLength = CurrentNoteLength
        End If
      End If
      If Chr$(NoteASCII) = "L" Then
        ' Set Length
        LenStr = ""
        Do Until i = Len(NoteCaps) Or _
(Val(Mid$(NoteCaps, i + 1, 1)) = 0 And _
Mid$(NoteCaps, i + 1, 1) <> "0")
          LenStr = LenStr + Mid$(NoteCaps, i + 1, 1)
          i = i + 1
        Loop
        If LenStr <> "" Then
          CurrentNoteLength = Val(LenStr)
        End If
      End If
      If Chr$(NoteASCII) = "O" Then
        ' Set Octave
        If i < Len(NoteCaps) Then
          NoteASCII = Asc(Mid$(NoteCaps, i + 1, 1))
          If NoteASCII > 47 And NoteASCII < 55 Then
            Octave = NoteASCII - 48
            i = i + 1
          End If
        End If
      End If
    End If
    If (NoteASCII > 64 And NoteASCII < 73) Or NoteASCII = -1 Then
      ' Select Note
      Sharp = 0
      If NoteASCII <> -1 Then
        If i < Len(NoteCaps) Then
          If Mid$(NoteCaps, i + 1, 1) = "#" Or _
Mid$(NoteCaps, i + 1, 1) = "+" Then
            i = i + 1
            Sharp = 1
          ElseIf Mid$(NoteCaps, i + 1, 1) = "-" Then
            i = i + 1
            Sharp = -1
          End If
        End If
      End If
      CurrentNote = CurrentNote + 1
      ReDim Preserve PlayNote(CurrentNote)
      If NoteASCII <> -1 Then
        PlayNote(CurrentNote).Pitch = (Octave * 12) + _
Note(NoteASCII - 65) + Sharp + 24
        PlayNote(CurrentNote).Length = CurrentNoteLength
      Else
        PlayNote(CurrentNote).Pitch = -1
        PlayNote(CurrentNote).Length = PauseNoteLength
      End If
      PlayNote(CurrentNote).Volume = CurrentVolume
      PlayNote(CurrentNote).Style = MusicStyle
    End If
    If NoteASCII > -1 Then
      If Chr$(NoteASCII) = "V" Then
        ' Set Volume
        If i < Len(NoteCaps) Then
          i = i + 1
          Select Case Mid$(NoteCaps, i, 1)
          Case "F"  ' Forte
            CurrentVolume = 127
          Case "O"  ' Mezzo-Forte
            CurrentVolume = 96
          Case "I"  ' Mezzo-Piano
            CurrentVolume = 65
          Case "P"  ' Piano
            CurrentVolume = 34
          Case Else
            i = i - 1
          End Select
        End If
      End If
      If Chr$(NoteASCII) = "M" Then
        ' Set Music Style
        If i < Len(NoteCaps) Then
          i = i + 1
          Select Case Mid$(NoteCaps, i, 1)
          Case "S"  ' Staccato
            MusicStyle = Style_Staccato
          Case "N"  ' Normal
            MusicStyle = Style_Normal
          Case "L"  ' Legato
            MusicStyle = Style_Legato
          Case "D"  ' Sustained
            MusicStyle = Style_Sustained
          Case Else
            i = i - 1
          End Select
        End If
      End If
    End If
  Loop
  ' Play Notes
  For i = 0 To CurrentNote
    ' Send Note
    If PlayNote(i).Pitch <> -1 Then SendMidiOut 144, _
PlayNote(i).Pitch, PlayNote(i).Volume
    ' Wait until next note should be played
    If i < CurrentNote Then
      PlayLength = ((((60 / Tempo) * 4) * (1 / _
PlayNote(i).Length)) * 1000)
      If PlayNote(i).Length > 0 Then
        Select Case PlayNote(i).Style
        Case Style_Sustained
          ' Play the full note value and don't stop it
		  ' afterwards
          SleepAPI Int(PlayLength + 0.5)
        Case Style_Normal
          ' Play 7/8 of the note value
          SleepAPI Int(PlayLength * (7 / 8) + 0.5)
          Call midiOutReset(MIDIDevice)
          SleepAPI Int((PlayLength * (1 / 8)) + 0.5)
        Case Style_Legato
          ' Play the full note value
          SleepAPI Int(PlayLength + 0.5)
          Call midiOutReset(MIDIDevice)
          SleepAPI 1
        Case Style_Staccato
          ' Play half the note value and pause for
		  ' the remainder
          SleepAPI Int(PlayLength * (1 / 2) + 0.5)
          Call midiOutReset(MIDIDevice)
          SleepAPI Int((PlayLength * (1 / 2)) + 0.5)
        End Select
      End If
    End If
    DoEvents
  Next i
  SleepAPI 1   ' This must be done in order for the last
		  ' note to be played
  ' Disable MIDI
  Call EnablePlay(PlayState_Disable)
End Sub
Private Function EnablePlay(Enable As Integer) As Boolean
  ' Enables/Disables MIDI Playing
  ' Enable = PlayState_?
  Dim MIDIOut As Long, ReturnValue As Long
  Static MIDIEnabled As Boolean
  If (Enable <> PlayState_Disable) And MIDIEnabled = False Then
    ' Enable MIDI
    ReturnValue = midiOutOpen(MIDIOut, -1, 0&, 0&, 0&)
    If ReturnValue = 0 Then
      MIDIEnabled = True
      EnablePlay = True
      MIDIDevice = MIDIOut
    Else
      EnablePlay = False
    End If
  ElseIf (Enable <> PlayState_Enable) And MIDIEnabled = True Then
    ' Disable MIDI
    ReturnValue = midiOutClose(MIDIDevice)
    If ReturnValue = 0 Then
      MIDIEnabled = False
      EnablePlay = True
    Else
      EnablePlay = False
    End If
  End If
End Function
Private Sub SendMidiOut(MidiEventOut As Long, MidiNoteOut As Long,_
MidiVelOut As Long)
  ' Sends the Note to the MIDI Device
  Dim LowInt As Long, VelOut As Long, HighInt As Long,_
MIDIMessage As Long
  Dim ReturnValue As Long
  LowInt = (MidiNoteOut * 256) + MidiEventOut
  VelOut = MidiVelOut * 256
  HighInt = VelOut * 256
  MIDIMessage = LowInt + HighInt
  ReturnValue = midiOutShortMsg(MIDIDevice, MIDIMessage)
End Sub
```

