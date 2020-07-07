Option Explicit
Option Base 1

Private Const INT_ACOUSTIC_GRAND        As Integer = 0
Private Const INT_DEFAULT_NOTE_LENGTH   As Integer = 500
Private Const INT_MAX                   As Integer = 127
Private Const INT_MIN                   As Integer = 0
Private Const INT_DEFAULT_OCTAVE_RANGE  As Integer = 5
Private Const INT_MIDDLE_OCTAVE_RANGE   As Integer = 5
Private Const STR_DEFAULT_NOTE_NAME     As String = "C"

Private mintInstrument                  As Integer
Private mintVolume                      As Integer
Private mintLength                      As Integer
Private mintNote                        As Integer
Private mintOctaveNo                    As Integer
Private mstrNoteName                    As String

Private Enum ENU_DIRECTION
    Higher = 1
    Lower = 2
End Enum

Private Sub Class_Initialize()
    mintInstrument = INT_ACOUSTIC_GRAND
    mintVolume = INT_MAX
    mintLength = INT_DEFAULT_NOTE_LENGTH
    mstrNoteName = STR_DEFAULT_NOTE_NAME
    mintOctaveNo = INT_DEFAULT_OCTAVE_RANGE
    mintNote = GetNoteNumber()
End Sub

Public Property Get Up() As Integer
    Up = ENU_DIRECTION.Higher
End Property

Public Property Get Down() As Integer
    Down = ENU_DIRECTION.Lower
End Property

Public Property Get Instrument()
     Instrument = mintInstrument
End Property

Public Property Let Volume(ByVal Value)
    If (Value > INT_MIN) And (Value < INT_MAX) Then
        mintVolume = Value
    End If
End Property

Public Property Get Volume()
    Volume = mintVolume
End Property
   
Public Property Let NoteLength(ByVal Value)
    mintLength = Value
End Property

Public Property Get NoteLength()
    NoteLength = mintLength
End Property

Public Property Let NoteName(ByVal Value)
    mstrNoteName = Value
    mintNote = GetNoteNumber()
End Property

Public Property Get NoteName()
    NoteName = mstrNoteName
End Property

Public Property Get NoteNumber()
    NoteNumber = mintNote
End Property

Public Property Get MiddleOctave()
    MiddleOctave = INT_MIDDLE_OCTAVE_RANGE
End Property

Public Property Let OctaveNo(ByVal Value)
    mintOctaveNo = Value
End Property

Public Sub MoveSemiTone(ByVal intDirection As Integer, Optional ByVal iNum As Integer = 1)
    MoveNote ((1 * iNum) * IIf(intDirection = ENU_DIRECTION.Lower, -1, 1))
End Sub

Public Sub MoveWholeTone(ByVal intDirection As Integer, Optional ByVal iNum As Integer = 1)
    MoveNote ((2 * iNum) * IIf(intDirection = ENU_DIRECTION.Lower, -1, 1))
End Sub

Private Sub MoveNote(ByVal intMargin As Integer)
    mintNote = mintNote + intMargin
End Sub

Private Function GetNoteNumber()
    Dim iBase           As Integer
    Dim arrBaseNoteC    As Variant
    arrBaseNoteC = Array(12, 24, 36, 48, 60, 72, 84, 96, 108)
    iBase = arrBaseNoteC(mintOctaveNo)
    Select Case mstrNoteName
    Case "C"
        iBase = iBase + 0
    Case "D"
        iBase = iBase + 2
    Case "E"
         iBase = iBase + 4
    Case "F"
         iBase = iBase + 5
    Case "G"
         iBase = iBase + 7
    Case "A"
         iBase = iBase + 9
    Case "B"
         iBase = iBase + 11
    End Select
    GetNoteNumber = iBase
End Function