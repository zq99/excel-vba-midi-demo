Option Explicit

Private Const MAXPNAMELEN               As Integer = 32
Private Const MMSYSERR_BASE             As Integer = 0
Private Const MMSYSERR_BADDEVICEID      As Integer = (MMSYSERR_BASE + 2)
Private Const MMSYSERR_INVALPARAM       As Integer = (MMSYSERR_BASE + 11)
Private Const MMSYSERR_NODRIVER         As Integer = (MMSYSERR_BASE + 6)
Private Const MMSYSERR_NOMEM            As Integer = (MMSYSERR_BASE + 7)
Private Const MMSYSERR_INVALHANDLE      As Integer = (MMSYSERR_BASE + 5)
Private Const MIDIERR_BASE              As Integer = 64
Private Const MIDIERR_STILLPLAYING      As Integer = (MIDIERR_BASE + 1)
Private Const MIDIERR_NOTREADY          As Integer = (MIDIERR_BASE + 3)
Private Const MIDIERR_BADOPENMODE       As Integer = (MIDIERR_BASE + 6)

Private Type MIDIOUTCAPS
   wMid             As Integer
   wPid             As Integer
   wTechnology      As Integer
   wVoices          As Integer
   wNotes           As Integer
   wChannelMask     As Integer
   vDriverVersion   As Long
   dwSupport        As Long
   szPname          As String * MAXPNAMELEN
End Type

Private Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Private Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

#If Win64 Then
Private Declare PtrSafe Function midiOutClose Lib “winmm.dll” (ByVal hMidiOut As LongPtr) As Long
Private Declare PtrSafe Function midiOutOpen Lib “winmm.dll” (lphMidiOut As LongPtr, ByVal uDeviceID As LongPtr, ByVal dwCallback As LongPtr, ByVal dwInstance As LongPtr, ByVal dwflags As LongPtr) As Long
Private Declare PtrSafe Function midiOutShortMsg Lib “winmm.dll” (ByVal hMidiOut As LongPtr, ByVal dwMsg As LongPtr) As Long
Private Declare PtrSafe Function timeGetTime Lib “winmm.dll” () As Long
#Else
Private Declare Function midiOutClose Lib “winmm.dll” (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutOpen Lib “winmm.dll” (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function midiOutShortMsg Lib “winmm.dll” (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Private Declare Function timeGetTime Lib “winmm.dll” () As Long
#End If


#If Win64 Then
Private mlngCurDevice As LongPtr
Private mlngHmidi As LongPtr
Private mlngRc As LongPtr
Private mlngMidiMsg As LongPtr
#Else
Private mlngCurDevice As Long
Private mlngHmidi As Long
Private mlngRc As Long
Private mlngMidiMsg As Long
#End If


Private mlngMiPrivatesg    As Long
Private mintChannel        As Integer
Private mintVolume         As Integer
Private mintNoteLength     As Integer
Private mintMidiNote       As Integer
Private mintInstrument     As Integer
Private mstrDeviceName     As String
Private mblnIsDeviceOpen   As Boolean
Private mlngNumDevices     As Long

Private Const INT_DEFAULT_CHANNEL       As Integer = 0
Private Const INT_DEFAULT_VOLUME        As Integer = 127
Private Const INT_DEFAULT_NOTE_LENGTH   As Integer = 1000
Private Const INT_DEFAULT_CUR_DEVICE    As Integer = 0

Private Sub Class_Initialize()
    mintChannel = INT_DEFAULT_CHANNEL
    mlngCurDevice = INT_DEFAULT_CUR_DEVICE
    mintVolume = INT_DEFAULT_VOLUME
    mintNoteLength = INT_DEFAULT_NOTE_LENGTH
    mblnIsDeviceOpen = False
    Call OpenDevice
End Sub

Private Sub Class_Terminate()
    Call CloseDevice
End Sub

Private Sub OpenDevice()
On Error GoTo ERR_HANDLER:
    If Not mblnIsDeviceOpen Then
        mlngRc = midiOutClose(mlngHmidi)
        mlngRc = midiOutOpen(mlngHmidi, mlngCurDevice, 0, 0, 0)
        If (mlngRc <> 0) Then
            MsgBox "Couldn't open midi out, lngc = " & mlngRc
            mblnIsDeviceOpen = False
        End If
        mblnIsDeviceOpen = True
    End If
    Exit Sub
ERR_HANDLER:
    Debug.Print "Open DLL Error:=" & Err.LastDllError
    mblnIsDeviceOpen = False
End Sub

Private Sub CloseDevice()
    If mblnIsDeviceOpen Then
        mlngRc = midiOutClose(mlngHmidi)
        mblnIsDeviceOpen = False
    End If
End Sub

Private Sub StartNote()
    mlngMidiMsg = &H90 + (mintMidiNote * &H100) + (mintVolume * &H10000) + mintChannel
    midiOutShortMsg mlngHmidi, mlngMidiMsg
End Sub

Private Sub StopNote()
    mlngMidiMsg = &H80 + (mintMidiNote * &H100) + mintChannel
    midiOutShortMsg mlngHmidi, mlngMidiMsg
End Sub

Private Sub PauseNote()
    Sleep mintNoteLength
End Sub

Public Sub PlayNote(ByVal note As csNote)
    mintNoteLength = note.NoteLength
    mintVolume = note.Volume
    mintInstrument = note.Instrument
    mintMidiNote = note.NoteNumber
    Call StartNote
    Call PauseNote
	Call StopNote
End Sub

Private Sub UpdateInstrument()
    If mblnIsDeviceOpen = True Then
        mlngMidiMsg = (mintInstrument * 256) + &HC0 + mintChannel + (0 * 256) * 256
        midiOutShortMsg mlngHmidi, mlngMidiMsg
    End If
End Sub

Private Sub getNumberOfDevices()
    mlngNumDevices = (midiOutGetNumDevs() - 1)
End Sub

Private Sub CurrentDeviceName()
    Dim caps    As MIDIOUTCAPS
    midiOutGetDevCaps mlngCurDevice, caps, Len(caps)
    mstrDeviceName = caps.szPname
End Sub

Public Function GetMIDIDevices() As String()
    Dim strRet()    As String
    Dim lngLoop     As Long
    Dim udtCap      As MIDIOUTCAPS
    
    mlngNumDevices = (midiOutGetNumDevs() - 1)
    ReDim strRet(0) As String
    strRet(0) = " MIDI Mapper"
    
    For lngLoop = 0 To mlngNumDevices
        mlngRc = midiOutGetDevCaps(lngLoop, udtCap, Len(udtCap))
        ReDim Preserve strRet(lngLoop + 1) As String
        strRet(lngLoop + 1) = udtCap.szPname
    Next
    GetMIDIDevices = strRet()
End Function