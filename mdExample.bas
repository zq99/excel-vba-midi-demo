Option Explicit

Public Sub PlayMajorScale()
    Dim oPiano    As New csMidi
    Dim oNote     As New csNote
    Dim intCount  As Integer

    oNote.NoteName = "C"
    oNote.OctaveNo = oNote.MiddleOctave
    For intCount = 1 To 8
        DoEvents
        oPiano.PlayNote oNote
        Select Case intCount
        Case 3, 7
            oNote.MoveSemiTone oNote.Up
        Case Else
            oNote.MoveWholeTone oNote.Up
        End Select
    Next
    Set oNote = Nothing
    Set oPiano = Nothing
End Sub