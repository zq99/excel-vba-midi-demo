
# Excel VBA MIDI Music Project

This project demonstrates how to play musical scales using VBA.

## Setup

Add the MIDI and note class to a VBA project. The MIDI class encapsulates all the low level Windows MIDI calls.

The Note class handles the logic for playing notes. To make use of the MIDI functionality, new code should interact with the Note class only (or whatever its
equivalent). The idea is to keep the MIDI class separate from any application logic.


## Example

This is how to play a note using the class:

    Dim oPiano As New csMidi
    Dim oNote as New csNote
    
    oNote.NoteName = "C"
    oNote.OctaveNo = oNote.MiddleOctave
    oPiano.PlayNote oNote


## Further Info

- This project : https://datapluscode.com/general/using-excel-vba-midi-play-scales/
- Windows MIDI : https://docs.microsoft.com/en-us/windows/win32/multimedia/midi-functions?redirectedfrom=MSDN
