Attribute VB_Name = "modMIDI"
'' P I A N O  by Armin Niki
'' Original code:
'' http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=64928&lngWId=1
'' Update - Apr 16 2006 by Paul Bahlawan
'' modified for min code Nov 2007 by Scott Smith
Option Explicit
Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long
Private hmidi      As Long
Private baseNote   As Long
Private channel    As Long
Private volume     As Long
Private lNote      As Long
Private Playin()   As String
Private playinc    As Long
Private timers     As Long
Private rec        As String
Private midimsg    As Long
Private notep      As Long



Public Sub PlayNote(mNote As Long)
    Dim midimsg    As Long
    'Play note
    midimsg = &H90 + ((baseNote + mNote) * &H100) + (volume * &H10000) + channel
    midiOutShortMsg hmidi, midimsg
    'record the key-down event

    timers = 0
    lNote = mNote
    'hi-light key being played
    'pKey(mNote - 1).BackColor = &H6060F0
End Sub
'Stop a note
Public Sub PlayNoteOff(mNote As Long)
    Dim midimsg    As Long
    midimsg = &H80 + ((baseNote + mNote) * &H100) + channel
    midiOutShortMsg hmidi, midimsg
    'record the key-up event
    timers = 0
    If mNote = lNote Then lNote = 0    'lNote = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    midiOutClose (hmidi)
End Sub

'Change the instrument
Private Sub InstChange(Instru)
    Dim midimsg    As Long
    midimsg = (Instru * 256) + &HC0 + channel
    'Text1.Text = midimsg
    midiOutShortMsg hmidi, midimsg
End Sub
'Private Sub sldVol_Change()
'    volume = sldVol.Value
'    Text3.Text = sldVol.Value
'End Sub


Public Sub OpenMIDI()
    Dim rc         As Long
    Dim curDevice  As Long
    midiOutClose (hmidi)
    rc = midiOutOpen(hmidi, curDevice, 0, 0, 0)
    If (rc <> 0) Then
        MsgBox "Couldn't open midi device - Error #" & rc
    End If
    baseNote = 23
    channel = 15
    volume = 127


End Sub
Public Sub CloseMIDI()
    midiOutClose (hmidi)

End Sub
