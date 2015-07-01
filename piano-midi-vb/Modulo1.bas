Attribute VB_Name = "Modulo1"
Option Explicit

Public Const MAXPNAMELEN = 32


Public Const MMSYSERR_BASE = 0
Public Const MMSYSERR_BADDEVICEID = (MMSYSERR_BASE + 2)     ' El ID del dispositivo fuera de rango
Public Const MMSYSERR_INVALPARAM = (MMSYSERR_BASE + 11)     ' Parámetro invalido pasado
Public Const MMSYSERR_NODRIVER = (MMSYSERR_BASE + 6)        ' No hay dispositivo detectado
Public Const MMSYSERR_NOMEM = (MMSYSERR_BASE + 7)           ' Error Alojamiento de memoria

Public Const MMSYSERR_INVALHANDLE = (MMSYSERR_BASE + 5)     ' handle del dispositivo es inválido
Public Const MIDIERR_BASE = 64
Public Const MIDIERR_STILLPLAYING = (MIDIERR_BASE + 1)
Public Const MIDIERR_NOTREADY = (MIDIERR_BASE + 3)
Public Const MIDIERR_BADOPENMODE = (MIDIERR_BASE + 6)       ' operación  o soportado en / modo de abrir


Type MIDIOUTCAPS
   wMid As Integer
                                     
                                     
   
   wPid As Integer
                                     
                                     
   
   vDriverVersion As Long
                                     
                                     ' version menor, mayor del midi
                                     
   szPname As String * MAXPNAMELEN   ' Nombre Producto
   
   wTechnology As Integer
                                     '     MOD_FMSYNTH-el didpositivo es FM synthesizer.
                                     '     MOD_MAPPER-el didpositivo es Microsoft MIDI mapper.
                                     '     MOD_MIDIPORT-el didpositivo es MIDI hardware port.
                                     '     MOD_SQSYNTH-el didpositivo es square wave synthesizer.
                                     '     MOD_SYNTH-el didpositivo es a .
                                     
   wVoices As Integer
   wNotes As Integer
                                     
   wChannelMask As Integer
End Type

Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long

Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long

Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long

Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long

