Attribute VB_Name = "Graficar_Audio"
Option Explicit

Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Private DevHandle As Long 'Handle des Audiodevice
                           
Private Declare Function waveInGetDevCaps Lib "winmm" Alias _
        "waveInGetDevCapsA" (ByVal uDeviceID As Long, ByVal _
        WaveInCapsPointer As Long, ByVal WaveInCapsStructSize _
        As Long) As Long

Private Declare Function waveInOpen Lib "winmm" _
        (WaveDeviceInputHandle As Long, ByVal WhichDevice As _
        Long, ByVal WaveFormatExPointer As Long, ByVal _
        CallBack As Long, ByVal CallBackInstance As Long, ByVal _
        flags As Long) As Long
                                                 
Private Declare Function waveInGetNumDevs Lib "winmm" () As Long

Private Declare Function waveInClose Lib "winmm" (ByVal _
        WaveDeviceInputHandle As Long) As Long
                
Private Declare Function waveInStart Lib "winmm" (ByVal _
        WaveDeviceInputHandle As Long) As Long
                
Private Declare Function waveInReset Lib "winmm" (ByVal _
        WaveDeviceInputHandle As Long) As Long
        
Private Declare Function waveInStop Lib "winmm" (ByVal _
        WaveDeviceInputHandle As Long) As Long
                
Private Declare Function sndplaysound Lib "winmm.dll" Alias _
        "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal _
        uFlags As Long) As Long
        
Private Declare Function waveInAddBuffer Lib "winmm" (ByVal _
        InputDeviceHandle As Long, ByVal WaveHdrPointer As _
        Long, ByVal WaveHdrStructSize As Long) As Long
                                                      
Private Declare Function waveInPrepareHeader Lib "winmm" _
        (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer _
        As Long, ByVal WaveHdrStructSize As Long) As Long
                                                          
Private Declare Function waveInUnprepareHeader Lib "winmm" _
        (ByVal InputDeviceHandle As Long, ByVal WaveHdrPointer _
        As Long, ByVal WaveHdrStructSize As Long) As Long
                
Private Type WAVEFORMATEX
    FormatTag As Integer
    Channels As Integer
    SamplesPerSec As Long
    AvgBytesPerSec As Long
    BlockAlign As Integer
    BitsPerSample As Integer
    ExtraDataSize As Integer
End Type

Private Type WaveHdr
    lpData As Long
    dwBufferLength As Long
    dwBytesRecorded As Long
    dwUser As Long
    dwFlags As Long
    dwLoops As Long
    lpNext As Long
    Reserved As Long
End Type

Private Type WaveInCaps
    ManufacturerID As Integer
    ProductID As Integer
    DriverVersion As Long
    ProductName(1 To 32) As Byte
    Formats As Long
    Channels As Integer
    Reserved As Integer
End Type

Const WAVE_INVALIDFORMAT As Long = &H0& 'invalid forma
Const WAVE_FORMAT_1M08 As Long = &H1&   '11.025 kHz,Mono,   8-bit
Const WAVE_FORMAT_1S08 As Long = &H2&   '11.025 kHz,Stereo, 8-bit
Const WAVE_FORMAT_1M16 As Long = &H4&   '11.025 kHz,Mono,  16-bit
Const WAVE_FORMAT_1S16 As Long = &H8&   '11.025 kHz,Stereo,16-bit
Const WAVE_FORMAT_2M08 As Long = &H10&  '22.05  kHz,Mono,   8-bit
Const WAVE_FORMAT_2S08 As Long = &H20&  '22.05  kHz,Stereo, 8-bit
Const WAVE_FORMAT_2M16 As Long = &H40&  '22.05  kHz,Mono,  16-bit
Const WAVE_FORMAT_2S16 As Long = &H80&  '22.05  kHz,Stereo,16-bit
Const WAVE_FORMAT_4M08 As Long = &H100& '44.1   kHz,Mono,   8-bit
Const WAVE_FORMAT_4S08 As Long = &H200& '44.1   kHz,Stereo, 8-bit
Const WAVE_FORMAT_4M16 As Long = &H400& '44.1   kHz,Mono,  16-bit
Const WAVE_FORMAT_4S16 As Long = &H800& '44.1   kHz,Stereo,16-bit

Const WAVE_FORMAT_PCM As Long = 1&

'Statuskonstanten
Const WHDR_DONE As Long = &H1&
Const WHDR_PREPARED As Long = &H2&
Const WHDR_BEGINLOOP As Long = &H4&
Const WHDR_ENDLOOP As Long = &H8&
Const WHDR_INQUEUE As Long = &H10&

Dim WaveHead As WaveHdr
Dim WavData() As Integer
Dim Ns As Long
Dim WaveFmt As WAVEFORMATEX
Dim WasActive As Long

Public Sub OpenDevice()
Ns = 1024
ReDim WavData(0 To Ns - 1)
WaveFmt.FormatTag = WAVE_FORMAT_PCM ' Stereo geht auch, aber dann das doppelte Ns vorsehen!
WaveFmt.Channels = 1
WaveFmt.SamplesPerSec = 10000 '11khz, altenativ 22050, 44100
WaveFmt.BitsPerSample = 16
WaveFmt.BlockAlign = (WaveFmt.Channels * WaveFmt.BitsPerSample) \ 8
WaveFmt.AvgBytesPerSec = WaveFmt.BlockAlign * WaveFmt.SamplesPerSec
WaveFmt.ExtraDataSize = 0
Call waveInOpen(DevHandle, 0, VarPtr(WaveFmt), 0, 0, 0)
If DevHandle = 0 Then MsgBox "NO se pudo iniciar la Grafica!", vbExclamation: Exit Sub
Call waveInStart(DevHandle)
WaveHead.lpData = VarPtr(WavData(0))
WaveHead.dwBufferLength = Ns
WaveHead.dwFlags = 0
Call waveInPrepareHeader(DevHandle, VarPtr(WaveHead), Len(WaveHead))
End Sub

Public Sub CloseDevice()
Call waveInUnprepareHeader(DevHandle, VarPtr(WaveHead), Len(WaveHead))
Call waveInReset(DevHandle)
Call waveInClose(DevHandle)
DevHandle = 0
End Sub


Public Sub GraficarAudio()
On Error Resume Next

Dim Amax As Long, i As Integer
Call waveInAddBuffer(DevHandle, VarPtr(WaveHead), Len(WaveHead))

For i = 0 To UBound(WavData)

    If WavData(i) > Amax Then Amax = WavData(i) 'buscar el maximo valor
Next i
Amax = 100 * Amax / 32768 'pasar a 100%


If Amax > 5 Then
    For i = 0 To 2

        Form1.LabelProgres(i).Caption = String(Amax / 3, "I")
    Next
Else
    For i = 0 To 2

        Form1.LabelProgres(i).Caption = ""
    Next
End If
End Sub


'Envia los comandos al dispositivo MCI:
Private Sub SendCommand(Command As String, Optional ReturnString As String, Optional ReturnLength As Long)
If ReturnString = vbNullString Then ReturnString = 0
Call mciSendString(Command, ReturnString, ReturnLength, 0)
End Sub

'Empieza a grabar la entrada del Microfono:
Public Sub RECORD_Start()
On Local Error Resume Next

Call SendCommand("open new type waveaudio alias WavFile")
Call SendCommand("record WavFile insert")

End Sub

'Termina la grabacion de la entrada del Microfono:
Public Sub RECORD_Finish()
On Local Error Resume Next

Call SendCommand("stop WavFile wait")

End Sub

'Guarda en un fichero WAV la grabacion:
Public Sub RECORD_Save()
On Local Error Resume Next

Call SendCommand("save WavFile C:\TempWave.wav")
Call SendCommand("close WavFile")

End Sub



