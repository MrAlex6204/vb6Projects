VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cMMKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------------------
' ENUM
'---------------------------------------------------------------------------------------
Public Enum emmKey
    mmKey_LAUNCH_MEDIA_SELECT = &HB5
    mmKey_MEDIA_NEXT_TRACK = &HB0
    mmKey_MEDIA_PLAY_PAUSE = &HB3
    mmKey_MEDIA_PREV_TRACK = &HB1
    mmKey_MEDIA_STOP = &HB2
    mmKey_VOLUME_UP = 175
    mmKey_VOLUME_MUTE = 173
    mmKey_VOLUME_DOWN = 174
End Enum

'=======================================================================================
' IMPLEMENTS
'=======================================================================================
Implements iHook

'=======================================================================================
' MEMBER VARIABLES
'=======================================================================================
Private c_cHook As cHook

'---------------------------------------------------------------------------------------
' EVENTS
'---------------------------------------------------------------------------------------
Public Event KeyEventDown(mmKey As emmKey)
Public Event KeyEventUp(mmKey As emmKey)

Private Sub Class_Initialize()
    Set c_cHook = New cHook                                     'Create the hook class
    Call c_cHook.Hook(Me, WH_KEYBOARD_LL, False)
End Sub

Private Sub Class_Terminate()
    Set c_cHook = Nothing
End Sub

Private Sub iHook_Proc(ByVal bBefore As Boolean, bHandled As Boolean, lReturn As Long, nCode As WinSubHook2.eHookCode, wParam As Long, lParam As Long)
    Dim kb   As WinSubHook2.tKBDLLHOOKSTRUCT
    Dim bDown As Boolean
    
    Debug.Print "ok"
    
    If wParam = WinSubHook2.eMsg.WM_KEYDOWN Then
        bDown = True
    Else
        bDown = False
    End If
        
    kb = c_cHook.xKBDLLHOOKSTRUCT(lParam)
    Select Case kb.vkCode
        Case mmKey_LAUNCH_MEDIA_SELECT, mmKey_MEDIA_NEXT_TRACK, mmKey_MEDIA_PLAY_PAUSE _
           , mmKey_MEDIA_PREV_TRACK, mmKey_MEDIA_STOP, mmKey_VOLUME_UP, mmKey_VOLUME_MUTE _
           , mmKey_VOLUME_DOWN
            If bDown Then
                RaiseEvent KeyEventDown(kb.vkCode)
            Else
                RaiseEvent KeyEventUp(kb.vkCode)
            End If
    End Select
        
End Sub
