Attribute VB_Name = "Insertar_Objeto"
Option Explicit
   'Deklarace API funkci a procedur
   Private Declare Function OleUIInsertObject _
                   Lib "oledlg.dll" _
                   Alias "OleUIInsertObjectA" _
                   (inParam As Any) As Long
   Private Declare Function ProgIDFromCLSID _
                   Lib "ole32.dll" _
                   (clsid As Any, _
                   strAddess As Long) As Long
   Private Declare Sub CoTaskMemFree _
                   Lib "ole32.dll" _
                   (ByVal pvoid As Long)
   Private Declare Sub CopyMemory _
                   Lib "kernel32" _
                   Alias "RtlMoveMemory" _
                   (Destination As Any, _
                   Source As Any, _
                   ByVal Length As Long)
   Private Declare Function lstrlenW _
                   Lib "kernel32" _
                   (ByVal lpString As Long _
                   ) As Long
   'Konstannty
   Const IOF_SHOWHELP = &H1
   Const IOF_SELECTCREATENEW = &H2
   Const IOF_SELECTCREATEFROMFILE = &H4
   Const IOF_CHECKLINK = &H8
   Const IOF_CHECKDISPLAYASICON = &H10
   Const IOF_CREATENEWOBJECT = &H20
   Const IOF_CREATEFILEOBJECT = &H40
   Const IOF_CREATELINKOBJECT = &H80
   Const IOF_DISABLELINK = &H100
   Const IOF_VERIFYSERVERSEXIST = &H200
   Const IOF_DISABLEDISPLAYASICON = &H400
   Const IOF_HIDECHANGEICON = &H800
   Const IOF_SHOWINSERTCONTROL = &H1000
   Const IOF_SELECTCREATECONTROL = &H2000

   'Konstanty navratovych kodu
   Const OLEUI_FALSE = 0
   Const OLEUI_SUCCESS = 1
   Const OLEUI_OK = 1
   Const OLEUI_CANCEL = 2

   ' GUID, IID, CLSID, atd
   Private Type GUID
       Data1 As Long
       Data2 As Integer
       Data3 As Integer
       Data4(0 To 7) As Byte
   End Type

   'UDT pouzite v OleUIInsertObject.
   Private Type OleUIInsertObjectType
       cbStruct As Long
       dwFlags As Long
       hWndOwner As Long
       lpszCaption  As String
       lpfnHook As Long
       lCustData As Long
       hInstance  As Long
       lpszTemplate As String
       hResource As Long
       clsid As GUID
       lpszFile As String
       cchFile As Long
       cClsidExclude As Long
       lpClsidExclude As Long
       IID As GUID
       oleRender As Long
       lpFormatEtc As Long
       lpIOleClientSite As Long
       lpIStorage As Long
       ppvObj As Long
       sc As Long
       hMetaPict As Long
   End Type
   '--------camara
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function capCreateCaptureWindow Lib "avicap32.dll" Alias "capCreateCaptureWindowA" (ByVal lpszWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hwndParent As Long, ByVal nID As Long) As Long
Public mCapHwnd As Long

Public Const CONNECT As Long = 1034
Public Const DISCONNECT As Long = 1035
Public Const GET_FRAME As Long = 1084
Public Const COPY As Long = 1054

'------------Audio
Public Declare Function sndplaysound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_ASYNC = (1)
Public Const SND_NODEFAULT = (2)
 Public Sub InsertarObjeto()
     Dim UIInsertObj As OleUIInsertObjectType
     Dim retValue As Long
     Dim lpolestr As Long
     Dim strsize As Long
     Dim ProgId As String

     On Error GoTo err 'Pri chybe skoc na err

     'Priprav strukturu
     With UIInsertObj
       .cbStruct = LenB(UIInsertObj)
       .dwFlags = IOF_SELECTCREATENEW
       .hWndOwner = Form1.hWnd
       .lpszFile = String(256, " ")
       .cchFile = Len(.lpszFile)
     End With

     'Zobraz dialog box
     retValue = OleUIInsertObject(UIInsertObj)

     If (retValue = OLEUI_OK) Then
       If ((UIInsertObj.dwFlags And _
            IOF_SELECTCREATENEW) = _
            IOF_SELECTCREATENEW) Then
          retValue = ProgIDFromCLSID(UIInsertObj.clsid, _
                                     lpolestr)
          strsize = lstrlenW(lpolestr) + 1
          ProgId = String(strsize, 0)
          CopyMemory ByVal StrPtr(ProgId), _
                     ByVal lpolestr, strsize * 2
          CoTaskMemFree lpolestr
          Form1.RichTextBox2.OLEObjects.Add , , "", ProgId
       Else
          Form1.RichTextBox2.OLEObjects.Add , _
          , UIInsertObj.lpszFile
       End If
      End If
   Exit Sub

err:
       MsgBox "Doslo k chybe cislo: " & _
              err.Number & _
              "Popis chyby:" & _
              vbNewLine & err.Description
   End Sub



Sub Main()
Form2.Show
Form2.icon = Form1.icon


End Sub
