Attribute VB_Name = "madMain"
'API function for setting form allways on top
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    Global Const conHwndTopmost = -1
    Global Const conSwpNoActivate = &H10
    Global Const conSwpShowWindow = &H40

'API function for playing WAV file
Declare Function sndPlaySound Lib "WINMM.DLL" Alias "sndPlaySoundA" (lpszSoundName As Any, ByVal uFlags As Long) As Long
    Public Const SND_ASYNC = &H1
    Public Const SND_NODEFAULT = &H2
    Public Const SND_MEMORY = &H4
    Public Const SND_LOOP = &H8
    Public Const SND_NOSTOP = &H10

'API calls for mouse drag without title bar
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function ReleaseCapture Lib "user32" () As Long

Global SoundBuffer() As Byte
    
Sub BeginPlaySound(ByVal ResourceId As Integer)
    SoundBuffer = LoadResData(ResourceId, "SOUND")
    sndPlaySound SoundBuffer(0), SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY Or SND_LOOP
End Sub

Sub EndPlaySound()
    sndPlaySound ByVal vbNullString, 0&
End Sub

Sub main()
frmMain.Show
End Sub

