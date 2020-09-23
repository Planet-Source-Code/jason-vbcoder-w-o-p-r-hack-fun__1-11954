VERSION 5.00
Object = "{27395F88-0C0C-101B-A3C9-08002B2F49FB}#1.1#0"; "PICCLP32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Running Program...."
   ClientHeight    =   630
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4320
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   630
   ScaleWidth      =   4320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrFlash 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   2400
      Top             =   240
   End
   Begin VB.Timer TmrLock 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2040
      Top             =   240
   End
   Begin VB.Timer TmrHack 
      Interval        =   100
      Left            =   1680
      Top             =   240
   End
   Begin PicClip.PictureClip LEDRed 
      Left            =   960
      Top             =   2640
      _ExtentX        =   5239
      _ExtentY        =   714
      _Version        =   393216
      Cols            =   11
      Picture         =   "frmMain.frx":0442
   End
   Begin VB.Label DragME 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
   Begin VB.Image LEDDisplay 
      Height          =   495
      Index           =   0
      Left            =   120
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'Program variables
Dim LEDDigits As Integer
Dim DigLock(255) As Integer 'Tracks the "Locked in" digets
Dim FlashCount As Integer
Dim WOPRWav As Variant
Private Sub Setup_Hack_Gen()
'This will setup the picture clip control and size the form
Dim i As Integer
'setup picture clip for Num of Digits requested
LEDDisplay(0) = LEDRed.GraphicCell(0)
For i = 1 To LEDDigits
    Load LEDDisplay(i)
    LEDDisplay(i) = LEDRed.GraphicCell(0)
    LEDDisplay(i).Left = LEDDisplay(i - 1).Left + LEDDisplay(i - 1).Width
    LEDDisplay(i).Visible = True
Next
'Size Form to wrap display digits
frmMain.Width = LEDDisplay(LEDDigits).Left + LEDDisplay(LEDDigits).Width + 120
'DrageMe is a transparent label used to drag the form around
DragME.Width = frmMain.Width
DragME.Height = frmMain.Height
End Sub

Private Sub Hack_Gen()
'Generate randum numbers for display
Dim i As Integer
Dim r As Integer
For i = 0 To LEDDigits
    DoEvents
    If DigLock(i) = 0 Then 'If Num not locked then change it
        r = Int((10 * Rnd) + 0)   ' Generate random value between 0 and 9.
        LEDDisplay(i) = LEDRed.GraphicCell(r)
    End If
Next
End Sub

Private Sub DragME_DblClick()
'Clean up API Wav player and edit
EndPlaySound
End
End Sub

Private Sub DragME_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'click label box to drag around
    ReleaseCapture
    SendMessage Me.hwnd, &HA1, 2, 0&
End Sub

Private Sub Form_Load()
Dim LEDinput As Integer
Dim Responce As Integer
'The error handler is for when nothing is entered in the "input box"
On Error GoTo SkipInput
LEDinput = InputBox("Input number between 1 - 255. Press Enter for default. (Default = 12)", "Launch Code Length")
SkipInput:
'Do you want sound?
Responce = MsgBox("Would you like to enable sound?", vbYesNo, "W.O.P.R. Sound")
If Responce = vbYes Then
    BeginPlaySound 102
End If
'setup LED Display number length
If LEDinput = 0 Then
    LEDDigits = 12
ElseIf LEDinput > 255 Then
    LEDDigits = 12
Else
    LEDDigits = LEDinput
End If
'Adjust for 0 (I can not remember why this is here)
LEDDigits = LEDDigits - 1
Setup_Hack_Gen
'Set window to allways on top
SetWindowPos frmMain.hwnd, -1, frmMain.Left, frmMain.Top, frmMain.Width / 15, frmMain.Height / 15, SWP_NOACTIVATE Or SWP_SHOWWINDOW
'Go
TmrLock.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Cleanup wav player
EndPlaySound
End Sub

Private Sub TmrFlash_Timer()
'Flash the number after getting all digits
If DragME.BackStyle = 0 Then
    DragME.BackStyle = 1
Else
    DragME.BackStyle = 0
End If
FlashCount = FlashCount + 1
'if flash count time is over then reset
If FlashCount > 150 Then 'AKA flash for 15 seconds (TmrFlash.interval = 100)
Dim i As Integer
    For i = 0 To LEDDigits
        DigLock(i) = 0
    Next
    DragME.BackStyle = 0
    TmrFlash.Enabled = False
    TmrHack.Enabled = True
    TmrLock.Enabled = True
    FlashCount = 0
End If
End Sub

Private Sub TmrHack_Timer()
'Call hack_Gen routine
Hack_Gen
End Sub

Private Sub TmrLock_Timer()
'This randomly picks one of the digits and sets the array lock.
'This "Locks in" the random number. No checking is done if it's
'allready locked. This adds a more random fell as to how long it takes
Dim x As Integer
'Dim i As Integer
'Generate number to lock
x = LEDDigits + 1 'Set upper limit on random number
x = Int((x * Rnd) + 0) 'Get random number
DigLock(x) = 1 'Set number lock
'Check array if all numbers are locked
'If all locked Flash number
For x = 0 To LEDDigits
    DoEvents
    If DigLock(x) = 1 Then
    'Loop throught the array and check if all digits locked
        If x = LEDDigits Then
        'All digits locked Flash the display
            TmrHack.Enabled = False
            TmrLock.Enabled = False
            TmrFlash.Enabled = True
        End If
    Else
        Exit For
    End If
Next
End Sub

