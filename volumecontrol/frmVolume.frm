VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmVolume 
   Caption         =   "Form1"
   ClientHeight    =   1575
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3525
   LinkTopic       =   "Form1"
   ScaleHeight     =   1575
   ScaleWidth      =   3525
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   1320
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Volume Up"
      Height          =   735
      Left            =   2040
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Volume Down"
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   10
   End
End
Attribute VB_Name = "frmVolume"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim q As String, ahgt, ghtr As String
Dim id As Long, v As Long, i As Long, lVol As lVolType, Vol As VolType, lv As Double, rv As Double


Private Declare Function auxGetVolume Lib "WINMM.DLL" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long


Private Declare Function mciGetDeviceID Lib "WINMM.DLL" Alias "mciGetDeviceIDA" (ByVal lpstrName As String) As Long


Private Declare Function waveOutGetVolume Lib "WINMM.DLL" (ByVal uDeviceID As Long, lpdwVolume As Long) As Long


Private Declare Function waveOutSetVolume Lib "WINMM.DLL" (ByVal uDeviceID As Long, ByVal dwVolume As Long) As Long


Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long


Private Type lVolType
    v As Long
    End Type


Private Type VolType
    lv As Integer
    rv As Integer
    End Type
Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If q = "1" Then
        Exit Sub
    End If
    id = -0
    i = waveOutGetVolume(id, v)
    lVol.v = v
    LSet Vol = lVol
    lv = Vol.lv: rv = Vol.rv
    lv = lv - &HFFF
    rv = rv - &HFFF
    If lv < -32768 Then lv = 65535 + lv
    If rv < -32768 Then rv = 65535 + rv
    Vol.lv = lv
    Vol.rv = rv
    LSet lVol = Vol
    v = lVol.v
    i = waveOutSetVolume(id, v)
    Call Findout
    StatusBar1.SimpleText = ProgressBar1.Value
End Sub

Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If q = "10" Then
        Exit Sub
    End If
    Dim dfre
    j = 1
    id = -0
    i = waveOutGetVolume(id, v)
    lVol.v = v
    LSet Vol = lVol
    lv = Vol.lv: rv = Vol.rv
    lv = lv + &HFFF
    rv = rv + &HFFF
    'If lv <= -30000 Then Exit Sub
    If lv > 32767 Then lv = lv - 65536
    If rv > 32767 Then rv = rv - 65536
    Vol.lv = lv
    Vol.rv = rv
    LSet lVol = Vol
    v = lVol.v
    i = waveOutSetVolume(id, v)
    Call Findout
    StatusBar1.SimpleText = ProgressBar1.Value
End Sub

Private Sub Form_Load()
    Call Findout
End Sub
Sub Findout()
    id = -0
    i = waveOutGetVolume(id, v)
    lVol.v = v
    LSet Vol = lVol
    lv = Vol.lv: rv = Vol.rv
    lv = lv - &HFFF
    rv = rv - &HFFF
    If lv < -32768 Then lv = 65535 + lv
    If rv < -32768 Then rv = 65535 + rv
    Vol.lv = lv
    Vol.rv = rv
    LSet lVol = Vol
    v = lVol.v
    ghtr = Left(lv, 1)


    If ghtr = "-" Then
        GoTo erre
    End If


    If lv < 5000 Then
        q = 1
        GoTo sayit
    End If


    If lv < 10000 Then
        q = 2
        GoTo sayit
    End If


    If lv < 15000 Then
        q = 3
        GoTo sayit
    End If


    If lv < 20000 Then
        q = 4
        GoTo sayit
    End If


    If lv < 25000 Then
        q = 5
        GoTo sayit
    End If


    If lv < 30000 Then
        q = 6
        GoTo sayit
    End If
erre:


    If lv < (-28000) Then
        q = 7
        GoTo sayit
    End If


    If lv < (-22000) Then
        q = 8
        GoTo sayit
    End If


    If lv < (-15000) Then
        q = 9
        GoTo sayit
    End If


    If lv < (-8000) Then
        q = 10
        GoTo sayit
    End If
sayit:
    ProgressBar1.Value = q
End Sub
