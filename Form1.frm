VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1215
   ClientLeft      =   14190
   ClientTop       =   9585
   ClientWidth     =   780
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1215
   ScaleWidth      =   780
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox taban 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   3
      Left            =   480
      Picture         =   "Form1.frx":57E2
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   9
      Top             =   900
      Width           =   240
   End
   Begin VB.PictureBox taban 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   480
      Picture         =   "Form1.frx":AE3D
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   630
      Width           =   240
   End
   Begin VB.PictureBox taban 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   480
      Picture         =   "Form1.frx":10498
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   360
      Width           =   240
   End
   Begin VB.PictureBox taban 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   480
      Picture         =   "Form1.frx":15AF3
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   6
      Top             =   90
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   1
      Left            =   360
      Picture         =   "Form1.frx":1B14E
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   1860
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   360
      Picture         =   "Form1.frx":207A9
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   240
   End
   Begin VB.Timer tmrKeyStates 
      Interval        =   100
      Left            =   4080
      Top             =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ins"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   3
      Left            =   90
      TabIndex        =   3
      Top             =   900
      Width           =   210
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Scr"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   2
      Top             =   660
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caps"
      ForeColor       =   &H00808080&
      Height          =   195
      Index           =   1
      Left            =   60
      TabIndex        =   1
      Top             =   390
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Num"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   330
   End
   Begin VB.Menu mnuForm 
      Caption         =   "Form"
      Visible         =   0   'False
      Begin VB.Menu mnuShow 
         Caption         =   "&Show"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuHide 
         Caption         =   "&Hide"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuauto 
         Caption         =   "&Bir dahaki sefere otomatik basla"
      End
      Begin VB.Menu mnuend 
         Caption         =   "&End"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private ontop As New clsOnTop
Option Explicit

Public Function CenterForm32(frm As Form)
Dim ScreenWidth&, ScreenHeight&, ScreenLeft&, ScreenTop&
Dim DeskTopArea As RECT
Call SystemParametersInfo(SPI_GETWORKAREA, 0, DeskTopArea, 0)
ScreenHeight = (DeskTopArea.Bottom - DeskTopArea.Top) * Screen.TwipsPerPixelY
ScreenWidth = (DeskTopArea.Right - DeskTopArea.Left) * Screen.TwipsPerPixelX
ScreenLeft = DeskTopArea.Left * Screen.TwipsPerPixelX
ScreenTop = DeskTopArea.Top * Screen.TwipsPerPixelY
'eðer ortalanmak isterse
'frm.Move (ScreenWidth - frm.Width) \ 2 + ScreenLeft, (ScreenHeight - frm.Height) \ 2 + ScreenTop
frm.Move (ScreenWidth - frm.Width) + ScreenLeft, (ScreenHeight - frm.Height) + ScreenTop
End Function
Public Function CapsLockOn() As Boolean
  Dim iKeyState As Integer
  iKeyState = GetKeyState(vbKeyCapital)
  CapsLockOn = (iKeyState = 1 Or iKeyState = -127)
End Function
Public Function NumLockOn() As Boolean
    Dim iKeyState As Integer
    iKeyState = GetKeyState(vbKeyNumlock)
    NumLockOn = (iKeyState = 1 Or iKeyState = -127)
End Function
Public Function ScrlLockOn() As Boolean
    Dim iKeyState As Integer
    iKeyState = GetKeyState(vbKeyScrollLock)
    ScrlLockOn = (iKeyState = 1 Or iKeyState = -127)
End Function
Public Function InsertOn() As Boolean
    Dim iKeyState As Integer
    iKeyState = GetKeyState(vbKeyInsert)
    InsertOn = (iKeyState = 1 Or iKeyState = -127)
End Function
Private Sub Form_Load()
On Error Resume Next
Dim Ret As Long
CenterForm32 Me
If GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Keystates") <> "" Then mnuauto.Checked = True Else mnuauto.Checked = False
ontop.MakeTopMost Me.hwnd
    'Set the window style to 'Layered'
    Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
    'Set the opacity of the layered window
    SetLayeredWindowAttributes Me.hwnd, 0, 170, LWA_ALPHA
    TrayAdd hwnd, Me.Icon, "System Tray", MouseMove
    'mnuHide_Click
End Sub
Private Sub Form_LostFocus()
ontop.MakeTopMost Me.hwnd
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
ReleaseCapture
SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
Private Sub Label1_Click(Index As Integer)
Select Case Index

Case 0
'num
            If CBool(GetKeyState(vbKeyNumlock) And 1) = False Then
               Call keybd_event(&H90, &H45, &H1 Or 0, 0)
               Call keybd_event(&H90, &H45, &H1 Or &H2, 0): Exit Sub
            Else
               Call keybd_event(&H90, &H45, &H1 Or 0, 0)
               Call keybd_event(&H90, &H45, &H1 Or &H2, 0)
            End If


Case 1
'caps
            If CBool(GetKeyState(vbKeyCapital) And 1) = False Then
               Call keybd_event(&H14, &H45, &H1 Or 0, 0)
               Call keybd_event(&H14, &H45, &H1 Or &H2, 0): Exit Sub
            Else
                Call keybd_event(&H14, &H45, &H1 Or 0, 0)
                Call keybd_event(&H14, &H45, &H1 Or &H2, 0)
            End If

Case 2
'scr
            If CBool(GetKeyState(vbKeyScrollLock) And 1) = False Then
               Call keybd_event(&H91, &H45, &H1 Or 0, 0)
               Call keybd_event(&H91, &H45, &H1 Or &H2, 0): Exit Sub
Else
               Call keybd_event(&H91, &H45, &H1 Or 0, 0)
               Call keybd_event(&H91, &H45, &H1 Or &H2, 0)
            End If

Case 3
'ins
            If CBool(GetKeyState(vbKeyInsert) And 1) = False Then
               Call keybd_event(&H2D, &H45, &H1 Or 0, 0)
               Call keybd_event(&H2D, &H45, &H1 Or &H2, 0): Exit Sub
Else
               Call keybd_event(&H2D, &H45, &H1 Or 0, 0)
               Call keybd_event(&H2D, &H45, &H1 Or &H2, 0)
            End If

End Select


End Sub

Private Sub mnuauto_Click()
If mnuauto.Checked = True Then
mnuauto.Checked = False: Exit Sub
Else
mnuauto.Checked = True
End If
End Sub

Private Sub mnuend_Click()
On Error Resume Next

If mnuauto.Checked = True And GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Keystates") = "" Then AddToRun "Keystates", "c:\Keystates.exe": GoTo ende
If mnuauto.Checked = False And GetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "Keystates") <> "" Then RemoveFromRun "Keystates": End

ende:
TrayDelete
End
End Sub

Private Sub taban_Click(Index As Integer)
Label1_Click Index
End Sub

Private Sub tmrKeyStates_Timer()
  
   If InsertOn() Then
      Label1(3).ForeColor = &HC0&
      taban(3).Picture = Picture1(0).Image
   Else
      Label1(3).ForeColor = &H808080
      taban(3).Picture = Picture1(1).Image
   End If
    
   If NumLockOn() Then
      Label1(0).ForeColor = &HC0&
      taban(0).Picture = Picture1(0).Image

   Else
      Label1(0).ForeColor = &H808080
      taban(0).Picture = Picture1(1).Image
   
   End If
  
   If CapsLockOn() Then
      Label1(1).ForeColor = &HC0&
      taban(1).Picture = Picture1(0).Image
   
   Else
      Label1(1).ForeColor = &H808080
      taban(1).Picture = Picture1(1).Image
   
   End If
  
   If ScrlLockOn() Then
      Label1(2).ForeColor = &HC0&
      taban(2).Picture = Picture1(0).Image
   
   Else
      Label1(2).ForeColor = &H808080
      taban(2).Picture = Picture1(1).Image
   
   End If
  
End Sub


Private Sub mnuHide_Click()
    If Not Me.WindowState = 1 Then WindowState = 1: Me.Hide
End Sub

Private Sub mnuShow_Click()
    If Me.WindowState = 1 Then WindowState = 0: Me.Show
    TrayDelete  '[Deleting Tray]
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim cEvent As Single
cEvent = x / Screen.TwipsPerPixelX
Select Case cEvent
    Case MouseMove
    Case LeftUp
    Case LeftDown
    Case LeftDbClick
    Case MiddleUp
    Case MiddleDown
    Case MiddleDbClick
    Case RightUp
        PopupMenu mnuForm
    Case RightDown
    Case RightDbClick
End Select
End Sub

