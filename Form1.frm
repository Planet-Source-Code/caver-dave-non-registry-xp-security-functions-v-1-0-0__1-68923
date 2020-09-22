VERSION 5.00
Begin VB.Form Form1 
   Caption         =   " 4 FUNCTIONS, 7 BUTTONS & 1 FORM"
   ClientHeight    =   4740
   ClientLeft      =   4290
   ClientTop       =   3105
   ClientWidth     =   8970
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   8970
   Begin VB.CommandButton Command1 
      Caption         =   "LC"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   375
      Left            =   2040
      TabIndex        =   4
      Top             =   3840
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Command4"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Command5"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   3840
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Command6"
      Height          =   375
      Left            =   3960
      TabIndex        =   1
      Top             =   4200
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8400
      TabIndex        =   0
      Top             =   4080
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "TRY CTRL + ALT + DEL WHEN RUNNING APP THEN CLOSE APP WITH THE  X BUTTON AND WHEN THE APP HAS CLOSED TRY IT AGAIN "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   720
      TabIndex        =   8
      Top             =   2640
      Width           =   7335
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   720
      TabIndex        =   7
      Top             =   240
      Width           =   7095
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' types for cursor limiting
Private Type RECT
left As Integer
top As Integer
right As Integer
bottom As Integer
End Type
Private Type POINT
X As Long
Y As Long
End Type
' api for cursor limiting
Private Declare Sub ClipCursor Lib "user32" (lpRect As Any)
Private Declare Sub GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT)
Private Declare Sub ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As POINT)
Private Declare Sub OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long)
' api for showing and hiding taskbar and desktop
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
' show and hide window constants
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40
Private Sub Command3_Click()
'************************************************************************
'*         to show taskbar and retain cursor restraint to form          *
'************************************************************************
Dim Thwnd As Long
Thwnd = FindWindow("Shell_traywnd", "")
Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)

Call Command1_Click ' cursor restraint
End Sub

Private Sub Command4_Click()
'************************************************************************
'*         to hide taskbar and retain cursor restraint to form          *
'************************************************************************
Dim Thwnd As Long
Thwnd = FindWindow("Shell_traywnd", "")
Call SetWindowPos(Thwnd, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)

Call Command1_Click ' cursor restraint
End Sub

Private Sub Command5_Click()
'************************************************************************
'*       to show the desktop and retain cursor restraint to form        *
'************************************************************************
Dim hWnd As Long
hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hWnd, 1

Call Command1_Click ' cursor restraint

End Sub

Private Sub Command6_Click()
'************************************************************************
'*       to hide the desktop and retain cursor restraint to form        *
'************************************************************************
Dim hWnd As Long
hWnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hWnd, 0

Call Command1_Click ' cursor restraint

End Sub

Private Sub Command7_Click()
Unload Me
End Sub

Private Sub Form_Load()
'************************************************************************
'*           centre the form, open and lock the task manager            *                                                              *
'************************************************************************
Dim TopCorner As Integer
  Dim LeftCorner As Integer
  'centres the form on the screen
  If Me.WindowState <> 0 Then Exit Sub

  TopCorner = (Screen.Height - Me.Height) \ 2
  LeftCorner = (Screen.Width - Me.Width) \ 2
  Me.Move LeftCorner, TopCorner
  
  
' effectively disables ctrl+alt+del
Open "c:\windows\system32\taskmgr.exe" For Random Lock Read As #1
'****** MUST be released on Unload ***************
 
Command1.Caption = "LOCK CURSOR"
Command2.Caption = "REALEASE CURSOR"
Command3.Caption = "SHOW TASKBAR"
Command4.Caption = "HIDE TASKBAR"
Command5.Caption = "SHOW DESKTOP"
Command6.Caption = "HIDE DESKTOP"
End Sub
Private Sub Command1_Click()
'************************************************************************
'*            limits the cursor movement to within the form             *
'************************************************************************
Dim client As RECT
Dim upperleft As POINT

GetClientRect Me.hWnd, client
upperleft.X = client.left
upperleft.Y = client.top
ClientToScreen Me.hWnd, upperleft
OffsetRect client, upperleft.X, upperleft.Y
ClipCursor client
End Sub
Private Sub Command2_Click()
'************************************************************************
'*                     Releases the cursor limits                       *
'************************************************************************
ClipCursor ByVal 0&
End Sub
Private Sub Form_Unload(Cancel As Integer)
'************************************************************************
'*        Releases the cursor limits and unlocks the task manager       *
'************************************************************************
ClipCursor ByVal 0&
Close #1 ' release task manager enable ctrl+alt+del
End Sub
