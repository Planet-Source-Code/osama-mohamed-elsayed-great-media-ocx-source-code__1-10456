VERSION 5.00
Begin VB.Form Authorfrm 
   BackColor       =   &H00000080&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2475
   ClientLeft      =   2325
   ClientTop       =   1605
   ClientWidth     =   4665
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   1
      Left            =   810
      TabIndex        =   5
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Vote:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Index           =   0
      Left            =   960
      TabIndex        =   4
      Top             =   1800
      Width           =   465
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   $"Authorfrm.frx":0000
      ForeColor       =   &H00FFFF80&
      Height          =   1170
      Left            =   1560
      TabIndex        =   3
      Top             =   600
      Width           =   2370
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Osama Mohamed El sayed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   240
      Left            =   1560
      TabIndex        =   2
      ToolTipText     =   "About Author"
      Top             =   240
      Width           =   2460
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   1080
      Picture         =   "Authorfrm.frx":00AE
      Top             =   2640
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "www.Planet-Source-Code.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1560
      TabIndex        =   1
      ToolTipText     =   "Panet-Source-Code.com"
      Top             =   1800
      Width           =   2715
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "osama_nt@intouch.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1920
      TabIndex        =   0
      ToolTipText     =   "Author E-mail"
      Top             =   2040
      Width           =   2160
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   120
      Picture         =   "Authorfrm.frx":0200
      ToolTipText     =   "About Author"
      Top             =   240
      Width           =   1305
   End
End
Attribute VB_Name = "Authorfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_DblClick()
Unload Me
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
ReleaseCapture
SendMessage hwnd, &HA1, 2, 0&
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Label1
  '.FontBold = False
  .FontUnderline = False
  .ForeColor = vbWhite
End With
With Label2
  .ForeColor = vbWhite
End With

   Me.MousePointer = 0
End Sub

Private Sub Image1_Click()
Dim Scr_hDC As Long
Dim ShellDoc As Long
Dim DocName As String
DocName = "http://planet-source-code.com/vb/authors/ShowAuthorBio.asp?lngAuthorId=37911&lngWId=1"
Scr_hDC = GetDesktopWindow()
ShellDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "C:\", 1)
End Sub

Private Sub Image1_DblClick()
Unload Me
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseDown(1, False, X, Y)
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Image1
   .MousePointer = 99
   .MouseIcon = Image2.Picture
End With
End Sub

Private Sub Label1_Click()
Dim Scr_hDC As Long
Dim ShellDoc As Long
Dim DocName As String
DocName = "mailto: osama_nt@intouch.com"
Scr_hDC = GetDesktopWindow()
ShellDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "C:\", 1)
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(1, False, X, Y)
With Label1
  '.FontBold = True
  .FontUnderline = True
  .ForeColor = vbBlue
  .MousePointer = 99
  .MouseIcon = Image2.Picture
End With
End Sub

Private Sub Label2_Click()
Dim Scr_hDC As Long
Dim ShellDoc As Long
Dim DocName As String
           
DocName = "http://planet-source-code.com/vb/scripts/ShowCode.asp?lngWId=1&txtCodeId=10456"
Scr_hDC = GetDesktopWindow()
ShellDoc = ShellExecute(Scr_hDC, "Open", DocName, "", "C:\", 1)

End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseMove(1, False, X, Y)
With Label2
  .ForeColor = vbGreen
  .MousePointer = 99
  .MouseIcon = Image2.Picture
End With
End Sub

Private Sub Label3_Click()
Call Image1_Click
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With Label3
   .MousePointer = 99
   .MouseIcon = Image2.Picture
End With

End Sub

Private Sub Label4_DblClick()
Unload Me
End Sub

Private Sub Label4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Call Form_MouseDown(1, False, X, Y)
End Sub
