VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{A59A6906-58E2-11D4-8F15-D4E343FCF02F}#30.0#0"; "OMEAMedia.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4515
   ClientLeft      =   1485
   ClientTop       =   900
   ClientWidth     =   7785
   LinkTopic       =   "Form1"
   ScaleHeight     =   4515
   ScaleWidth      =   7785
   Begin OMEAMedia.Media Media1 
      Left            =   120
      Top             =   3600
      _ExtentX        =   873
      _ExtentY        =   847
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Stop And Save To File"
      Height          =   495
      Index           =   3
      Left            =   0
      TabIndex        =   15
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Resume Record"
      Height          =   375
      Index           =   2
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Pause Record"
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   13
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Record"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Repeat"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2880
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6600
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Open Media File"
      Height          =   1095
      Left            =   6360
      TabIndex        =   10
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close All"
      Height          =   495
      Index           =   4
      Left            =   720
      TabIndex        =   9
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   8
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Index           =   0
      Left            =   6360
      TabIndex        =   7
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Play To"
      Height          =   375
      Index           =   1
      Left            =   6360
      TabIndex        =   6
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Play From"
      Height          =   375
      Index           =   0
      Left            =   6360
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Close"
      Height          =   495
      Index           =   3
      Left            =   1800
      TabIndex        =   4
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Resume"
      Height          =   495
      Index           =   2
      Left            =   2880
      TabIndex        =   3
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pause"
      Height          =   495
      Index           =   1
      Left            =   3960
      TabIndex        =   2
      Top             =   3840
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Play"
      Height          =   495
      Index           =   0
      Left            =   5040
      TabIndex        =   1
      Top             =   3840
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3255
      Left            =   1560
      TabIndex        =   0
      Top             =   240
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
Media1.Repeat = Check1.Value
End Sub
' You can now play all media files such as standard formats
' such as AVI,MID,MIDI,WAV and compression formats such as
' MP3, Mpeg, Mpg .... and other
' and you can do this without compund or complected code
' just simple code :)
' without using Dll files or any other component
' if you want a help about that
' you can contact with me E-mail osama_nt@intouch.com icq# 64529562
' and soon the other properties will be avilable you can't imagine that
' that's will be great for all programmers

Private Sub Command1_Click(Index As Integer)
With Media1
Select Case Index
Case 0
 .StartRecord
Case 1
  .PuaseRecord
Case 2
  .ResumeRecord
Case 3
With CommonDialog1
   .DialogTitle = "Save File"
   .Filter = "Wav Files|*.wav"
   .FileName = ""
   .ShowSave
 If .FileName = "" Then Exit Sub
End With
   .SaveRecordToFile CommonDialog1.FileName
End Select
End With
End Sub

Private Sub Command2_Click(Index As Integer)
Dim i As Integer
With Media1
Select Case Index
Case 0
  .PlayIt ' play file from start to end
Case 1
  .PauseIt ' pause playing the file
Case 2
  .ResumeIt ' resume it
Case 3
  .CloseIt ' close the current file
Case 4
  .CloseAll ' close all files if you have
End Select
End With
End Sub

Private Sub Command3_Click(Index As Integer)
With Media1
Select Case Index
Case 0
  .PlayFrom Val(Text1(0).Text) ' play from position
Case 1
  .PlayTo Val(Text1(1).Text) 'play to position
End Select
End With

End Sub


Private Sub Command5_Click()
Dim Fname As String
With CommonDialog1
  .DialogTitle = "open Media Files"
  .Filter = "All Files|*.*"
  .InitDir = "C:\"
  .FileName = ""
  .ShowOpen
  If .FileName = "" Then Exit Sub
  Fname = .FileName
End With
If UCase$(Right$(Fname, 4)) = ".WAV" Or _
      UCase$(Right$(Fname, 4)) = ".MP3" Or _
      UCase$(Right$(Fname, 4)) = ".MID" Or _
      UCase$(Right$(Fname, 5)) = ".MIDI" Or _
      UCase$(Right$(Fname, 4)) = ".RMI" Then
      Media1.OpenMediaFile Fname ', Frame1.hWnd
ElseIf UCase$(Right$(Fname, 4)) = ".AVI" Or _
       UCase$(Right$(Fname, 4)) = ".MPG" Or _
       UCase$(Right$(Fname, 5)) = ".MPEG" Or _
       UCase$(Right$(Fname, 4)) = ".ASF" Or _
       UCase$(Right$(Fname, 4)) = ".DAT" Then
       ' when you put in the second argument True the Video file
       ' will stretch in the control size (e.g Frame1)
       ' if you put False will show the video in it's size (Normal size)
       Media1.OpenMediaFile Fname, Frame1.hWnd
End If
End Sub

Private Sub Form_Load()
Media1.Author
End Sub

Private Sub Form_Unload(Cancel As Integer)
Media1.CloseAll
Media1.Author
End Sub

Private Sub Media1_Notify()
Caption = Media1.Length & "/" & Media1.Position
End Sub


Private Sub SpinButton1_Change()

End Sub
