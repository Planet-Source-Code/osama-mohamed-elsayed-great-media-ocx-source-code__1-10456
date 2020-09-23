VERSION 5.00
Begin VB.UserControl Media 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   Picture         =   "UserControl1.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   495
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   -120
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   -120
      Top             =   0
   End
End
Attribute VB_Name = "Media"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
' this OCX source code to play, pause,close and other processing
' about Media files and you can by it to record a file
' but that's not all it will playing at soon all media files as
' ra, ram and other
' but if you satisified about this ocx vote me :)
' if you have any questions or if you want to contact with me
' me E-mail: osama_nt@intouch.com or ICQ#: 64529562
'Detrmine the type of media
Dim TypeOfMedia As Integer ' 1 = Sound files / 2 = Video Files
'Default Property Values:
Const m_def_Repeat = False
'Property Variables:
Dim m_Repeat As Boolean
Dim m_ControlHwnd As Long
'Event Declarations:
Event Notify()
' some variables
Dim Media_FileName As String
Dim MediaControlHwnd  As Long
Dim State As Boolean
Dim CurrentDeviceID As Long
Dim RecordDeviceID As Long
Public Function StartRecord() As Long
Dim MCIOpen As MCI_OPEN_PARMS
Dim MCIRecord As MCI_RECORD_PARMS
On Local Error Resume Next
Dim ReturnVal As Long
CloseAll
With MCIOpen
  .dwCallback = 0
  .lpstrAlias = ""
  .lpstrDeviceType = "waveaudio"
  .lpstrElementName = ""
  .wDeviceID = 0
End With
ReturnVal = mciSendCommand(0, MCI_OPEN, _
                        MCI_WAIT Or MCI_OPEN_TYPE Or MCI_OPEN_ELEMENT, _
                        MCIOpen)
RecordDeviceID = MCIOpen.wDeviceID

With MCIRecord
   .dwCallback = 0
   .dwFrom = 0
   .dwTo = 0
End With
ReturnVal = mciSendCommand(RecordDeviceID, MCI_RECORD, _
               MCI_NOTIFY, MCIRecord)
               
If ReturnVal Then
   Dim B As String * 255
   mciGetErrorString ReturnVal, B, 255
   MsgBox B
End If
End Function
Public Sub PuaseRecord()
Dim ReturnVal As Long
Dim MCIGeneric As MCI_GENERIC_PARMS
MCIGeneric.dwCallback = 0
ReturnVal = mciSendCommand(RecordDeviceID, MCI_PAUSE, MCI_WAIT, _
            MCIGeneric)
End Sub
Public Sub ResumeRecord()
Dim ReturnVal As Long
Dim MCIGeneric As MCI_GENERIC_PARMS
MCIGeneric.dwCallback = 0
ReturnVal = mciSendCommand(RecordDeviceID, MCI_RESUME, MCI_WAIT, _
            MCIGeneric)
End Sub
Public Sub CloseRecord()
Dim ReturnVal As Long
Dim MCIGeneric As MCI_GENERIC_PARMS
MCIGeneric.dwCallback = 0
ReturnVal = mciSendCommand(RecordDeviceID, MCI_STOP, MCI_WAIT, _
                     MCIGeneric)
ReturnVal = mciSendCommand(RecordDeviceID, MCI_CLOSE, MCI_WAIT, _
            MCIGeneric)
End Sub
Public Sub SaveRecordToFile(FileNameToRecord As String)
Dim MCISaveRecord As MCI_SAVE_PARMS
Dim MCIGeneric As MCI_GENERIC_PARMS
Dim ReturnVal As Long

MCIGeneric.dwCallback = 0
ReturnVal = mciSendCommand(RecordDeviceID, MCI_STOP, MCI_WAIT, _
                     MCIGeneric)
                     
With MCISaveRecord
  .dwCallback = 0
  .lpFileName = FileNameToRecord
End With
ReturnVal = mciSendCommand(RecordDeviceID, MCI_Save, MCI_WAIT Or _
            MCI_SAVE_FILE, MCISaveRecord)
ReturnVal = mciSendCommand(RecordDeviceID, MCI_CLOSE, MCI_WAIT, _
            MCIGeneric)
End Sub
Private Function GetDeviceId(TypeOfDevice As String) As Long
GetDeviceId = mciGetDeviceID(TypeOfDevice)
End Function
Public Sub Author()
  Authorfrm.Show 1
End Sub
Private Sub Timer1_Timer()
If m_Repeat = True And Length = Position Then
   OpenMediaFile Media_FileName, MediaControlHwnd
   PlayIt
End If
End Sub

Private Sub Timer2_Timer()
RaiseEvent Notify
End Sub
Public Property Get EnableNotify() As Boolean
    EnableNotify = Timer2.Enabled
End Property

Public Property Let EnableNotify(ByVal New_EnableNotify As Boolean)
    Timer2.Enabled() = New_EnableNotify
    PropertyChanged "EnableNotify"
End Property
Public Sub PlayIt()
Dim MCISoundPlay As MCI_PLAY_PARMS
Dim MCIVideoPlay As MCI_ANIM_PLAY_PARMS
Dim ReturnVal As Long
Select Case TypeOfMedia
Case 1 ' Sound Files
With MCISoundPlay
  .dwCallback = 0
  .dwFrom = 0
  .dwTo = 0
End With
ReturnVal = mciSendCommand(CurrentDeviceID, MCI_PLAY, _
             MCI_NOTIFY, MCISoundPlay)
Case 2 ' Video Files
With MCIVideoPlay
  .dwCallback = 0
  .dwFrom = 0
  .dwTo = 0
End With
ReturnVal = mciSendCommand(CurrentDeviceID, MCI_PLAY, _
             MCI_NOTIFY, MCIVideoPlay)
End Select
If ReturnVal Then
 Dim r As String * 255
 mciGetErrorString ReturnVal, r, 255
 MsgBox r
 
End If
Timer1.Enabled = True
Timer2.Enabled = True
End Sub


Public Sub ResumeIt()
Dim ReturnVal As Long
Dim MCIGeneric As MCI_GENERIC_PARMS
Timer1.Enabled = True
Timer2.Enabled = True
MCIGeneric.dwCallback = 0
  ReturnVal = mciSendCommand(CurrentDeviceID, MCI_RESUME, _
     MCI_NOTIFY Or MCI_WAIT, MCIGeneric)
     
End Sub
Public Sub CloseIt()
Dim MCIGeneric As MCI_GENERIC_PARMS
Dim ReturnVal As Long
Timer1.Enabled = False
Timer2.Enabled = False
MCIGeneric.dwCallback = 0
ReturnVal = mciSendCommand(CurrentDeviceID, MCI_STOP, MCI_WAIT, MCIGeneric)
ReturnVal = mciSendCommand(CurrentDeviceID, MCI_CLOSE, MCI_WAIT, MCIGeneric)
End Sub
Private Sub UserControl_Initialize()
Dim ReturnVal As Long
Dim MCISetDevice As MCI_SET_PARMS
TypeOfMedia = 1
With MCISetDevice
  .dwCallback = 0
  .dwAudio = MCI_SET_AUDIO_ALL
End With
ReturnVal = mciSendCommand(MCI_ALL_DEVICE_ID, MCI_SET, _
                MCI_SET_AUDIO, MCISetDevice)
End Sub

Private Sub UserControl_InitProperties()
        m_Repeat = m_def_Repeat
    

End Sub



Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
        m_Repeat = PropBag.ReadProperty("Repeat", m_def_Repeat)
    Timer2.Enabled = PropBag.ReadProperty("EnableNotify", False)
    
End Sub

Private Sub UserControl_Resize()
Width = 495
Height = 480
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Repeat", m_Repeat, m_def_Repeat)
        Call PropBag.WriteProperty("EnableNotify", Timer2.Enabled, False)
End Sub
Public Function OpenMediaFile(MediaFileName As String, Optional ControlHwnd As Long) As Long
Dim MCIVideoOpen As MCI_ANIM_OPEN_PARMS
Dim MCISoundOpen As MCI_OPEN_PARMS
Dim MCIGeneric As MCI_GENERIC_PARMS
Dim MCIRecord As MCI_RECORD_PARMS
Dim MCISave As MCI_SAVE_PARMS
Dim MCIWindow As MCI_ANIM_WINDOW_PARMS
Dim MCIVideoRect As MCI_ANIM_RECT_PARMS
Dim ReturnVal As Long
Dim DeviceType As String
Dim MediaType As String
Dim FName As String
MediaControlHwnd = ControlHwnd
'State = Stretch
CloseAll
If MediaFileName = "" Then Exit Function
If UCase$(Right$(MediaFileName, 4)) = ".AVI" Then
       MediaType = "AviVideo"
ElseIf UCase$(Right$(MediaFileName, 4)) = ".RMI" Or _
       UCase$(Right$(MediaFileName, 4)) = ".MID" Or _
       UCase$(Right$(MediaFileName, 5)) = ".MIDI" Then
       MediaType = "sequencer"
       ControlHwnd = 0
ElseIf UCase$(Right$(MediaFileName, 4)) = ".WAV" Then
       MediaType = "WaveAudio"
Else
       MediaType = "MPEGVideo"
End If
Media_FileName = MediaFileName
If ControlHwnd Then 'There is a video File
  TypeOfMedia = 2
 With MCIVideoOpen
   .dwCallback = 0
   .hWndParent = ControlHwnd
   .dwStyle = WS_CHILD
   .lpstrAlias = ""
   .lpstrDeviceType = MediaType
   .lpstrElementName = MediaFileName
   .wDeviceID = 0
 End With
    
             
   ReturnVal = mciSendCommand(0, MCI_OPEN, MCI_OPEN_ELEMENT Or _
           MCI_ANIM_OPEN_PARENT, MCIVideoOpen)
   CurrentDeviceID = MCIVideoOpen.wDeviceID
  
     Dim Rec As RECT
     Dim GRet As Long
     GRet = GetWindowRect(ControlHwnd, Rec)
     
     MCIGeneric.dwCallback = 0
    
     ' Detrmine who's window parent
   With MCIWindow
       .dwCallback = 0
       .hwnd = ControlHwnd
       .lpstrText = ""
   End With
   ReturnVal = mciSendCommand(CurrentDeviceID, MCI_WINDOW, _
             MCI_ANIM_WINDOW_HWND, MCIWindow)
  If ReturnVal Then
     Dim r As String * 255
  mciGetErrorString ReturnVal, r, 255 ' you can here read the Variable r and it will tell you what's error happened
  OpenMediaFile = -1 ' the function failed in open the file other or 0 means success
  Exit Function
  End If
 
Else ' just Sound
TypeOfMedia = 1
With MCISoundOpen
  .dwCallback = ControlHwnd
  .wDeviceID = 0
  .lpstrDeviceType = MediaType
  .lpstrElementName = MediaFileName
  .lpstrAlias = ""
End With
ReturnVal = mciSendCommand(0, MCI_OPEN, MCI_WAIT Or MCI_OPEN_TYPE Or _
          MCI_OPEN_ELEMENT, MCISoundOpen)
CurrentDeviceID = MCISoundOpen.wDeviceID
End If
Exit Function
If ReturnVal Then
 OpenMediaFile = -1
MsgBox "You have a problem with this file, or this's Not Media File"
End If
End Function
Public Sub PauseIt()
Dim ReturnVal As Long
Dim MCIGeneric As MCI_GENERIC_PARMS
Timer1.Enabled = False
Timer2.Enabled = False
MCIGeneric.dwCallback = 0
ReturnVal = mciSendCommand(CurrentDeviceID, MCI_PAUSE, _
      MCI_NOTIFY Or MCI_WAIT, MCIGeneric)
End Sub
Public Sub PlayFrom(FromPos As Long)
Dim MCISoundPlay As MCI_PLAY_PARMS
Dim MCIVideoPlay As MCI_ANIM_PLAY_PARMS
Dim ReturnVal As Long
'CloseIt
'OpenMediaFile Media_FileName, MediaControlHwnd
Select Case TypeOfMedia
Case 1
With MCISoundPlay
  .dwCallback = 0
  .dwFrom = FromPos
'  .dwTo = 0
End With
ReturnVal = mciSendCommand(CurrentDeviceID, MCI_PLAY, _
          MCI_FROM, MCISoundPlay)
Case 2
With MCIVideoPlay
  .dwCallback = 0
  .dwFrom = FromPos
'  .dwTo = 0
End With
ReturnVal = mciSendCommand(CurrentDeviceID, MCI_PLAY, _
          MCI_FROM, MCIVideoPlay)
End Select
Timer1.Enabled = True
Timer2.Enabled = True
End Sub
Public Sub PlayTo(ToPos As Long)
Dim MCISoundPlay As MCI_PLAY_PARMS
Dim MCIVideoPlay As MCI_ANIM_PLAY_PARMS
Dim ReturnVal As Long
'CloseIt
'OpenMediaFile Media_FileName, MediaControlHwnd
Select Case TypeOfMedia
Case 1 'Sound Files
With MCISoundPlay
  .dwCallback = 0
'  .dwFrom = 0
  .dwTo = ToPos
End With
ReturnVal = mciSendCommand(CurrentDeviceID, MCI_PLAY, _
          MCI_TO, MCISoundPlay)
Case 2 ' Video Files
With MCIVideoPlay
   .dwCallback = 0
'   .dwFrom = 0
   .dwTo = ToPos
End With
ReturnVal = mciSendCommand(CurrentDeviceID, MCI_PLAY, _
          MCI_TO, MCIVideoPlay)
End Select
If ReturnVal Then
  Dim r As String * 255
   mciGetErrorString ReturnVal, r, 255
   MsgBox r
   Exit Sub
 End If
Timer1.Enabled = True
Timer2.Enabled = True
End Sub
Public Sub CloseAll()
Dim MCIGeneric As MCI_GENERIC_PARMS
Dim ReturnVal As Long
Timer1.Enabled = False
Timer2.Enabled = False
MCIGeneric.dwCallback = 0
ReturnVal = mciSendCommand(MCI_ALL_DEVICE_ID, MCI_STOP, MCI_WAIT, MCIGeneric)
ReturnVal = mciSendCommand(MCI_ALL_DEVICE_ID, MCI_CLOSE, MCI_WAIT, MCIGeneric)

End Sub

Public Function Length() As Long
Dim ReturnVal As Long
Dim MCIStatus As MCI_STATUS_PARMS
Dim Buffer As String * 255
With MCIStatus
   .dwCallback = 0
   .dwItem = MCI_STATUS_LENGTH
End With
ReturnVal = mciSendCommand(CurrentDeviceID, MCI_STATUS, _
           MCI_STATUS_ITEM Or MCI_STATUS_LENGTH, MCIStatus)
If ReturnVal Then
  mciGetErrorString ReturnVal, Buffer, 255
MsgBox Buffer
 Timer2.Enabled = False
  Timer1.Enabled = False
End If
Length = MCIStatus.dwReturn
End Function

Public Function Position() As Variant
Dim ReturnVal As Long
Dim MCIStatus As MCI_STATUS_PARMS
Dim Buffer As String * 255
With MCIStatus
   .dwCallback = 0
   .dwItem = MCI_STATUS_POSITION
End With
ReturnVal = mciSendCommand(CurrentDeviceID, MCI_STATUS, _
           MCI_STATUS_ITEM Or MCI_STATUS_POSITION, MCIStatus)
If ReturnVal Then
  mciGetErrorString ReturnVal, Buffer, 255
  MsgBox Buffer
  Timer2.Enabled = False
  Timer1.Enabled = False
End If
Position = MCIStatus.dwReturn
End Function

Public Property Get Repeat() As Boolean
    Repeat = m_Repeat
End Property

Public Property Let Repeat(ByVal New_Repeat As Boolean)
    m_Repeat = New_Repeat
        PropertyChanged "Repeat"
End Property

