Attribute VB_Name = "Types"
Option Explicit
' Custom Types
Public Type POINTAPI
        X As Long
        Y As Long
End Type

Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type MMCKINFO
    ckid As Long
    ckSize As Long
    fccType As Long
    dwDataOffset As Long
    dwFlags As Long
End Type
Public Type MMTIME
        wType As Long
        u As Long
End Type

Public Type MMIOINFO
        dwFlags As Long
        fccIOProc As Long
        pIOProc As Long
        wErrorRet As Long
        htask As Long
        cchBuffer As Long
        pchBuffer As String
        pchNext As String
        pchEndRead As String
        pchEndWrite As String
        lBufOffset As Long
        lDiskOffset As Long
        adwInfo(4) As Long
        dwReserved1 As Long
        dwReserved2 As Long
        hmmio As Long
End Type
Public Type MCI_OPEN_PARMS
        dwCallback As Long
        wDeviceID As Long
        lpstrDeviceType As String
        lpstrElementName As String
        lpstrAlias As String
End Type
Public Type MCI_ANIM_OPEN_PARMS
        dwCallback As Long
        wDeviceID As Long
        lpstrDeviceType As String
        lpstrElementName As String
        lpstrAlias As String
        dwStyle As Long
        hWndParent As Long
End Type
Public Type MCI_ANIM_PLAY_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
        dwSpeed As Long
End Type
Public Type MCI_ANIM_RECT_PARMS
        dwCallback As Long
        rc As RECT
End Type
Public Type MCI_ANIM_STEP_PARMS
        dwCallback As Long
        dwFrames As Long
End Type
Public Type MCI_ANIM_UPDATE_PARMS
        dwCallback As Long
        rc As RECT
        hdc As Long
End Type
Public Type MCI_ANIM_WINDOW_PARMS
        dwCallback As Long
        hwnd As Long
        nCmdShow As Long
        lpstrText As String
End Type
Public Type MCI_BREAK_PARMS
        dwCallback As Long
        nVirtKey As Long
        hwndBreak As Long
End Type
Public Type MCI_GENERIC_PARMS
        dwCallback As Long
End Type
Public Type MCI_GETDEVCAPS_PARMS
        dwCallback As Long
        dwReturn As Long
        dwIten As Long
End Type
Public Type MCI_INFO_PARMS
        dwCallback As Long
        lpstrReturn As String
        dwRetSize As Long
End Type
Public Type MCI_LOAD_PARMS
        dwCallback As Long
        lpFileName As String
End Type
Public Type MCI_OVLY_LOAD_PARMS
        dwCallback As Long
        lpFileName As String
        rc As RECT
End Type
Public Type MCI_OVLY_OPEN_PARMS
        dwCallback As Long
        wDeviceID As Long
        lpstrDeviceType As String
        lpstrElementName As String
        lpstrAlias As String
        dwStyle As Long
        hWndParent As Long
End Type
Public Type MCI_OVLY_RECT_PARMS
        dwCallback As Long
        rc As RECT
End Type
Public Type MCI_OVLY_SAVE_PARMS
        dwCallback As Long
        lpFileName As String
        rc As RECT
End Type
Public Type MCI_OVLY_WINDOW_PARMS
        dwCallback As Long
        hwnd As Long
        nCmdShow As Long
        lpstrText As String
End Type
Public Type MCI_PLAY_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
End Type
Public Type MCI_RECORD_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
End Type
Public Type MCI_SAVE_PARMS
        dwCallback As Long
        lpFileName As String
End Type
Public Type MCI_SEEK_PARMS
        dwCallback As Long
        dwTo As Long
End Type
Public Type MCI_SEQ_SET_PARMS
        dwCallback As Long
        dwTimeFormat As Long
        dwAudio As Long
        dwTempo As Long
        dwPort As Long
        dwSlave As Long
        dwMaster As Long
        dwOffset As Long
End Type
Public Type MCI_SET_PARMS
        dwCallback As Long
        dwTimeFormat As Long
        dwAudio As Long
End Type
Public Type MCI_SOUND_PARMS
        dwCallback As Long
        lpstrSoundName As String
End Type
Public Type MCI_STATUS_PARMS
        dwCallback As Long
        dwReturn As Long
        dwItem As Long
        dwTrack As Integer
End Type
Public Type MCI_SYSINFO_PARMS
        dwCallback As Long
        lpstrReturn As String
        dwRetSize As Long
        dwNumber As Long
        wDeviceType As Long
End Type
Public Type MCI_VD_ESCAPE_PARMS
        dwCallback As Long
        lpstrCommand As String
End Type
Public Type MCI_VD_PLAY_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
        dwSpeed As Long
End Type
Public Type MCI_VD_STEP_PARMS
        dwCallback As Long
        dwFrames As Long
End Type
Public Type MCI_WAVE_DELETE_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
End Type
Public Type MCI_WAVE_OPEN_PARMS
        dwCallback As Long
        wDeviceID As Long
        lpstrDeviceType As String
        lpstrElementName As String
        lpstrAlias As String
        dwBufferSeconds As Long
End Type
Public Type MCI_WAVE_SET_PARMS
        dwCallback As Long
        dwTimeFormat As Long
        dwAudio As Long
        wInput As Long
        wOutput As Long
        wFormatTag As Integer
        wReserved2 As Integer
        nChannels As Integer
        wReserved3 As Integer
        nSamplesPerSec As Long
        nAvgBytesPerSec As Long
        nBlockAlign As Integer
        wReserved4 As Integer
        wBitsPerSample As Integer
        wReserved5 As Integer
End Type




