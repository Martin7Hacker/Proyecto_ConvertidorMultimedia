VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFFPresetMPG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' -- declaraciones de funciones
Private Declare Function PathFindExtension Lib "shlwapi" Alias "PathFindExtensionA" (ByVal pPath As String) As Long
Private Declare Function lstrcpyA Lib "kernel32" (ByVal RetVal As String, ByVal Ptr As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal Ptr As Any) As Long

Private mPresetsCount               As Integer
Private mPresetDescription          As String
Private mPresetExtension            As String
Private mPresetCategory             As String
Private mPresetParameters           As String


Enum ePresets
    [01_CDWavStereo] = 0
    [02_GigabeatFS] = 1
    [03_GigabeatWS] = 2
    [04_H1020FS] = 3
    [05_H1020WS] = 4
    [06_H105FS] = 5
    [07_H105WS] = 6
    [08_H300FS] = 7
    [09_H300WS] = 8
    [10_I5GFS] = 9
    [11_I5GWS] = 10
    [12_M4a] = 11
    [13_NeurosOSDTVHQNTSC] = 12
    [14_NeurosOSDTVHQPAL] = 13
    [15_NeurosOSDTVSFNTSC] = 14
    [16_NeurosOSDTVSFPAL] = 15
    [17_NeurosOSDTVWSHQNTSC] = 16
    [18_NeurosOSDTVWSHQPAL] = 17
    [19_NeurosOSDTVWSSFNTSC] = 18
    [20_NeurosOSDTVWSSFPAL] = 19
    [21_PalmXviDFS] = 20
    [22_PalmXviDWS] = 21
    [23_PhotoFS] = 22
    [24_PhotoWS] = 23
    [25_X5FS] = 24
    [26_X5WS] = 25
    [27_XviDAVIFS] = 26
    [28_XviDAVIWS] = 27
    [29_XviDAVIWSAna] = 28
    [30_XviDPSP169] = 29
    [31_XviDPSP43] = 30
    [32_ac3dvd192] = 31
    [33_ac3dvd384] = 32
    [34_divx] = 33
    [35_dvntsc] = 34
    [36_dvpal] = 35
    [37_e200FS] = 36
    [38_e200WS] = 37
    [39_mp3] = 38
    [40_mp3m] = 39
    [41_nanoFS] = 40
    [42_nanoWS] = 41
    [43_ntscdvdgoodqFS] = 42
    [44_ntscdvdgoodqWS] = 43
    [45_ntscdvdhqFS] = 44
    [46_ntscdvdhqWS] = 45
    [47_ntscdvdlq] = 46
    [48_ntscvcdhq] = 47
    [49_paldvdgoodqFS] = 48
    [50_paldvdgoodqWS] = 49
    [51_paldvdhqFS] = 50
    [52_paldvdhqWS] = 51
    [53_paldvdlq] = 52
    [54_palvcdhq] = 53
    [55_qmov] = 54
    [56_wma] = 55
    [57_wmv] = 56
    [58_x264HQFS] = 57
    [59_x264HQPSP43] = 58
    [60_x264HQWS] = 59
    [61_cdma3g] = 60
    [62_XviDAVIZENFS] = 61
    [63_XviDAVIZENWS] = 62
    [64_BlackberryCurvefs] = 63
    [65_BlackberryCurvews] = 64
    [66_lgchocolate] = 65
    [67_blackberrymp3] = 66
    [68_blackberryws] = 67
    [69_blackberryfs] = 68
    [70_flvwebFS] = 69
    [71_flvwebWS] = 70
    [72_x264HQPSP169robertswain] = 71
    [73_iPodiTunesSmallCRF21FS] = 72
    [74_iPodiTunesSmallCRF21WS] = 73
    [75_iPodiTunesSmallCRF21WSAna] = 74
    [76_iPodiTunesSmallCRF21WSLBFix] = 75
    [77_iPodiTunesTVOutCRF21FS] = 76
    [78_iPodiTunesTVOutCRF21WS] = 77
    [79_iPodiTunesTVOutCRF21WSAna] = 78
    [80_iPodiTunesTVOutCRF21WSLBFix] = 79
    [81_iPodXviDFS] = 80
    [82_iPodXviDWS] = 81
    [83_iPodXviDWSAna] = 82
    [84_ogg4] = 83
    [85_ogg2] = 84
    [86_ogg5] = 85
    [87_ogg1] = 86
    [88_wma1] = 87
    [89_wmv3a] = 88
    [90_wma2] = 89
    [91_wmv3d] = 90
    [92_wmv3c] = 91
    [93_wmv3b] = 92
End Enum



Property Get PresetsCount() As Integer
    PresetsCount = mPresetsCount
End Property

Private Sub Class_Initialize()
    mPresetsCount = 93
End Sub

Property Get PresetDescription() As String
    PresetDescription = mPresetDescription
End Property
    
Property Get PresetExtension() As String
    PresetExtension = mPresetExtension
End Property

Property Get PresetCategory() As String
    PresetCategory = mPresetCategory
End Property

Property Get PresetParameters() As String
    PresetParameters = mPresetParameters
End Property

Function setPreset(lPreset As ePresets) As String
    If lPreset = 0 Then
        mPresetDescription = "Wav for CD"
        mPresetParameters = "-vn -ar 44100"
        mPresetExtension = "wav"
        mPresetCategory = "Audio"
    End If
    If lPreset = 1 Then
        mPresetDescription = "RB Toshiba Gigabeat F/X 4:3"
        mPresetParameters = "-acodec libmp3lame -ab 128 -ar 44100 -vcodec mpeg2video -s 320x240 -b 600kb -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 2 Then
        mPresetDescription = "RB Toshiba Gigabeat F/X 16:9"
        mPresetParameters = "-acodec libmp3lame -ab 128 -ar 44100 -vcodec mpeg2video -s 320x176 -b 600kb -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 3 Then
        mPresetDescription = "RB iRiver H10 20GB 4:3"
        mPresetParameters = "-acodec libmp3lame -ab 128 -ar 44100 -vcodec mpeg2video -s 160x128 -b 224kb -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 4 Then
        mPresetDescription = "RB iRiver H10 20GB 16:9"
        mPresetParameters = "-acodec libmp3lame -ab 128 -ar 44100 -vcodec mpeg2video -s 160x96 -b 224kb -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 5 Then
        mPresetDescription = "RB iRiver H10 5/6GB 4:3"
        mPresetParameters = "-acodec libmp3lame -ab 128 -ar 44100 -vcodec mpeg2video -s 128x96 -b 224kb -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 6 Then
        mPresetDescription = "RB iRiver H10 5/6GB 16:9"
        mPresetParameters = "-acodec libmp3lame -ab 128 -ar 44100 -vcodec mpeg2video -s 128x80 -b 224kb -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 7 Then
        mPresetDescription = "RB iRiver H300 4:3"
        mPresetParameters = "-acodec libmp3lame -ab 96 -ar 44100 -vcodec mpeg2video -s 224x176 -b 224kb -r 10 -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 8 Then
        mPresetDescription = "RB iRiver H300 16:9"
        mPresetParameters = "-acodec libmp3lame -ab 96 -ar 44100 -vcodec mpeg2video -s 224x128 -b 256kb -r 15 -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 9 Then
        mPresetDescription = "RB Apple iPod Video 4:3"
        mPresetParameters = "-acodec libmp3lame -ab 128 -ar 44100 -vcodec mpeg2video -s 320x240 -b 400kb -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 10 Then
        mPresetDescription = "RB Apple iPod Video 16:9"
        mPresetParameters = "-acodec libmp3lame -ab 128 -ar 44100 -vcodec mpeg2video -s 320x176 -b 400kb -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 11 Then
        mPresetDescription = "M4a"
        mPresetParameters = "-vn -acodec libfaac -ab 112k -ac 2"
        mPresetExtension = "m4a"
        mPresetCategory = "Audio"
    End If
    If lPreset = 12 Then
        mPresetDescription = "Neuros High Quality NTSC (4:3)"
        mPresetParameters = "-r 29.97 -vcodec libxvid -s 640x480 -aspect 4:3 -maxrate 2500k -b 2000k -qmin 3 -qmax 5 -bufsize 4096 -mbd 2 -flags +4mv -trellis -aic -cmp 2 -subcmp 2 -g 300 -acodec libfaac -ar 48000 -ab 128k -ac 2"
        mPresetExtension = "mp4"
        mPresetCategory = "Neuros OSD"
    End If
    If lPreset = 13 Then
        mPresetDescription = "Neuros High Quality PAL (4:3)"
        mPresetParameters = "-r 25 -vcodec libxvid -s 640x480 -aspect 4:3 -maxrate 2500k -b 2000k -qmin 3 -qmax 5 -bufsize 4096 -mbd 2 -flags +4mv -trellis -aic -cmp 2 -subcmp 2 -g 300 -acodec libfaac -ar 48000 -ab 128k -ac 2"
        mPresetExtension = "mp4"
        mPresetCategory = "Neuros OSD"
    End If
    If lPreset = 14 Then
        mPresetDescription = "Neuros Small File NTSC (4:3)"
        mPresetParameters = "-b 800k -r 29.97 -s 320x240 -aspect 4:3 -vcodec libxvid -ar 48000 -ab 80k -ac 2 -acodec libfaac"
        mPresetExtension = "mp4"
        mPresetCategory = "Neuros OSD"
    End If
    If lPreset = 15 Then
        mPresetDescription = "Neuros Small File PAL (4:3)"
        mPresetParameters = "-b 800k -r 25 -s 320x240 -aspect 4:3 -vcodec libxvid -ar 48000 -ab 80k -ac 2 -acodec libfaac"
        mPresetExtension = "mp4"
        mPresetCategory = "Neuros OSD"
    End If
    If lPreset = 16 Then
        mPresetDescription = "Neuros High Quality NTSC (16:9)"
        mPresetParameters = "-r 29.97 -vcodec libxvid -s 704x384 -aspect 16:9 -maxrate 3000k -b 2500k -qmin 3 -qmax 5 -bufsize 4096 -mbd 2 -bf 2 -flags +4mv -trellis -aic -cmp 2 -subcmp 2 -g 300 -acodec libfaac -ar 48000 -ab 128k -ac 2"
        mPresetExtension = "mp4"
        mPresetCategory = "Neuros OSD"
    End If
    If lPreset = 17 Then
        mPresetDescription = "Neuros High Quality PAL (16:9)"
        mPresetParameters = "-r 25 -vcodec libxvid -s 704x384 -aspect 16:9 -maxrate 3000k -b 2500k -qmin 3 -qmax 5 -bufsize 4096 -mbd 2 -bf 2 -flags +4mv -trellis -aic -cmp 2 -subcmp 2 -g 300 -acodec libfaac -ar 48000 -ab 128k -ac 2"
        mPresetExtension = "mp4"
        mPresetCategory = "Neuros OSD"
    End If
    If lPreset = 18 Then
        mPresetDescription = "Neuros Small File NTSC (16:9)"
        mPresetParameters = "-b 800k -r 29.97 -s 352x240 -aspect 16:9 -vcodec libxvid -ar 48000 -ab 80k -ac 2 -acodec libfaac"
        mPresetExtension = "mp4"
        mPresetCategory = "Neuros OSD"
    End If
    If lPreset = 19 Then
        mPresetDescription = "Neuros Small File PAL (16:9)"
        mPresetParameters = "-b 800k -r 25 -s 352x240 -aspect 16:9 -vcodec libxvid -ar 48000 -ab 80k -ac 2 -acodec libfaac"
        mPresetExtension = "mp4"
        mPresetCategory = "Neuros OSD"
    End If
    If lPreset = 20 Then
        mPresetDescription = "Xvid for Palm (4:3)"
        mPresetParameters = "-r 29.97 -vcodec libxvid -s 428x320 -aspect 4:3 -maxrate 550k -b 500k -qmin 3 -qmax 5 -bufsize 4096 -mbd 2 -flags +4mv -trellis -aic -cmp 2 -subcmp 2 -g 300 -acodec libfaac -ar 44100 -ab 80k -ac 2"
        mPresetExtension = "mp4"
        mPresetCategory = "Palm"
    End If
    If lPreset = 21 Then
        mPresetDescription = "XviD for Palm (3:2)"
        mPresetParameters = "-r 29.97 -vcodec libxvid -s 480x320 -aspect 3:2 -maxrate 450k -b 430k -qmin 3 -qmax 5 -bufsize 4096 -mbd 2 -flags +4mv -trellis -aic -cmp 2 -subcmp 2 -g 300 -acodec libfaac -ar 44100 -ab 80k -ac 2"
        mPresetExtension = "mp4"
        mPresetCategory = "Palm"
    End If
    If lPreset = 22 Then
        mPresetDescription = "RB Apple iPod Photo/Color 4:3"
        mPresetParameters = "-acodec libmp3lame -ab 128 -ar 44100 -vcodec mpeg2video -s 224x176 -b 320kb -r 15 -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 23 Then
        mPresetDescription = "RB Apple iPod Photo/Color 16:9"
        mPresetParameters = "-acodec libmp3lame -ab 128 -ar 44100 -vcodec mpeg2video -s 224x128 -b 320kb -r 24 -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 24 Then
        mPresetDescription = "RB Cowon iAudio X5 4:3"
        mPresetParameters = "-acodec libmp3lame -ab 128 -ar 44100 -vcodec mpeg2video -s 160x128 -b 176kb -r 15 -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 25 Then
        mPresetDescription = "RB Cowon iAudio X5 16:9"
        mPresetParameters = "-acodec libmp3lame -ab 128 -ar 44100 -vcodec mpeg2video -s 160x96 -b 224kb -r 24 -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 26 Then
        mPresetDescription = "XviD in AVI (4:3)"
        mPresetParameters = "-r 29.97 -vcodec libxvid -vtag XVID -s 640x480 -aspect 4:3 -maxrate 1800k -b 1500k -qmin 3 -qmax 5 -bufsize 4096 -mbd 2 -bf 2 -flags +4mv -trellis -aic -cmp 2 -subcmp 2 -g 300 -acodec libmp3lame -ar 48000 -ab 128k -ac 2"
        mPresetExtension = "avi"
        mPresetCategory = "AVI"
    End If
    If lPreset = 27 Then
        mPresetDescription = "XviD in AVI (16:9)"
        mPresetParameters = "-r 29.97 -vcodec libxvid -vtag XVID -s 704x384 -aspect 16:9 -maxrate 1800k -b 1500k -qmin 3 -qmax 5 -bufsize 4096 -mbd 2 -bf 2 -flags +4mv -trellis -aic -cmp 2 -subcmp 2 -g 300 -acodec libmp3lame -ar 48000 -ab 128k -ac 2"
        mPresetExtension = "avi"
        mPresetCategory = "AVI"
    End If
    If lPreset = 28 Then
        mPresetDescription = "XviD in AVI (16:9 Anamorphic)"
        mPresetParameters = "-r 29.97 -croptop 58 -cropbottom 62 -vcodec libxvid -vtag XVID -s 640x272 -aspect 2.35 -maxrate 1800k -b 1500k -qmin 3 -qmax 5 -bufsize 4096 -mbd 2 -bf 2 -flags +4mv -trellis -aic -cmp 2 -subcmp 2 -g 300 -acodec libmp3lame -ar 48000 -ab 128k -ac 2"
        mPresetExtension = "avi"
        mPresetCategory = "AVI"
    End If
    If lPreset = 29 Then
        mPresetDescription = "XviD for PSP (16:9)"
        mPresetParameters = "-r 29.97 -vcodec libxvid -s 640x480 -aspect 16:9 -maxrate 1500k -b 1250k -qmin 3 -qmax 5 -bufsize 4096 -mbd 2 -flags +4mv -trellis -aic -cmp 2 -subcmp 2 -g 300 -acodec libfaac -ar 48000 -ab 80k -ac 2 -s 320x240"
        mPresetExtension = "mp4"
        mPresetCategory = "PSP"
    End If
    If lPreset = 30 Then
        mPresetDescription = "XviD for PSP (4:3)"
        mPresetParameters = "-r 29.97 -vcodec libxvid -s 640x480 -aspect 4:3 -maxrate 1500k -b 1250k -qmin 3 -qmax 5 -bufsize 4096 -mbd 2 -flags +4mv -trellis -aic -cmp 2 -subcmp 2 -g 300 -acodec libfaac -ar 48000 -ab 80k -ac 2 -s 320x240"
        mPresetExtension = "mp4"
        mPresetCategory = "PSP"
    End If
    If lPreset = 31 Then
        mPresetDescription = "Ac3 DVD - 192kbps Stereo"
        mPresetParameters = "-acodec ac3 -ab 192k -ar 48000 -ac 2"
        mPresetExtension = "ac3"
        mPresetCategory = "Audio"
    End If
    If lPreset = 32 Then
        mPresetDescription = "Ac3 DVD - 384kbps Stereo"
        mPresetParameters = "-acodec ac3 -ab 384k -ar 48000 -ac 2"
        mPresetExtension = "ac3"
        mPresetCategory = "Audio"
    End If
    If lPreset = 33 Then
        mPresetDescription = "MS Compatible AVI"
        mPresetParameters = "-f avi -acodec libmp3lame -vcodec msmpeg4 -ab 192 -b 1250"
        mPresetExtension = "avi"
        mPresetCategory = "AVI"
    End If
    If lPreset = 34 Then
        mPresetDescription = "Raw DV for NTSC 4:3"
        mPresetParameters = "-target ntsc-dv -aspect 4:3"
        mPresetExtension = "dv"
        mPresetCategory = "DV"
    End If
    If lPreset = 35 Then
        mPresetDescription = "Raw DV for PAL 4:3"
        mPresetParameters = "-target pal-dv -aspect 4:3"
        mPresetExtension = "dv"
        mPresetCategory = "DV"
    End If
    If lPreset = 36 Then
        mPresetDescription = "RB Sandisk Sansa e200 4:3"
        mPresetParameters = "-acodec libmp3lame -ab 128 -ar 44100 -vcodec mpeg2video -s 224x176 -b 320kb -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 37 Then
        mPresetDescription = "RB Sandisk Sansa e200 16:9"
        mPresetParameters = "-acodec libmp3lame -ab 128 -ar 44100 -vcodec mpeg2video -s 224x128 -b 320kb -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 38 Then
        mPresetDescription = "MP3"
        mPresetParameters = "-acodec libmp3lame -ab 160k -ac 2 -ar 44100"
        mPresetExtension = "mp3"
        mPresetCategory = "Audio"
    End If
    If lPreset = 39 Then
        mPresetDescription = "MP3 (Mono)"
        mPresetParameters = "-acodec libmp3lame -ab 32k -ac 1 -ar 22050"
        mPresetExtension = "mp3"
        mPresetCategory = "Mobile Phones"
    End If
    If lPreset = 40 Then
        mPresetDescription = "RB Apple iPod Nano 4:3"
        mPresetParameters = "-acodec libmp3lame -ab 128 -ar 44100 -vcodec mpeg2video -s 176x128 -b 256kb -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 41 Then
        mPresetDescription = "RB Apple iPod Nano 16:9"
        mPresetParameters = "-acodec libmp3lame -ab 128 -ar 44100 -vcodec mpeg2video -s 176x128 -b 256kb -strict -1"
        mPresetExtension = "mpg"
        mPresetCategory = "Rockbox"
    End If
    If lPreset = 42 Then
        mPresetDescription = "NTSC DVD (4:3)"
        mPresetParameters = "-vcodec mpeg2video -r 29.97 -s 352x480 -aspect 4:3 -b 4000k -mbd rd -trellis -mv0 -cmp 2 -subcmp 2 -acodec mp2 -ab 192k -ar 48000 -ac 2"
        mPresetExtension = "mpg"
        mPresetCategory = "DVD"
    End If
    If lPreset = 43 Then
        mPresetDescription = "NTSC DVD (16:9)"
        mPresetParameters = "-vcodec mpeg2video -r 29.97 -s 352x480 -aspect 16:9 -b 4000k -mbd rd -trellis -mv0 -cmp 2 -subcmp 2 -acodec mp2 -ab 192k -ar 48000 -ac 2"
        mPresetExtension = "mpg"
        mPresetCategory = "DVD"
    End If
    If lPreset = 44 Then
        mPresetDescription = "NTSC DVD HQ (4:3)"
        mPresetParameters = "-target ntsc-dvd -r 29.97 -s 720x480 -aspect 4:3 -b 8000k -mbd rd -trellis -mv0 -cmp 0 -subcmp 2"
        mPresetExtension = "mpg"
        mPresetCategory = "DVD"
    End If
    If lPreset = 45 Then
        mPresetDescription = "NTSC DVD HQ (16:9)"
        mPresetParameters = "-target ntsc-dvd -r 29.97 -s 720x480 -aspect 16:9 -b 8000k -g 12 -mbd rd -trellis -mv0 -cmp 0 -subcmp 2"
        mPresetExtension = "mpg"
        mPresetCategory = "DVD"
    End If
    If lPreset = 46 Then
        mPresetDescription = "NTSC DVD Fast (LQ)"
        mPresetParameters = "-target ntsc-dvd -b 5000kb -r 29.97 -s 720x480 -ar 48000 -ab 384kb"
        mPresetExtension = "mpg"
        mPresetCategory = "DVD"
    End If
    If lPreset = 47 Then
        mPresetDescription = "NTSC VCD (HQ)"
        mPresetParameters = "-target ntsc-vcd -mbd rd -trellis -mv0 -cmp 0 -subcmp 2"
        mPresetExtension = "mpg"
        mPresetCategory = "VCD"
    End If
    If lPreset = 48 Then
        mPresetDescription = "PAL DVD (4:3)"
        mPresetParameters = "-vcodec mpeg2video -r 25.00 -s 352x576  -aspect 4:3 -b 4000k -mbd rd -trellis -mv0 -cmp 2 -subcmp 2 -acodec mp2 -ab 192k -ar 48000 -ac 2"
        mPresetExtension = "mpg"
        mPresetCategory = "DVD"
    End If
    If lPreset = 49 Then
        mPresetDescription = "PAL DVD (16:9)"
        mPresetParameters = "-vcodec mpeg2video -r 25.00 -s 352x576 -aspect 16:9 -b 4000k -mbd rd -trellis -mv0 -cmp 2 -subcmp 2 -acodec mp2 -ab 192k -ar 48000 -ac 2"
        mPresetExtension = "mpg"
        mPresetCategory = "DVD"
    End If
    If lPreset = 50 Then
        mPresetDescription = "PAL DVD HQ (4:3)"
        mPresetParameters = "-target pal-dvd -aspect 4:3 -b 8000k -mbd rd -trellis -mv0 -cmp 0 -subcmp 2"
        mPresetExtension = "mpg"
        mPresetCategory = "DVD"
    End If
    If lPreset = 51 Then
        mPresetDescription = "PAL DVD HQ (16:9)"
        mPresetParameters = "-target pal-dvd -aspect 16:9 -b 8000k -mbd rd -trellis -mv0 -cmp 0 -subcmp 2"
        mPresetExtension = "mpg"
        mPresetCategory = "DVD"
    End If
    If lPreset = 52 Then
        mPresetDescription = "PAL DVD Fast (LQ)"
        mPresetParameters = "-target pal-dvd -b 5000kb -r 25 -s 720x576 -ar 48000 -ab 384kb"
        mPresetExtension = "mpg"
        mPresetCategory = "DVD"
    End If
    If lPreset = 53 Then
        mPresetDescription = "PAL VCD (HQ)"
        mPresetParameters = "-target pal-vcd -mbd rd -trellis -mv0 -cmp 0 -subcmp 2"
        mPresetExtension = "mpg"
        mPresetCategory = "VCD"
    End If
    If lPreset = 54 Then
        mPresetDescription = "Quicktime MOV"
        mPresetParameters = "-acodec libfaac -b 1250 -r 25 -ab 128"
        mPresetExtension = "mov"
        mPresetCategory = "Quicktime"
    End If
    If lPreset = 55 Then
        mPresetDescription = "WMA"
        mPresetParameters = "-vn  -acodec wmav2"
        mPresetExtension = "wma"
        mPresetCategory = "Audio"
    End If
    If lPreset = 56 Then
        mPresetDescription = "WMV2"
        mPresetParameters = "-vcodec wmv2  -acodec wmav2 -b 1000kb"
        mPresetExtension = "wmv"
        mPresetCategory = "WMV"
    End If
    If lPreset = 57 Then
        mPresetDescription = "H.264 in MP4(4:3)"
        mPresetParameters = "-r 29.97 -vcodec libx264 -s 640x480 -aspect 4:3 -flags +loop -cmp +chroma -deblockalpha 0 -deblockbeta 0 -b 1250k -maxrate 1500k -bufsize 4M -bt 256k -refs 1 -bf 3 -coder 1 -me_method umh -me_range 16 -subq 7 -partitions +parti4x4+parti8x8+partp8x8+partb8x8 -g 250 -keyint_min 25 -level 30 -qmin 10 -qmax 51 -qcomp 0.6 -trellis 2 -sc_threshold 40 -i_qfactor 0.71 -acodec libfaac -ab 112k -ar 48000 -ac 2"
        mPresetExtension = "mp4"
        mPresetCategory = "MP4"
    End If
    If lPreset = 58 Then
        mPresetDescription = "H.264 for PSP (4:3)"
        mPresetParameters = "-r 29.97 -vcodec libx264 -s 640x480 -aspect 4:3 -b 1250k -maxrate 4M -bufsize 4M -flags +loop -cmp +chroma -partitions +parti4x4+partp8x8+partb8x8 -me_method umh -subq 6 -trellis 1 -refs 2 -bf 1 -coder 1 -me_range 16 -g 300 -keyint_min 25 -sc_threshold 40 -i_qfactor 0.71 -bt 1250k -qcomp 0.6 -qmin 10 -qmax 51 -qdiff 4 -level 21 -acodec libfaac -ab 128k -ar 48000 -ac 2"
        mPresetExtension = "mp4"
        mPresetCategory = "PSP"
    End If
    If lPreset = 59 Then
        mPresetDescription = "H.264 in MP4 (16:9)"
        mPresetParameters = "-r 29.97 -vcodec libx264 -s 704x384 -aspect 16:9 -flags +loop -cmp +chroma -deblockalpha 0 -deblockbeta 0 -b 1250k -maxrate 1500k -bufsize 4M -bt 256k -refs 1 -bf 3 -coder 1 -me_method umh -me_range 16 -subq 7 -partitions +parti4x4+parti8x8+partp8x8+partb8x8 -g 250 -keyint_min 25 -level 30 -qmin 10 -qmax 51 -qcomp 0.6 -trellis 2 -sc_threshold 40 -i_qfactor 0.71 -acodec libfaac -ab 112k -ar 48000 -ac 2"
        mPresetExtension = "mp4"
        mPresetCategory = "MP4"
    End If
    If lPreset = 60 Then
        mPresetDescription = "CDMA phone 3g2"
        mPresetParameters = "-ar 22050 -ab 128k -acodec libfaac -s qcif -b 128kb -r 14.985"
        mPresetExtension = "3g2"
        mPresetCategory = "Mobile Phones"
    End If
    If lPreset = 61 Then
        mPresetDescription = "XviD for Zen (4:3)"
        mPresetParameters = "-r 29.97 -vcodec libxvid -vtag XVID -s 320x240 -aspect 4:3 -maxrate 1800k -b 1500k  -acodec libmp3lame -ar 48000 -ab 128k -ac 2"
        mPresetExtension = "avi"
        mPresetCategory = "Creative Zen"
    End If
    If lPreset = 62 Then
        mPresetDescription = "XviD for Zen (16:9)"
        mPresetParameters = "-r 29.97 -vcodec libxvid -vtag XVID -s 320x240 -aspect 16:9 -maxrate 1800k -b 1500k  -acodec libmp3lame -ar 48000 -ab 128k -ac 2"
        mPresetExtension = "avi"
        mPresetCategory = "Creative Zen"
    End If
    If lPreset = 63 Then
        mPresetDescription = "Blackberry Curve FS"
        mPresetParameters = "-vcodec mpeg4 -b 400k -r 24 -s 320x240 -aspect 4:3 -acodec libfaac -ar 22050 -ac 2 -ab 48k"
        mPresetExtension = "mp4"
        mPresetCategory = "Blackberry"
    End If
    If lPreset = 64 Then
        mPresetDescription = "Blackberry Curve WS"
        mPresetParameters = "-vcodec mpeg4 -b 400k -r 24 -s 320x180 -aspect 16:9 -acodec libfaac -ar 22050 -ac 2 -ab 48k"
        mPresetExtension = "mp4"
        mPresetCategory = "Blackberry"
    End If
    If lPreset = 65 Then
        mPresetDescription = "LG Chocolate"
        mPresetParameters = "-s 240x192 -r 11.988 -b 192 -ab 56 -vcodec libxvid -acodec libfaac"
        mPresetExtension = "mp4"
        mPresetCategory = "LG"
    End If
    If lPreset = 66 Then
        mPresetDescription = "Blackberry Music"
        mPresetParameters = "-acodec libmp3lame -ab 128k -ac 2 -ar 44100"
        mPresetExtension = "mp3"
        mPresetCategory = "Blackberry"
    End If
    If lPreset = 67 Then
        mPresetDescription = "Blackberry Video 16:9"
        mPresetParameters = "-r 24 -vcodec libx264 -s 320x180 -aspect 16:9 -maxrate 800k -bufsize 80k -b 400k -acodec libfaac -ar 44100 -ab 80k -ac 2"
        mPresetExtension = "mp4"
        mPresetCategory = "Blackberry"
    End If
    If lPreset = 68 Then
        mPresetDescription = "Blackberry Video 4:3"
        mPresetParameters = "-r 24 -vcodec libx264 -s 240x180 -aspect 4:3 -maxrate 800k -bufsize 80k -b 400k -acodec libfaac -ar 44100 -ab 80k -ac 2"
        mPresetExtension = "mp4"
        mPresetCategory = "Blackberry"
    End If
    If lPreset = 69 Then
        mPresetDescription = "Flash: Video (flv) for Web use (4:3)"
        mPresetParameters = "-vcodec flv -f flv -r 29.97 -s 320x240 -aspect 4:3 -b 300kb -g 160 -cmp dct  -subcmp dct  -mbd 2 -flags +aic+cbp+mv0+mv4 -trellis 1 -ac 1 -ar 22050 -ab 56k"
        mPresetExtension = "flv"
        mPresetCategory = "Websites"
    End If
    If lPreset = 70 Then
        mPresetDescription = "Flash: Video (flv) for Web use (16:9)"
        mPresetParameters = "-vcodec flv -f flv -r 29.97 -s 320x180 -aspect 16:9 -b 300kb -g 160 -cmp dct -subcmp dct -mbd 2 -flags +aic+cbp+mv0+mv4 -trellis 1 -ac 1 -ar 22050 -ab 56k"
        mPresetExtension = "flv"
        mPresetCategory = "Websites"
    End If
    If lPreset = 71 Then
        mPresetDescription = "H.264 for PSP (16:9)"
        mPresetParameters = "-r 29.97 -vcodec libx264 -s 480x272 -aspect 16:9 -b 1250k -maxrate 4M -bufsize 4M -flags +loop -cmp +chroma -partitions +parti4x4+partp8x8+partb8x8 -me_method umh -subq 6 -trellis 1 -refs 2 -bf 1 -coder 1 -me_range 16 -g 300 -keyint_min 25 -sc_threshold 40 -i_qfactor 0.71 -bt 1250k -qcomp 0.6 -qmin 10 -qmax 51 -qdiff 4 -level 21 -acodec libfaac -ab 128k -ar 48000 -ac 2"
        mPresetExtension = "mp4"
        mPresetCategory = "PSP"
    End If
    If lPreset = 72 Then
        mPresetDescription = "iPod Small Screen 4:3 CRF 21 (iTunes)"
        mPresetParameters = "-r 29.97 -vcodec libx264 -flags +loop -cmp +chroma -deblockalpha 0 -deblockbeta 0 -crf 21 -bt 256k -refs 1 -coder 0 -me_method full -me_range 16 -subq 5 -partitions +parti4x4+parti8x8+partp8x8 -g 250 -keyint_min 25 -level 30 -trellis 2 -sc_threshold 40 -i_qfactor 0.71 -s 320x240 -aspect 4:3 -acodec libfaac -ab 112k -ar 48000 -ac 2"
        mPresetExtension = "m4v"
        mPresetCategory = "iPod-iTunes"
    End If
    If lPreset = 73 Then
        mPresetDescription = "iPod Small Screen 16:9 CRF 21 (iTunes)"
        mPresetParameters = "-r 29.97 -vcodec libx264 -flags +loop -cmp +chroma -deblockalpha 0 -deblockbeta 0 -crf 21 -bt 256k -refs 1 -coder 0 -me_method full -me_range 16 -subq 5 -partitions +parti4x4+parti8x8+partp8x8 -g 250 -keyint_min 25 -level 30 -trellis 2 -sc_threshold 40 -i_qfactor 0.71 -s 320x176 -aspect 16:9 -acodec libfaac -ab 112k -ar 48000 -ac 2"
        mPresetExtension = "m4v"
        mPresetCategory = "iPod-iTunes"
    End If
    If lPreset = 74 Then
        mPresetDescription = "iPod Small Screen 16:9 Anamorphic CRF 21 (iTunes)"
        mPresetParameters = "-r 29.97 -croptop 60 -cropbottom 60 -vcodec libx264 -flags +loop -cmp +chroma -deblockalpha 0 -deblockbeta 0 -crf 21 -bt 256k -refs 1 -coder 0 -me_method full -me_range 16 -subq 5 -partitions +parti4x4+parti8x8+partp8x8 -g 250 -keyint_min 25 -level 30 -trellis 2 -sc_threshold 40 -i_qfactor 0.71 -s 320x144 -aspect 2.35 -acodec libfaac -ab 112k -ar 48000 -ac 2"
        mPresetExtension = "m4v"
        mPresetCategory = "iPod-iTunes"
    End If
    If lPreset = 75 Then
        mPresetDescription = "iPod Small Screen 4:3 to 16:9 CRF  21 (iTunes)"
        mPresetParameters = "-r 29.97 -croptop 64 -cropbottom 64 -vcodec libx264 -flags +loop -cmp +chroma -deblockalpha 0 -deblockbeta 0 -crf 21 -bt 256k -refs 1 -coder 0 -me_method full -me_range 16 -subq 5 -partitions +parti4x4+parti8x8+partp8x8 -g 250 -keyint_min 25 -level 30 -trellis 2 -sc_threshold 40 -i_qfactor 0.71 -s 320x176 -aspect 16:9 -acodec libfaac -ab 112k -ar 48000 -ac 2"
        mPresetExtension = "m4v"
        mPresetCategory = "iPod-iTunes"
    End If
    If lPreset = 76 Then
        mPresetDescription = "iPod TV Out 4:3 CRF 21 (iTunes)"
        mPresetParameters = "-r 29.97 -vcodec libx264 -flags +loop -cmp +chroma -deblockalpha 0 -deblockbeta 0 -crf 21 -bt 256k -refs 1 -coder 0 -me_method full -me_range 16 -subq 5 -partitions +parti4x4+parti8x8+partp8x8 -g 250 -keyint_min 25 -level 30 -trellis 2 -sc_threshold 40 -i_qfactor 0.71 -s 512x384 -aspect 4:3 -acodec libfaac -ab 112k -ar 48000 -ac 2"
        mPresetExtension = "m4v"
        mPresetCategory = "iPod-iTunes"
    End If
    If lPreset = 77 Then
        mPresetDescription = "iPod TV Out 16:9 CRF 21 (iTunes)"
        mPresetParameters = "-r 29.97 -vcodec libx264 -flags +loop -cmp +chroma -deblockalpha 0 -deblockbeta 0 -crf 21 -bt 256k -refs 1 -coder 0 -me_method full -me_range 16 -subq 5 -partitions +parti4x4+parti8x8+partp8x8 -g 250 -keyint_min 25 -level 30 -trellis 2 -sc_threshold 40 -i_qfactor 0.71 -s 624x352 -aspect 16:9 -acodec libfaac -ab 112k -ar 48000 -ac 2"
        mPresetExtension = "m4v"
        mPresetCategory = "iPod-iTunes"
    End If
    If lPreset = 78 Then
        mPresetDescription = "iPod TV Out 16:9 Anamorphic CRF 21 (iTunes)"
        mPresetParameters = "-r 29.97 -croptop 60 -cropbottom 60 -vcodec libx264 -flags +loop -cmp +chroma -deblockalpha 0 -deblockbeta 0 -crf 21 -bt 256k -refs 1 -coder 0 -me_method full -me_range 16 -subq 5 -partitions +parti4x4+parti8x8+partp8x8 -g 250 -keyint_min 25 -level 30 -trellis 2 -sc_threshold 40 -i_qfactor 0.71 -s 640x272 -aspect 2.35 -acodec libfaac -ab 112k -ar 48000 -ac 2"
        mPresetExtension = "m4v"
        mPresetCategory = "iPod-iTunes"
    End If
    If lPreset = 79 Then
        mPresetDescription = "iPod TV Out 4:3 to 16:9 CRF  21 (iTunes)"
        mPresetParameters = "-r 29.97 -croptop 64 -cropbottom 64 -vcodec libx264 -flags +loop -cmp +chroma -deblockalpha 0 -deblockbeta 0 -crf 21 -bt 256k -refs 1 -coder 0 -me_method full -me_range 16 -subq 5 -partitions +parti4x4+parti8x8+partp8x8 -g 250 -keyint_min 25 -level 30 -trellis 2 -sc_threshold 40 -i_qfactor 0.71 -s 640x352 -aspect 16:9 -acodec libfaac -ab 112k -ar 48000 -ac 2"
        mPresetExtension = "m4v"
        mPresetCategory = "iPod-iTunes"
    End If
    If lPreset = 80 Then
        mPresetDescription = "XviD for iPod (4:3)"
        mPresetParameters = "-r 29.97 -vcodec libxvid -s 640x480 -aspect 4:3 -maxrate 1500k -b 1250k -qmin 3 -qmax 5 -bufsize 4096 -mbd 2 -flags +4mv -aic 2 -cmp 2 -subcmp 2 -trellis 1 -g 300 -acodec libfaac -ar 48000 -ab 80k -ac 2"
        mPresetExtension = "mp4"
        mPresetCategory = "Ipod"
    End If
    If lPreset = 81 Then
        mPresetDescription = "XviD for iPod (16:9)"
        mPresetParameters = "-r 29.97 -vcodec libxvid -s 704x384 -aspect 16:9 -maxrate 1500k -b 1250k -qmin 3 -qmax 5 -bufsize 4096 -mbd 2 -flags +4mv -aic 2 -cmp 2 -subcmp 2 -trellis 1 -g 300 -acodec libfaac -ar 48000 -ab 80k -ac 2"
        mPresetExtension = "mp4"
        mPresetCategory = "Ipod"
    End If
    If lPreset = 82 Then
        mPresetDescription = "XviD for iPod (16:9 Anamorphic)"
        mPresetParameters = "-r 29.97 -croptop 60 -cropbottom 60 -vcodec libxvid -s 640x272 -aspect 2.35 -maxrate 1500k -b 1250k -qmin 3 -qmax 5 -bufsize 4096 -mbd 2 -flags +4mv -aic 2 -cmp 2 -subcmp 2 -trellis 1 -g 300 -acodec libfaac -ar 48000 -ab 80k -ac 2"
        mPresetExtension = "mp4"
        mPresetCategory = "Ipod"
    End If
    If lPreset = 83 Then
        mPresetDescription = "OGG-64-kbps"
        mPresetParameters = "-acodec vorbis -aq 13 -vn"
        mPresetExtension = "OGG"
        mPresetCategory = "Audio"
    End If
    If lPreset = 84 Then
        mPresetDescription = "OGG-96-kbps"
        mPresetParameters = "-acodec vorbis -aq 23 -vn"
        mPresetExtension = "OGG"
        mPresetCategory = "Audio"
    End If
    If lPreset = 85 Then
        mPresetDescription = "OGG-128-kbps"
        mPresetParameters = "-acodec vorbis -aq 39 -vn"
        mPresetExtension = "OGG"
        mPresetCategory = "Audio"
    End If
    If lPreset = 86 Then
        mPresetDescription = "OGG-160-kbps"
        mPresetParameters = "-acodec vorbis -aq 60 -vn"
        mPresetExtension = "OGG"
        mPresetCategory = "Audio"
    End If
    If lPreset = 87 Then
        mPresetDescription = "WMA-32k-mono-22050"
        mPresetParameters = "-vn  -acodec wmav2 -ab 32k -ac 1 -ar 22050"
        mPresetExtension = "wma"
        mPresetCategory = "Audio"
    End If
    If lPreset = 88 Then
        mPresetDescription = "WMV-4:3-320-x240-757kbps-32k-mono"
        mPresetParameters = "-vcodec wmv2  -acodec wmav2 -aspect 4:3 -b 757k -ab 32k -ac 1 -ar 22050 -s 320x240"
        mPresetExtension = "wmv"
        mPresetCategory = "WMV"
    End If
    If lPreset = 89 Then
        mPresetDescription = "WMV-4:3-320-x240-500kbps-32k-mono"
        mPresetParameters = "-vcodec wmv2  -acodec wmav2 -aspect 4:3 -b 500k -ab 32k -ac 1 -ar 22050 -s 320x240"
        mPresetExtension = "wmv"
        mPresetCategory = "WMV"
    End If
    If lPreset = 90 Then
        mPresetDescription = "WMA-64k-stereo-22050"
        mPresetParameters = "-vn  -acodec wmav2 -ab 64k -ac 2 -ar 22050"
        mPresetExtension = "wma"
        mPresetCategory = "Audio"
    End If
    If lPreset = 91 Then
        mPresetDescription = "WMV-16:9-320-x240-757kbps-32k-mono"
        mPresetParameters = "-vcodec wmv2  -acodec wmav2 -aspect 16:9 -b 757k -ab 32k -ac 1 -ar 22050 -s 320x240"
        mPresetExtension = "wmv"
        mPresetCategory = "WMV"
    End If
    If lPreset = 92 Then
        mPresetDescription = "WMV-4:3-320-x240-200kbps-32k-mono"
        mPresetParameters = "-vcodec wmv2  -acodec wmav2 -aspect 4:3 -b 200k -ab 32k -ac 1 -ar 22050 -s 320x240"
        mPresetExtension = "wmv"
        mPresetCategory = "WMV"
    End If
    If lPreset = 93 Then
        mPresetDescription = "WMV-4:3-320-x240-300kbps-16k-mono"
        mPresetParameters = "-vcodec wmv2  -acodec wmav2 -aspect 4:3 -b 300k -ab 16k -ac 1 -ar 22050 -s 320x240"
        mPresetExtension = "wmv"
        mPresetCategory = "WMV"
    End If
End Function

' ----------------------------------------------------------------------
' -- Función de ajuste para usar con PathFindExtension
' ----------------------------------------------------------------------
Function GetExtension(ByVal sPath As String) As String
  Dim sTemp As String
  'Given a path and filename, return only the filename extension.
  sTemp = pvGetStrFromPtrA(PathFindExtension(sPath))
     
  If Len(sTemp) Then GetExtension = Replace(sTemp, ".", vbNullString)
End Function
  
' ----------------------------------------------------------------------------------
' \\ -- Devuelve un string desde un puntero
' ----------------------------------------------------------------------------------
Private Function pvGetStrFromPtrA(ByVal lpszA As Long) As String
  'Given a pointer to a string, return the string
   pvGetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal pvGetStrFromPtrA, ByVal lpszA)
End Function

