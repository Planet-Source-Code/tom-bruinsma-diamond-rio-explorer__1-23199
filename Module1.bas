Attribute VB_Name = "Module1"
Global Const LISTVIEW_MODE0 = "View Large Icons"
Global Const LISTVIEW_MODE1 = "View Small Icons"
Global Const LISTVIEW_MODE2 = "View List"
Global Const LISTVIEW_MODE3 = "View Details"
Public fMainForm As frmMain
Dim bitrate_lookup(7, 15) As Integer
Public actual_bitrate As Long


Public Declare Function SetWindowPos Lib "user32" _
    (ByVal hwnd As Long, _
    ByVal hWndInsertAfter As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal cx As Long, _
    ByVal cy As Long, _
    ByVal wFlags As Long) As Long

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpbuffer As String, ByVal nSize As Long) As Long
Public syspath As String


Sub Main()
    Set fMainForm = New frmMain

    Load frmLogin
    If frmLogin.ActiveLock1.RegisteredUser Then
        fMainForm.Show
        InitializeMp3Lookup

        Unload frmLogin
    Else
        frmLogin.Show
    End If
End Sub

Function WindowsDirectory() As String 'get windows directory
Dim WinPath As String
Dim temp
    WinPath = String(145, Chr(0))
    temp = GetWindowsDirectory(WinPath, 145)
    WindowsDirectory = Left(WinPath, InStr(WinPath, Chr(0)) - 1)
syspath = WindowsDirectory
End Function
   
Public Function Getmp3data(MP3File As String) As String
    Dim f As Object, fs As Object
     Dim dIN As String
     cr = Chr(10)
     Open MP3File For Binary As #21
     ' read in 1st 4k of .mp3 file to find a frame header
     dIN = Input(4096, #21)
     filesize = LOF(21) ' needed to calculate track duration
     Close #21
     
     ' frame header starts with 12 set bits [sync]
     ' NB this ignores MPEG-2.5 which is 11 set bits, 1 zero bit.
     
     ' my search for the sync bits only works on nibble boundaries,
     ' I'm not sure if it is necessary to search on bit boundaries -
     ' if so then this search will be 4* slower and require a rewrite
     ' of this search section and shift_those_bits.
     Do Until i = 4095
       i = i + 1
       d1 = Asc(Mid(dIN, i, 1))
       d2 = Asc(Mid(dIN, i + 1, 1))
       If d1 = &HFF And (d2 And &HF0) = &HF0 Then
         ''debug.print "Found at"; i
         ' get 20 hdr bits - they are last 20 bits of next 3 bytes
         temp_string = Mid(dIN, i + 1, 3)
         mp3bits_string = shift_those_bits(Mid(dIN, i + 1, 3))
         Exit Do
       End If
       ' if we haven't found the sync yet then shift left by 4 bits
       dSHIFT = shift_those_bits(Mid(dIN, i, 3))
       dd1 = Asc(Left(dSHIFT, 1))
       dd2 = Asc(Right(dSHIFT, 1))
       If dd1 = &HFF And (dd2 And &HF0) = &HF0 Then
         ''debug.print "Found at"; i; "& a nibble"
         ' get 20 hdr bits - they are first 20 bits of next 3 bytes
         mp3bits_string = Mid(dIN, i + 2, 3)
         Exit Do
       End If
     Loop
     
     ' 1st 20 bits of mp3bits_string are hdr info for this frame
     ' 1st bit is ID - 0=MPG-2, 1=MPG-1
     mp3_id = (&H80 And Asc(Left(mp3bits_string, 1))) / 128
     ' next 2 bits are Layer
     mp3_layer = (&H60 And Asc(Left(mp3bits_string, 1))) / 32
     ' next bit is Protection
     mp3_prot = &H10 And Asc(Left(mp3bits_string, 1))
     ' next 4 bits are bitrate
     mp3_bitrate = &HF And Asc(Left(mp3bits_string, 1))
     'next 2 bits are frequency
     mp3_freq = &HC0 And Asc(Mid(mp3bits_string, 2, 1))
     ' next bit is Padding
     mp3_pad = (&H20 And Asc(Mid(mp3bits_string, 2, 1))) / 2
     actual_bitrate = 1000 * CLng((bitrate_lookup((mp3_id * 4) Or mp3_layer, mp3_bitrate)))



     dat = Str(actual_bitrate) & vbTab
      
      Select Case (mp3_id * 4) Or mp3_freq
        Case 0
          sample_rate = 22050
        Case 1
          sample_rate = 24000
        Case 2
          sample_rate = 16000
        Case 4
          sample_rate = 44100
        Case 5
          sample_rate = 48000
        Case 6
          sample_rate = 32000
      End Select
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(MP3File)

         

    
    dat = Format(f.Size, "###,###,###") & vbTab
    dat = dat & Str(actual_bitrate / 1000) & vbTab
    dat = dat & Str(sample_rate) & vbTab
    dat = dat & Format(Date & " " & Time(), "m/d/yy h:mm AMPM")
      ' calculate track time

Getmp3data = dat

   End Function
   
   Public Function shift_those_bits(dIN As String) As String
     ' need to left shift 4 bits losing most significant 4 bits
     Dim sd1, sd2, sd3, do1, do2 As Integer
     duff = Left(dIN, 1)
     duff2 = Asc(duff)
     sd1 = Asc(Left(dIN, 1))
     sd2 = Asc(Mid(dIN, 2, 1))
     sd3 = Asc(Right(dIN, 1))
     
     do1 = ((sd1 And &HF) * 16) Or ((sd2 And &HF0) / 16)
     do2 = ((sd2 And &HF) * 16) Or ((sd3 And &HF0) / 16)
     shift_those_bits = Chr(do1) + Chr(do2)
   End Function



Public Sub InitializeMp3Lookup()
  
  ' setup array for mpeg bitrate info
  bitrate_data = "032,032,032,032,008,008,"
  bitrate_data = bitrate_data + "064,048,040,048,016,016,"
  bitrate_data = bitrate_data + "096,056,048,056,024,024,"
  bitrate_data = bitrate_data + "128,064,056,064,032,032,"
  bitrate_data = bitrate_data + "160,080,064,080,040,040,"
  bitrate_data = bitrate_data + "192,096,080,096,048,048,"
  bitrate_data = bitrate_data + "224,112,096,112,056,056,"
  bitrate_data = bitrate_data + "256,128,112,128,064,064,"
  bitrate_data = bitrate_data + "288,160,128,144,080,080,"
  bitrate_data = bitrate_data + "320,192,160,160,096,096,"
  bitrate_data = bitrate_data + "352,224,192,176,112,112,"
  bitrate_data = bitrate_data + "384,256,224,192,128,128,"
  bitrate_data = bitrate_data + "416,320,256,224,144,144,"
  bitrate_data = bitrate_data + "448,384,320,256,160,160,"
    
  For Y = 1 To 14
    For X = 7 To 5 Step -1
      bitrate_lookup(X, Y) = Left(bitrate_data, 3)
      bitrate_data = Right(bitrate_data, Len(bitrate_data) - 4)
    Next
    For X = 3 To 1 Step -1
      bitrate_lookup(X, Y) = Left(bitrate_data, 3)
      bitrate_data = Right(bitrate_data, Len(bitrate_data) - 4)
    Next
  Next
End Sub
