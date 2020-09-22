Attribute VB_Name = "ModWav"
'Wav Play API & Consts
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" _
(ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10

'==================================================================

'==========Copyright Notice
'Adopted module from the Web. Information below
'==========

'Write sound data to wave files
'By Axel Brink
'21-May-2002
'Write comments to axel@fmf.nl
'You can use this file freely, at your own risk
'Credits would be appreciated, but are not necessary
'Please do not redistribute this source code without my name
'Wavefile format information used from http://ccrma-www.stanford.edu/CCRMA/Courses/422:1998/projects/WaveFormat/
'Some comments from this address are copied here (marked ##)
'
'Example use:
'
'Private Sub cmdHighQuality_Click()
'  Const Hertz As Single = 2 * 3.14159265 / 44100 'use sin(sam*hertz) to get a tone of 1 hertz
'  Dim DataBytes(0 To 1, 0 To 44099) As Long
'  Dim sam As Long
'  Dim chan As Integer
'
'  MousePointer = vbHourglass
'  For sam = 0 To 44099
'    For chan = 0 To 1
'      DataBytes(0, sam) = 4000 * Sin(500 * sam * Hertz)
'      DataBytes(1, sam) = 4000 * Sin(300 * sam * Hertz)
'    Next chan
'  Next sam
'
'  WriteWave txtOutputPath.Text, CreateWaveArray(DataBytes, 44100, 16)
'  MousePointer = vbDefault
'End Sub
'
'Private Sub cmdLowQuality_Click()
'  Const Hertz As Single = 2 * 3.14159265 / 22050 'use sin(sam*hertz) to get a tone of 1 hertz
'  Dim DataBytes(0 To 0, 0 To 22049) As Long
'  Dim sam As Long
'
'  MousePointer = vbHourglass
'  For sam = 0 To 22049
'    DataBytes(0, sam) = 127 + 40 * Sin(800 * sam * Hertz) + 40 * Sin(300 * sam * Hertz)
'  Next sam
'
'  WriteWave txtOutputPath.Text, CreateWaveArray(DataBytes, 22050, 8)
'  MousePointer = vbDefault
'End Sub


Option Explicit
Public Const WAVE_MIN = -32768 'Extreme sample values for 16-bit samples
Public Const WAVE_MAX = 32767
  

Public Function CreateWaveArray(WaveData() As Long, SampleRate As Long, BitsPerSample As Long) As Byte()
'Reads sample data and returns an array with bytes which can be
'written to a .wav file
'Uses PCM (i.e. no compression)

'WaveData:  * Contents:  Sample values of type Long
'           * Structure: An array(channel, samplenumber)
'           * Channels:  The number of elements in the first dimension
'                        determines the number of channels
'           * Samples:   The number of elements in the second dimension
'                        determines the number of samples
'                        8-bit samples must be between 0 and 255
'                        16-bit samples must be between -32768 and 32767
'           Note that indexing is supposed to start at 0.
'SampleRate:             Number of samples per second. Typically 44100.
'BitsPerSample:          Number of bits per sample. 8 or 16.

  Const Byte1Mask As Long = 255      'Long integers are stored little endian.
  Const Byte2Mask As Long = 65280    'These bit masks select bytes from a number to store.
  Const Byte3Mask As Long = 16711680
  Const Byte2Divisor As Long = 256
  Const Byte3Divisor As Long = 65536
  Const Byte4Divisor As Long = 16777216
  
  Dim NumSamples As Long
  Dim NumChannels As Long
  Dim FileSize As Long
  Dim ByteRate As Long
  Dim BlockAlign As Long '## "The number of bytes for one sample including all channels.
                         '    I wonder what happens when this number isn't an integer?"
  Dim OutputFile() As Byte
  Dim ChunkSize As Long  '## "36 + SubChunk2Size, or more precisely:
'                               4 + (8 + SubChunk1Size) + (8 + SubChunk2Size)
'                               This is the size of the rest of the chunk
'                               following this number.  This is the size of the
'                               entire file in bytes minus 8 bytes for the
'                               two fields not included in this count:
'                               ChunkID and ChunkSize."
  Dim Subchunk2Size As Long '## "== NumSamples * NumChannels * BitsPerSample/8
'                               This is the number of bytes in the data.
'                               You can also think of this as the size
'                               of the read of the subchunk following this
'                               number."
  Dim SampleNr As Long
  Dim ChannelNr As Long

  If BitsPerSample <> 8 And BitsPerSample <> 16 Then Stop

  NumChannels = UBound(WaveData, 1) + 1
  NumSamples = UBound(WaveData, 2) + 1

  ByteRate = SampleRate * NumChannels * BitsPerSample / 8
  BlockAlign = NumChannels * BitsPerSample / 8

  FileSize = 44 + NumSamples * BlockAlign
  ChunkSize = FileSize - 8
  Subchunk2Size = NumSamples * NumChannels * BitsPerSample / 8

  ReDim OutputFile(0 To FileSize - 1) As Byte
  
  '** ChunkID **
  OutputFile(0) = Asc("R")
  OutputFile(1) = Asc("I")
  OutputFile(2) = Asc("F")
  OutputFile(3) = Asc("F")
  
  '** ChunkSize **
  OutputFile(4) = ChunkSize And Byte1Mask 'Little endian
  OutputFile(5) = (ChunkSize And Byte2Mask) \ Byte2Divisor
  OutputFile(6) = (ChunkSize And Byte3Mask) \ Byte3Divisor
  OutputFile(7) = ChunkSize \ Byte4Divisor
  
  '** Format **
  OutputFile(8) = Asc("W")
  OutputFile(9) = Asc("A")
  OutputFile(10) = Asc("V")
  OutputFile(11) = Asc("E")
  
  '** Subchunk1ID **
  OutputFile(12) = Asc("f")
  OutputFile(13) = Asc("m")
  OutputFile(14) = Asc("t")
  OutputFile(15) = Asc(" ")
  
  '** Subchunk1Size ** (16 for PCM)
  OutputFile(16) = 16 And Byte1Mask
  OutputFile(17) = (16 And Byte2Mask) \ Byte2Divisor
  OutputFile(18) = (16 And Byte3Mask) \ Byte3Divisor
  OutputFile(19) = 16 \ Byte4Divisor
  
  '** AudioFormat ** (1 for PCM)
  OutputFile(20) = 1 And Byte1Mask
  OutputFile(21) = (1 And Byte2Mask) \ Byte2Divisor
  
  '** NumChannels **
  OutputFile(22) = NumChannels And Byte1Mask
  OutputFile(23) = (NumChannels And Byte2Mask) \ Byte2Divisor

  '** SampleRate **
  OutputFile(24) = SampleRate And Byte1Mask
  OutputFile(25) = (SampleRate And Byte2Mask) \ Byte2Divisor
  OutputFile(26) = (SampleRate And Byte3Mask) \ Byte3Divisor
  OutputFile(27) = SampleRate \ Byte4Divisor
  
  '** ByteRate **
  OutputFile(28) = ByteRate And Byte1Mask
  OutputFile(29) = (ByteRate And Byte2Mask) \ Byte2Divisor
  OutputFile(30) = (ByteRate And Byte3Mask) \ Byte3Divisor
  OutputFile(31) = ByteRate \ Byte4Divisor
  
  '** BlockAlign **
  OutputFile(32) = BlockAlign And Byte1Mask
  OutputFile(33) = (BlockAlign And Byte2Mask) \ Byte2Divisor
  
  '** BitsPerSample **
  OutputFile(34) = BitsPerSample And Byte1Mask
  OutputFile(35) = (BitsPerSample And Byte2Mask) \ Byte2Divisor
  
  '** Subchunk2ID **
  OutputFile(36) = Asc("d")
  OutputFile(37) = Asc("a")
  OutputFile(38) = Asc("t")
  OutputFile(39) = Asc("a")
  
  '** Subchunk2Size **
  OutputFile(40) = Subchunk2Size And Byte1Mask
  OutputFile(41) = (Subchunk2Size And Byte2Mask) \ Byte2Divisor
  OutputFile(42) = (Subchunk2Size And Byte3Mask) \ Byte3Divisor
  OutputFile(43) = Subchunk2Size \ Byte4Divisor
  
  If BitsPerSample = 8 Then 'Samples are unsigned bytes; from 0 to 255
    For SampleNr = 0 To NumSamples - 1
      For ChannelNr = 0 To NumChannels - 1
        OutputFile(44 + BlockAlign * SampleNr + ChannelNr) = WaveData(ChannelNr, SampleNr) And Byte1Mask
      Next ChannelNr
    Next SampleNr
  ElseIf BitsPerSample = 16 Then 'Samples are 2's complement signed bytes; from -32768 to 32767
    For SampleNr = 0 To NumSamples - 1
      For ChannelNr = 0 To NumChannels - 1
        OutputFile(44 + BlockAlign * SampleNr + ChannelNr * BitsPerSample \ 8) = WaveData(ChannelNr, SampleNr) And Byte1Mask
        OutputFile(45 + BlockAlign * SampleNr + ChannelNr * BitsPerSample \ 8) = (WaveData(ChannelNr, SampleNr) And Byte2Mask) \ Byte2Divisor
      Next ChannelNr
    Next SampleNr
  End If
  CreateWaveArray = OutputFile
  
  Exit Function
End Function

Public Sub WriteWave(Path As String, WaveArray() As Byte)
  Dim ErrorNumber As Long
  
  On Error GoTo WriteWaveError
  Open Path For Input As #1
  Close #1
  
  If ErrorNumber = 0 Then 'File exists
    Kill Path             'Delete
  End If
  
  Open Path For Binary As #1
  Put #1, , WaveArray
  Close #1
  
WriteWaveError:
  ErrorNumber = Err.Number
  Resume Next
End Sub


