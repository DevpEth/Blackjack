Attribute VB_Name = "Module1"
'SOUND
'''''''''''''''''''''''''''''''''''''''''
Public Const SND_ASYNC As Long = &H1            '  play asynchronously
Public Const SND_FILENAME As Long = &H20000     '  name is a file name
Public Const SND_NODEFAULT As Long = &H2        '  silence not default, if sound not found
Public Const SND_NOWAIT As Long = &H2000        '  don't wait if the driver is busy
Public Const SND_RESOURCE As Long = &H40004     '  name is a resource name or atom
Public Const SND_MEMORY = &H4                   ' lpszSoundName points to a memory file

Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function sndPlaySound Lib "winmm" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
'''''''''''''''''''''''''''''''''''''''''

Public gCheckStart As Boolean
Public Moni As Long
Public NumOfDecks As Integer
Public MoniStartC As Long

'play sound
'SoundBuffer = StrConv(LoadResData("x", "SOUND"), vbUnicode)
'retVal = sndPlaySound(SoundBuffer, SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY)
'Environ("COMPUTERNAME")
