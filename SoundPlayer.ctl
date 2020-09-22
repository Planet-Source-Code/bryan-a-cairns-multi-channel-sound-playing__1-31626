VERSION 5.00
Begin VB.UserControl SoundPlayer 
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   630
   InvisibleAtRuntime=   -1  'True
   Picture         =   "SoundPlayer.ctx":0000
   ScaleHeight     =   525
   ScaleWidth      =   630
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   0
   End
End
Attribute VB_Name = "SoundPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Multi Channel Sound Player
'By Bryan Cairns
'Plays...
'qt, mov, dat, snd, mpg, mpa, mpv, enc, m1v, mp2, mp3, mpe, mpeg, mpm, au, snd, aif, aiff, aifc, wav, wmv, wma, avi, midi, mid, rmi, avi


Private Type TheSound
    LID As Long
    bLoop As Boolean
End Type
Dim SoundFiles() As TheSound

Public Function PlaySound(sFilename As String, bLoop As Boolean) As Long
On Error GoTo EH
'Returns the INDEX in our collection
Dim AliasName As String
Dim typeDevice As String
Dim Result As String
Dim INUM As Long

INUM = GetFreeID(bLoop)
typeDevice = "MPEGVideo"
'MPEGVideo is a generic type in windows - this means it will try to play...
'qt , mov, dat,snd, mpg, mpa, mpv, enc, m1v, mp2,mp3, mpe, mpeg, mpm
'au , snd, aif, aiff, aifc,wav,wmv,wma,avi,midi,mid,rmi,avi

AliasName = "movie" & INUM
Result = OpenMultimedia(UserControl.hWnd, AliasName, sFilename, typeDevice)
StartSound INUM
If Timer1.Enabled = False And bLoop = True Then Timer1.Enabled = True
PlaySound = INUM
Exit Function
EH:
MsgBox Err.Description, vbCritical, "Sound Player"
Exit Function
End Function

Private Sub StartSound(I As Long)
'Start Playing a Sound
On Error GoTo EH
Dim AliasName As String
Dim Result As String

AliasName = "movie" & I
Result = PlayMultimedia(AliasName, "", "")
Exit Sub
EH:
MsgBox Err.Description, vbCritical, "Sound Player"
Exit Sub
End Sub

Public Sub StopSound(I As Long)
'Stop one sound
On Error GoTo EH
Dim AliasName As String
Dim Result As String
Dim H As Integer
AliasName = "movie" & I

Result = StopMultimedia(AliasName)
Result = CloseMultimedia(AliasName)

Timer1.Enabled = False
DoEvents
For H = LBound(SoundFiles) To UBound(SoundFiles)
    If SoundFiles(H).LID = I Then
        SoundFiles(H).LID = 0
        SoundFiles(H).bLoop = False
    End If
Next H
DoEvents
If ISAnySoundOpen = True Then
    Timer1.Enabled = True
End If
Exit Sub
EH:
MsgBox Err.Description, vbCritical, "Sound Player"
Exit Sub
End Sub

Private Function ISAnySoundOpen() As Boolean
'See if we have any sound files playing
On Error GoTo EH
Dim I As Long
Dim bFound As Boolean

bFound = False
    For I = LBound(SoundFiles) To UBound(SoundFiles)
        If SoundFiles(I).LID <> 0 Then
            bFound = True
        End If
    Next I
ISAnySoundOpen = bFound
Exit Function
EH:
MsgBox Err.Description, vbCritical, "Sound Player"
Exit Function
End Function

Public Sub StopAll()
'Stop ALL sounds that are playing
On Error GoTo EH
    Dim I As Long
    For I = LBound(SoundFiles) To UBound(SoundFiles)
        StopSound SoundFiles(I).LID
    Next I
Timer1.Enabled = False
DoEvents

ReDim Preserve SoundFiles(1)
I = UBound(SoundFiles)
SoundFiles(I).bLoop = False
SoundFiles(I).LID = 0
Exit Sub
EH:
MsgBox Err.Description, vbCritical, "Sound Player"
Exit Sub
End Sub

Private Function GetFreeID(bLoop As Boolean) As Long
On Error GoTo EH
Dim I As Integer
Dim bFound As Boolean
For I = LBound(SoundFiles) To UBound(SoundFiles)
If SoundFiles(I).LID = 0 Then
    bFound = True
    SoundFiles(I).LID = GetInternalID
    SoundFiles(I).bLoop = bLoop
    GetFreeID = SoundFiles(I).LID
    Exit Function
End If
Next I
If bFound = False Then
I = UBound(SoundFiles) + 1
ReDim Preserve SoundFiles(I)
    SoundFiles(I).LID = GetInternalID
    SoundFiles(I).bLoop = bLoop
    GetFreeID = SoundFiles(I).LID
End If
Exit Function
EH:
MsgBox Err.Description, vbCritical, "Sound Player"
Exit Function
End Function

Private Function GetInternalID() As Integer
On Error GoTo EH
Dim I As Integer
Dim H As Integer
Dim bFound As Boolean
For H = 1 To 1000
bFound = False
    For I = LBound(SoundFiles) To UBound(SoundFiles)
    If SoundFiles(I).LID = H Then bFound = True
    Next I
    If bFound = False Then
        GetInternalID = H
        Exit Function
    End If
Next H
GetInternalID = 1000
Exit Function
EH:
MsgBox Err.Description, vbCritical, "Sound Player"
Exit Function
End Function


Private Sub Timer1_Timer()
On Error GoTo EH
Dim I As Long
Dim AliasName As String
Dim sVal As String

For I = LBound(SoundFiles) To UBound(SoundFiles)


DoEvents
AliasName = "movie" & SoundFiles(I).LID
'check for sound files we have to loop
    If SoundFiles(I).bLoop = True Then
           If LCase(GetStatusMultimedia(AliasName)) = "stopped" Then 'If AreMultimediaAtEnd(AliasName, Val("")) = True Then  ' alias name for e.g.:"movie"
                StartSound SoundFiles(I).LID
            End If
    End If
'clear old sound files as they end
    If SoundFiles(I).LID <> 0 And SoundFiles(I).bLoop = False Then
            If LCase(GetStatusMultimedia(AliasName)) = "stopped" Then  ' alias name for e.g.:"movie"
               StopSound SoundFiles(I).LID
            End If
    End If
Next I
Exit Sub
EH:
Timer1.Enabled = False
StopAll
MsgBox Err.Description, vbCritical, "Sound Player"
Exit Sub
End Sub

Private Sub UserControl_Initialize()
On Error GoTo EH
Dim I As Integer
ReDim Preserve SoundFiles(1)
I = UBound(SoundFiles)
SoundFiles(I).bLoop = False
SoundFiles(I).LID = 0
Exit Sub
EH:
MsgBox Err.Description, vbCritical, "Sound Player"
Exit Sub
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
UserControl.Width = 495
UserControl.Height = 495
End Sub

Private Sub UserControl_Terminate()
StopAll
End Sub
