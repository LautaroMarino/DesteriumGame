VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAudio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'**************************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'**************************************************************************

Option Explicit

Implements wGL_Audio_Callback ' Implements to allow reacting when music is over

Private Const CHANNEL_COUNT     As Long = 3
Private Const CHANNEL_MUSIC     As Long = 0
Private Const CHANNEL_EFFECT    As Long = 1
Private Const CHANNEL_UI        As Long = 2

Private Const EFFECT_FADE_TIME As Long = 1750 ' It Takes 1.75sec to fade in/out between music


Private mMasterMuted   As Boolean
Private mMasterVolume  As Long

Private mMasterEnabled As Boolean
Private mMusicEnabled  As Boolean
Private mMusicPath     As String
Private mMusicBuffer   As Long
Private mMusicID       As Long
Private mMusicFilename As String
Private mMusicList()   As Long

Private mFadeBuffer    As Long
Private mFadeID        As Long
Private mFadeFilename  As String
Private mFadeTick      As Long

Private mEffectEnabled As Boolean
Private mEffectPath    As String
Private mEffectVolume  As Long
Private mEffectBank    As Collection
Private mEffect3D      As Boolean

Private mInterfaceEnabled As Boolean
Private mInterfaceVolume  As Long

Public Function Initialize(ByVal MusicPath As String, ByVal EffectPath As String) As Boolean

On Error GoTo clsAudio_Initialize_Error

    Call wGL_Audio.Initialise(CHANNEL_COUNT, Me)
    
    ' Reset reverb effect
    Call wGL_Audio.SetReverb(REVERB_Default)

    'Set sound & music path
    mMusicPath = MusicPath
    mEffectPath = EffectPath
    
    mEffect3D = True
    
    ' Empty playlist
    ReDim mMusicList(0) As Long
    
    Set mEffectBank = New Collection
    
    Initialize = True
    
    Exit Function
    
clsAudio_Initialize_Error:
    Call LogError("Error" & err.Number & "(" & err.Description & ") en Initialize de clsAudio.cls")
    
End Function

Public Sub Update(ByVal Tick As Long)
        
#If SonidoPC = 0 Then
    ' Handle Mute/Unmute on background
    If (IsAppActive()) Then
        If (mMasterMuted) Then Call Unmute
    Else
        If (Not mMasterMuted) Then Call Mute
    End If
#End If
    
    ' Handle Fade In/Out of the music channel
    If (mFadeID <> 0) Then
        Call OnFadeEffect(Tick - mFadeTick)
    End If

    Call wGL_Audio.Update
    
End Sub

'Public Sub SetReverb(ByVal Reverb As wGL_Audio_Reverb)
    
    'Call wGL_Audio.SetReverb(IIf(Reverb = REVERB_Off, REVERB_Default, Reverb))  ' There is a glitch with Off
    
'End Sub

Public Sub Mute()
    
    Call wGL_Audio.SetMasterVolume(0)
    
    mMasterMuted = True
    
End Sub

Public Sub Unmute()
    
    Call wGL_Audio.SetMasterVolume(mMasterVolume * 0.01)

    mMasterMuted = False

End Sub

Public Sub SetListener(ByVal X As Single, ByVal Y As Single)

    Call wGL_Audio.SetMasterListener(X, Y, 0#)

End Sub

Public Sub Track(ByVal ID As Long)
    
    Call wGL_Audio.Track(ID)
    
End Sub

Public Function CreateSource(ByVal X As Single, ByVal Y As Single) As Long

    CreateSource = wGL_Audio.CreateSource
    
    If (CreateSource <> 0) Then
        Call UpdateSource(CreateSource, X, Y)
    End If
    
End Function

Public Sub UpdateSource(ByVal ID As Long, ByVal X As Single, ByVal Y As Single)

    If (ID <> 0) Then
        Call wGL_Audio.UpdateSource(ID, CreateDefaultEmitter(X, Y))
    End If

End Sub

Public Sub DeleteSource(ByRef ID As Long, ByVal StopImmediately As Boolean)
    
    If (ID <> 0) Then
        Call wGL_Audio.StopBySource(ID, StopImmediately)
        Call wGL_Audio.DeleteSource(ID)
    End If

    ID = 0
        
End Sub

Public Sub SetPlaylist(ByRef Playlist() As Long)
    
    mMusicList = Playlist
    
End Sub

Public Sub Halt()

    Call wGL_Audio.Halt(CHANNEL_EFFECT)
    
End Sub

Public Sub PlayMusic(ByVal FileName As String, Optional ByVal Repeat As Boolean = False, Optional ByVal Fade As Boolean = False)
   
    ' Check if music is enabled, or if parameters are valid
    If (Not mMusicEnabled Or FileName = vbNullString Or Not FileExist(mMusicPath & FileName, vbNormal)) Then Exit Sub
    
    ' Check if we we're already playing the same song
    If (mMusicFilename = FileName And mMusicID <> 0 And mFadeID = 0) Then Exit Sub
        
    ' Continue playing the music if we're trying to play the same music again while is active or fading
    If (Fade) Then
    
        ' If we were already fading, then exchange the fading song to master
        If (mFadeID <> 0) Then
            
            If (mFadeFilename = FileName) Then Exit Sub
        
            If (mMusicID <> 0) Then
                Call wGL_Audio.Stop(mMusicID)
            End If
            
            If (mMusicBuffer <> 0) Then
                Call wGL_Audio.DeleteSound(mMusicBuffer)
            End If
            
            mMusicID = mFadeID
            mMusicBuffer = mFadeBuffer
            mMusicFilename = mFadeFilename
            
        End If
        
        mFadeFilename = FileName
        mFadeBuffer = wGL_Audio.CreateSound(mMusicPath & FileName, AUDIO_FORMAT_MP3)
        mFadeID = wGL_Audio.Play(CHANNEL_MUSIC, mFadeBuffer, 0, MusicVolume, Repeat, False)
        mFadeTick = FrameTime
    Else
        Call StopMusic
        
        If (mMusicBuffer <> 0) Then
            Call wGL_Audio.DeleteSound(mMusicBuffer)
        End If
            
        mMusicFilename = FileName
        mMusicBuffer = wGL_Audio.CreateSound(mMusicPath & FileName, AUDIO_FORMAT_MP3)
        mMusicID = wGL_Audio.Play(CHANNEL_MUSIC, mMusicBuffer, 0, 1#, Repeat, False)
        
        Call wGL_Audio.Track(mMusicID)
    End If

End Sub

Public Sub PlayMusicRandomly(ByVal Fade As Boolean)
    
    If (UBound(mMusicList) > 0) Then
        Call PlayMusic(mMusicList(RandomNumber(1, UBound(mMusicList))) & ".mp3", False, Fade)
    End If
    
End Sub

Public Sub StopMusic()
    
    If (mFadeID <> 0) Then
        Call wGL_Audio.Stop(mFadeID)
        
        mFadeID = 0
    End If
    
    If (mFadeBuffer <> 0) Then
        Call wGL_Audio.DeleteSound(mFadeBuffer)
        
        mFadeBuffer = 0
    End If
    
    mFadeFilename = vbNullString
    
    If (mMusicID <> 0) Then
        Call wGL_Audio.Stop(mMusicID)
        
        mMusicID = 0
    End If

End Sub

Public Function PlayEffect(ByVal FileName As String, Optional ByVal Source As Long = 0, Optional ByVal Repeat As Boolean = False) As Long

    If Not mEffectEnabled Then Exit Function

    PlayEffect = Play(CHANNEL_EFFECT, FileName, Source, Repeat, True)
    
End Function
Public Sub PlayInterface(ByVal FileName As String)

    If Not mInterfaceEnabled Then Exit Sub

    Call Play(CHANNEL_UI, FileName, 0, False, False)
    
End Sub

Public Sub PlayEffectAt(ByVal FileName As String, ByVal X As Single, ByVal Y As Single)

    If Not mEffectEnabled Then Exit Sub

    Dim Effect As Long
    Effect = FindEffectFileInBank(FileName)
    
    If (Effect = 0) Then
        Effect = wGL_Audio.CreateSound(mEffectPath & FileName, AUDIO_FORMAT_WAV)
        
        If (Effect = 0) Then Exit Sub
        
        Call mEffectBank.Add(Item:=Effect, Key:=FileName)
    End If
    
    Call wGL_Audio.PlayAt(CHANNEL_EFFECT, Effect, CreateDefaultEmitter(X, Y), True)

End Sub

Public Function PlayAmbient(ByVal FileName As String, Optional ByVal Repeat As Boolean = False, _
                            Optional ByVal Channel As Long = CHANNEL_EFFECT) As Long

    If Not mEffectEnabled Then Exit Function
    
    PlayAmbient = Play(Channel, FileName, 0, Repeat, True)
    
End Function
Public Sub StopAmbient(ByRef ID As Long)
    
    If (ID <> 0) Then
        Call wGL_Audio.Stop(ID)
    End If
    
    ID = 0

End Sub
Public Property Get Effect3D() As Boolean

    Effect3D = mEffect3D
    
End Property

Public Property Let Effect3D(ByVal nValue As Boolean)

    If mEffect3D = nValue Then Exit Property

    mEffect3D = nValue

    Call wGL_Audio.Set3D(nValue)
    
End Property
Public Property Get MasterActivated() As Boolean

    MasterActivated = mMasterEnabled
    
End Property

Public Property Let MasterActivated(ByVal nValue As Boolean)

    If mMasterEnabled = nValue Then Exit Property

    mMasterEnabled = nValue

    If nValue Then
        Call wGL_Audio.SetMasterVolume(mMasterVolume * 0.01)
    Else
        Call wGL_Audio.SetMasterVolume(0)
    End If
    
End Property
Public Property Let MasterVolume(ByVal Volume As Long)

    If Volume < 0 Or Volume > 100 Then Exit Property

    Call wGL_Audio.SetMasterVolume(Volume * 0.01)
    
    mMasterVolume = Volume
    
End Property

Public Property Get MasterVolume() As Long

    MasterVolume = wGL_Audio.GetMasterVolume() * 100
    
End Property

Public Property Get MusicActivated() As Boolean

    MusicActivated = mMusicEnabled
    
End Property

Public Property Let MusicActivated(ByVal nValue As Boolean)

    If mMusicEnabled = nValue Then Exit Property

    mMusicEnabled = nValue

    If nValue Then
        Call PlayMusicRandomly(True)
    Else
        Call StopMusic
    End If
    
End Property

Public Property Let MusicVolume(ByVal Volume As Long)

    If Volume < 0 Or Volume > 100 Then Exit Property

    Call wGL_Audio.SetSubmixVolume(CHANNEL_MUSIC, Volume * 0.01)
    
End Property

Public Property Get MusicVolume() As Long

    MusicVolume = wGL_Audio.GetSubmixVolume(CHANNEL_MUSIC) * 100
    
End Property

Public Property Get EffectActivated() As Boolean

    EffectActivated = mEffectEnabled
    
End Property

Public Property Let EffectActivated(ByVal nValue As Boolean)

    If mEffectEnabled = nValue Then Exit Property

    mEffectEnabled = nValue

    If nValue Then
        Call wGL_Audio.SetSubmixVolume(CHANNEL_EFFECT, mEffectVolume * 0.01)
    Else
        Call wGL_Audio.SetSubmixVolume(CHANNEL_EFFECT, 0)
    End If
    
End Property

Public Property Let EffectVolume(ByVal Volume As Long)

    If Volume < 0 Or Volume > 100 Then Exit Property

    Call wGL_Audio.SetSubmixVolume(CHANNEL_EFFECT, Volume * 0.01)
    
    mEffectVolume = Volume
    
End Property

Public Property Get EffectVolume() As Long

    EffectVolume = mEffectVolume
    
End Property

Public Property Get InterfaceActivated() As Boolean

    InterfaceActivated = mInterfaceEnabled
    
End Property

Public Property Let InterfaceActivated(ByVal nValue As Boolean)

    If mInterfaceEnabled = nValue Then Exit Property

    mInterfaceEnabled = nValue

    If nValue Then
        Call wGL_Audio.SetSubmixVolume(CHANNEL_UI, mInterfaceVolume * 0.01)
    Else
        Call wGL_Audio.SetSubmixVolume(CHANNEL_UI, 0)
    End If
    
End Property

Public Property Let InterfaceVolume(ByVal Volume As Long)

    If Volume < 0 Or Volume > 100 Then Exit Property

    Call wGL_Audio.SetSubmixVolume(CHANNEL_UI, Volume * 0.01)
    
    mInterfaceVolume = Volume

End Property

Public Property Get InterfaceVolume() As Long

    InterfaceVolume = mInterfaceVolume
    
End Property

Private Function Play(ByVal Channel As Long, ByVal FileName As String, ByVal Source As Long, ByVal Repeat As Boolean, ByVal Reverb As Boolean) As Long

    If Not mEffectEnabled Then Exit Function

    Dim Effect As Long
    Effect = FindEffectFileInBank(FileName)
    
    If (Effect = 0) Then
        Effect = wGL_Audio.CreateSound(mEffectPath & FileName, AUDIO_FORMAT_WAV) '
        
        If (Effect = 0) Then Exit Function
        
        Call mEffectBank.Add(Item:=Effect, Key:=FileName)
    End If
    
    Play = wGL_Audio.Play(Channel, Effect, Source, 1#, Repeat, Reverb)

End Function

Private Function FindEffectFileInBank(ByVal Key As String) As Long

On Error GoTo clsAudio_FindEffectFileInBank_Error
    

    FindEffectFileInBank = mEffectBank.Item(Key)

    Exit Function
    
clsAudio_FindEffectFileInBank_Error:
    
    FindEffectFileInBank = 0
    
End Function

Private Static Sub wGL_Audio_Callback_OnStreamFinished(ByVal Instance As Long)
    
    If (Instance = mMusicID) Then
        mMusicID = 0
    
        Call PlayMusicRandomly(False)
    End If
    
End Sub

Private Function CreateDefaultEmitter(ByVal X As Single, ByVal Y As Single) As wGL_Audio_Emitter

    CreateDefaultEmitter.PositionX = X
    CreateDefaultEmitter.PositionY = Y
    CreateDefaultEmitter.PositionZ = 0#
    CreateDefaultEmitter.VelocityX = 0#
    CreateDefaultEmitter.VelocityY = 0#
    CreateDefaultEmitter.VelocityZ = 0#
    CreateDefaultEmitter.ChannelRadius = 1#
    CreateDefaultEmitter.CurveDistanceScaler = 12.25
    CreateDefaultEmitter.DopplerScaler = 1#
    CreateDefaultEmitter.InnerRadius = 2#
    CreateDefaultEmitter.InnerRadiusAngle = 3.14 / 4#
    
End Function
Private Sub OnFadeEffect(ByVal delta As Long)
    
    Dim Time As Single
    Time = (delta / EFFECT_FADE_TIME)
    
    ' Calculate the fade in/out gain between 0.0 and 1.0 using easeInOutCubic sigmoid function
    Dim Factor As Single
    Factor = IIf(Time < 0.5, 4 * Time * Time * Time, 1 - ((-2# * Time + 2#) ^ 3#) / 2#)
    
    ' Normalize in-case the user alt-tab and tick went to heaven (you will thank me later!)
    If (Factor > 1#) Then Factor = 1#
    
    ' Adjust both sound's gain value
    Call wGL_Audio.Adjust(mFadeID, Factor)
    
    If (mMusicID <> 0) Then
        Call wGL_Audio.Adjust(mMusicID, 1# - Factor)
    End If
    
    ' Release the other sound if the fade has finished
    If (delta >= EFFECT_FADE_TIME) Then
        If (mMusicID <> 0) Then
            Call wGL_Audio.Stop(mMusicID)
        End If
        
        If (mMusicBuffer <> 0) Then
            Call wGL_Audio.DeleteSound(mMusicBuffer)
        End If
        
        mMusicID = mFadeID
        mMusicBuffer = mFadeBuffer
        mMusicFilename = mFadeFilename
            
        mFadeID = 0
        mFadeBuffer = 0
        mFadeFilename = vbNullString
    End If
        
End Sub

