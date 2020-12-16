Imports System.IO

Namespace Global.JDS.VAProxyWrapper.Core
  Public Class Speech
    Private ReadOnly Property vaProxy As Object
    Private ReadOnly Property Parent As VAProxy
    
    Public Sub New(proxy As Object, vaParent As VAProxy)
      ' Dynamically typecast the proxy as one of the VoiceAttackInitProxyClass types
      Dim isProxyType As Type = proxy.GetType
      ' And save reference to the passed proxy and parent 
      _vaProxy = Convert.ChangeType(proxy, isProxyType)
      _Parent = vaParent
    End Sub

    Public Function GetCapturedAudio(type As AudioType) As MemoryStream
      ' Returns memory stream of captured voice attack audio
      Return vaProxy.Utility.CapturedAudio(type)
    End Function
    
    Public Sub ResetSpeechEngine
      vaProxy.Utility.ResetSpeechEngine()
    End Sub
    
    ' Returns flag for whether the recording device is muted
    Public ReadOnly Property IsMuted As Boolean
      Get
        Return vaProxy.Utility.GetSpeechRecordingDeviceMute()
      End Get
    End Property

    ' Sets mute value for the recording device
    Public Sub SetMuted(bMuteDevice As Boolean)
      vaProxy.Utility.SetSpeechRecordingDeviceMute(bMuteDevice)
    End Sub

  End Class
  
  Public Enum AudioType
    ' Last recognized voice command audio
    LastRecognized = 0
    
    ' Previous recognized voice command audio
    PreviousRecognized = 1

    ' Last unrecognized voice command audio
    LastUnrecognized = 2

    ' Last voice command audio
    Last = 3

    ' Previous voice command audio
    Previous = 4

    ' Voice command audio from dictation mode
    Dictated = 5
  End Enum
End Namespace
