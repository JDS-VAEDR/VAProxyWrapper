Imports VA = VoiceAttack
Imports JDS.VAProxyWrapper.Core
Imports JDS.VAProxyWrapper.Interfaces

Namespace Global.JDS.VAProxyWrapper
  Public Class VAProxy
    Implements IVoiceAttackProxy
    
    Public ReadOnly Property ProxyType As String Implements IVoiceAttackProxy.ProxyType
    Private ReadOnly Property VAProxy As Object Implements IVoiceAttackProxy.VAProxy
    Public ReadOnly Property LogLevel As LogLevel Implements IVoiceAttackProxy.LogLevel
    Public ReadOnly Property LogPrefix As String Implements IVoiceAttackProxy.LogPrefix
    
    Public Sub New(proxy As Object, logLevel As LogLevel, LogPrefix As String)
      ' Dynamically typecast the proxy as one of the VoiceAttackInitProxyClass types
      Dim isProxyType As Type = proxy.GetType
      ' And save reference to proxy 
      _VAProxy = Convert.ChangeType(proxy, isProxyType)

      ' Save Proxy Type
      ProxyType = _VAProxy.GetType().ToString()
      
      ' Save Logging values passed during construction
      _LogLevel = LogLevel
      _LogPrefix = LogPrefix

      ' Create Proxy wrapper class instances
      ActiveWindow = New ActiveWindow(VAProxy, Me)
      Command = New Command(VAProxy, Me)
      Commands = New Commands(VAProxy, Me)
      EventHandlers = New EventHandlers(VAProxy, Me)
      Log = New Log(VAProxy, Me)
      Options = New Options(VAProxy, Me)
      Paths = New Paths(VAProxy, Me)
      Profile = New Profile(VAProxy, Me)
      Queue = New Queue(VAProxy, Me)
      Speech = New Speech(VAProxy, Me)
      Variables = New Variables(VAProxy, Me)
      Versions = New Versions(VAProxy, Me)
    End Sub


    Public Sub Invoke(proxy As Object) Implements IVoiceAttackProxy.Invoke
      ' Dynamically update the typecast of the proxy as VoiceAttackInvokeProxyClass
      Dim isProxyType As Type = proxy.GetType
      ' And save reference to proxy 
      _VAProxy = Convert.ChangeType(proxy, isProxyType)

      ' Save Proxy Type
      _ProxyType = _VAProxy.GetType().ToString()
    End Sub

    Public Sub [Exit](proxy As Object) Implements IVoiceAttackProxy.Exit
      ' Dynamically update the typecast of the proxy as VoiceAttackProxyClass
      Dim isProxyType As Type = proxy.GetType
      ' And save reference to proxy 
      _VAProxy = Convert.ChangeType(proxy, isProxyType)

      ' Save Proxy Type
      _ProxyType = _VAProxy.GetType().ToString()
    End Sub

    Public Function ExtractPhrases(query As String, Optional trimSpaces As Boolean = False, Optional lowercase As Boolean = False) As IReadOnlyCollection(Of String) Implements IVoiceAttackProxy.ExtractPhrases
      Return VAProxy.ExtractPhrases(query, trimSpaces, lowercase)
    End Function

    Public Function ParseTokens(token As String) As String Implements IVoiceAttackProxy.ParseTokens
      Return VAProxy.ParseTokens(token)
    End Function

    Public Sub Opacity(percentage As Integer) Implements IVoiceAttackProxy.Opacity
      VAProxy.SetOpacity(percentage)
    End Sub

    Public Sub ResetStopFlag() Implements IVoiceAttackProxy.ResetStopFlag
      VAProxy.ResetStopFlag()
    End Sub

    Public Sub Close() Implements IVoiceAttackProxy.Close
      VAProxy.Close()
    End Sub
   
    Public ReadOnly Property Context As String Implements IVoiceAttackProxy.Context
      Get
        Return VAProxy.Context
      End Get
    End Property

    Public ReadOnly Property Stopped As Boolean Implements IVoiceAttackProxy.Stopped
      Get
        Return VAProxy.Stopped
      End Get
    End Property

    ' Returns list of Profile Names
    Public ReadOnly Property Profiles As IReadOnlyCollection(Of String) Implements IVoiceAttackProxy.ProfileNames
      Get
        Return VAProxy.ProfileNames()
      End Get
    End Property

    ' Returns list of Profile internal IDs
    Public ReadOnly Property ProfileInternalIDs As IReadOnlyCollection(Of Guid) Implements IVoiceAttackProxy.ProfileInternalIDs
    
    ' Returns the "Session State" provided by Voice Attack (we instantiate, so this isn't really necessary)
    Public ReadOnly Property SessionState As IReadOnlyDictionary(Of String, Object) Implements IVoiceAttackProxy.SessionState

    ' The Main Window Handle of Voice Attack
    Public ReadOnly Property MainWindowHandle As IntPtr Implements IVoiceAttackProxy.MainWindowHandle
      Get
        Return VAProxy.MainWindowHandle
      End Get
    End Property

    Public ReadOnly Property ActiveWindow As ActiveWindow Implements IVoiceAttackProxy.ActiveWindow
    Public ReadOnly Property Command As Command Implements IVoiceAttackProxy.Command
    Public ReadOnly Property Commands As Commands Implements IVoiceAttackProxy.Commands
    Public ReadOnly Property EventHandlers As EventHandlers Implements IVoiceAttackProxy.EventHandlers
    Public ReadOnly Property Log As Log Implements IVoiceAttackProxy.Log
    Public ReadOnly Property Options As Options Implements IVoiceAttackProxy.Options
    Public ReadOnly Property Paths As Paths Implements IVoiceAttackProxy.Paths
    Public ReadOnly Property Profile As Profile Implements IVoiceAttackProxy.Profile
    Public ReadOnly Property Queue As Queue Implements IVoiceAttackProxy.Queue
    Public ReadOnly Property Speech As Speech Implements IVoiceAttackProxy.Speech
    Public ReadOnly Property Variables As Variables Implements IVoiceAttackProxy.Variables
    Public ReadOnly Property Versions As Versions Implements IVoiceAttackProxy.Versions

  End Class
End Namespace