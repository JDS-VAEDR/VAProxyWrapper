Imports System.IO

Namespace Global.JDS.VAProxyWrapper.Core
  Public Class Paths
    Private ReadOnly Property vaProxy As Object
    Private ReadOnly Property Parent As VAProxy
    
    Public Sub New(proxy As Object, vaParent As VAProxy)
      ' Dynamically typecast the proxy as one of the VoiceAttackInitProxyClass types
      Dim isProxyType As Type = proxy.GetType
      ' And save reference to the passed proxy and parent 
      _vaProxy = Convert.ChangeType(proxy, isProxyType)
      _Parent = vaParent
    End Sub

    ' The currently executing plugin path (VA Help file notes this is a function - requires ())
    Public ReadOnly Property PluginPath As DirectoryInfo
      Get
        Return New DirectoryInfo(vaProxy.PluginPath())
      End Get
    End Property

    ' The Voice Attack installation directory
    Public ReadOnly Property VoiceAttack As DirectoryInfo
      Get
        Return New DirectoryInfo(vaProxy.InstallDir)
      End Get
    End Property

    ' The Voice Attack sounds directory
    Public ReadOnly Property Sounds As DirectoryInfo
      Get
        Return New DirectoryInfo(vaProxy.SoundsDir)
      End Get
    End Property

    ' The Voice Attack plugins directory
    Public ReadOnly Property Plugins As DirectoryInfo
      Get
        Return New DirectoryInfo(vaProxy.AppsDir)
      End Get
    End Property

    ' The Voice Attack assemblies directory
    Public ReadOnly Property Assemblies As DirectoryInfo
      Get
        Return New DirectoryInfo(vaProxy.AssembliesDir)
      End Get
    End Property

  End Class
End Namespace
