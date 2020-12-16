Imports System

Namespace Global.JDS.VAProxyWrapper.Core
  Public Class Versions
    Private ReadOnly Property vaProxy As Object
    Private ReadOnly Property Parent As VAProxy
    
    Public Sub New(proxy As Object, vaParent As VAProxy)
      ' Dynamically typecast the proxy as one of the VoiceAttackInitProxyClass types
      Dim isProxyType As Type = proxy.GetType
      ' And save reference to the passed proxy and parent 
      _vaProxy = Convert.ChangeType(proxy, isProxyType)
      _Parent = vaParent
    End Sub

    ' The Voice Attack Proxy interface version
    Public ReadOnly Property Proxy As Version 
      Get
        Return vaProxy.ProxyVersion
      End Get
    End Property

    ' The Voice Attack application version
    Public ReadOnly Property VoiceAttack As Version
      Get
        Return vaProxy.VAVersion
      End Get
    End Property

    ' The Voice Attack beta flag (is this a stable release?)
    Public ReadOnly Property IsStableRelease As Boolean
      Get
        Return vaProxy.IsRelease
      End Get
    End Property

    ' The Voice Attack trial flag (is this a trial version?)
    Public ReadOnly Property IsTrialVersion As Boolean
      Get
        Return vaProxy.IsTrial
      End Get
    End Property

  End Class
End Namespace