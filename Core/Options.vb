Namespace Global.JDS.VAProxyWrapper.Core
  Public Class Options
    Private ReadOnly Property vaProxy As Object
    Private ReadOnly Property Parent As VAProxy
    
    Public Sub New(proxy As Object, vaParent As VAProxy)
      ' Dynamically typecast the proxy as one of the VoiceAttackInitProxyClass types
      Dim isProxyType As Type = proxy.GetType
      ' And save reference to the passed proxy and parent 
      _vaProxy = Convert.ChangeType(proxy, isProxyType)
      _Parent = vaParent
    End Sub

    ' Voice Attack flag for plugins enabled
    Public ReadOnly Property PluginsEnabled As Boolean
      Get
        Return vaProxy.PluginsEnabled
      End Get
    End Property

    ' Voice Attack flag for nested tokens enabled
    Public ReadOnly Property NestedTokensEnabled As Boolean
      Get
        Return vaProxy.NestedTokensEnabled
      End Get
    End Property

    ' Voice Attack flag for auto profile switching enabled
    Public ReadOnly Property AutoProfileSwitchingEnabled As Boolean
      Get
        Return vaProxy.AutoProfileSwitchingEnabled
      End Get
    End Property

  End Class
End Namespace
