Namespace Global.JDS.VAProxyWrapper.Core

  ' This class is a wrapper for the command object of Voice Attack's proxy
  Public Class Queue
    Private ReadOnly Property VAProxy As Object
    Private ReadOnly Property Parent As VAProxy
    
    Public Sub New(proxy As Object, vaParent As VAProxy)
      ' Dynamically typecast the proxy as one of the VoiceAttackInitProxyClass types
      Dim isProxyType As Type = proxy.GetType
      ' And save reference to the passed proxy and parent 
      _vaProxy = Convert.ChangeType(proxy, isProxyType)
      _Parent = vaParent
    End Sub

    ' Returns number of command execution queues currently available
    Public Function Count() As Integer
      Return VAProxy.Queue.Count()
    End Function

    ' Returns status of command execution queue "name"
    Public Function Status(name As String) As String
      Return VAProxy.Queue.Status(name)
    End Function

    ' Returns number of commands in the named command execution queue
    Public Function CommandCount(name As String) As Integer
      Return VAProxy.Queue.CommandCount(name)
    End Function

    ' Returns the name of the actively executing command in the named command execution queue
    Public Function ActiveCommandName(name As String) As String
      Return VAProxy.Queue.ActiveCommandName(name)
    End Function

    ' Enumeration of possible Status values
    Public Enum StatusValues
      NotInitialized
      Running
      Idle
      Paused
      Stopped
    End Enum
  
  End Class
End Namespace
