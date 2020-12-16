Namespace Global.JDS.VAProxyWrapper.Core
  Public Class Log

    Private ReadOnly Property vaProxy As Object
    Private ReadOnly Property Parent As VAProxy
    
    Public Sub New(proxy As Object, vaParent As VAProxy)
      ' Dynamically typecast the proxy as one of the VoiceAttackInitProxyClass types
      Dim isProxyType As Type = proxy.GetType
      ' And save reference to the passed proxy and parent 
      _vaProxy = Convert.ChangeType(proxy, isProxyType)
      _Parent = vaParent
    End Sub
    
    ' Writes to the Voice Attack log
    Public Sub Write(content As String, Optional color As Color = Color.Black)
      content = $"{Parent.LogPrefix} {content}"
      vaProxy.WriteToLog(content, color.ToString())
    End Sub
    
    ' Clears the Voice Attack log
    Public Sub Clear()
      vaProxy.ClearLog()
    End Sub

    ' Set up logging methods
    ' Log Trace messages (only for development!!!)
    Public Sub Trace(message As String)
      If IsLoggingEnabled(LogLevel.Trace) Then
        Write(message, Color.Gray)
      End If
    End Sub
    
    ' Log Debug messages
    Public Sub Debug(message As String)
      If IsLoggingEnabled(LogLevel.Debug) Then
        Write(message, Color.Blue)
      End If
    End Sub
    
    ' Log informational messages
    Public Sub Information(message As String)
      If IsLoggingEnabled(LogLevel.Information) Then
        Write(message, Color.Pink)
      End If
    End Sub
    
    ' Log warning messages
    Public Sub Warning(message As String)
      If IsLoggingEnabled(LogLevel.Warning) Then
        Write(message, Color.Yellow)
      End If
    End Sub
    
    ' Log error messages
    Public Sub [Error](message As String)
      If IsLoggingEnabled(LogLevel.Error) Then
       Write(message, Color.Orange)
      End If
    End Sub
    
    ' Log critical/fatal messages
    Public Sub Critical(message As String)
      If IsLoggingEnabled(LogLevel.Critical) Then
       Write(message, Color.Red)
      End If
    End Sub
    
    ' If the current loglevel (for writing the log) is >= the level defined in the core, we can log the message
    Private Function IsLoggingEnabled(logLevel As LogLevel) As Boolean
      Return logLevel >= Parent.LogLevel
    End Function

  End Class

    ' Colors Voice Attack accepts for logging
  Public Enum Color
    Red
    Blue
    Green
    Yellow
    Orange
    Purple
    Blank
    Black
    Gray
    Pink
  End Enum
  
  Public Enum LogLevel
    Trace = 0	
    Debug = 1	
    Information = 2	
    Warning = 3	
    [Error] = 4	
    Critical = 5	
    None = 6	
  End Enum

End Namespace
