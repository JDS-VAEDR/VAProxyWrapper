Namespace Global.JDS.VAProxyWrapper.Core
  ' This class is a wrapper for the command object of Voice Attack's proxy
  Public Class Command
    Private ReadOnly Property vaProxy As Object
    Private ReadOnly Property Parent As VAProxy
    
    Public Sub New(proxy As Object, vaParent As VAProxy)
      ' Dynamically typecast the proxy as one of the VoiceAttackInitProxyClass types
      Dim isProxyType As Type = proxy.GetType
      ' And save reference to the passed proxy and parent 
      _vaProxy = Convert.ChangeType(proxy, isProxyType)
      _Parent = vaParent
    End Sub
    
    ' The name of the executing command
    Public ReadOnly Property Name As String
      Get
        Return vaProxy.Command.Name()
      End Get
    End Property
    
    ' The Internal identifier (InternalId) of the executing command
    Public ReadOnly Property Identifier As Guid
      Get
        Return vaProxy.Command.InternalID()
      End Get
    End Property

    ' Gets a segment from a dynamic command (zero indexed)
    Public ReadOnly Property GetSegment(index As Integer) As String
      Get
        Return vaProxy.Command.Segment(index)
      End Get
    End Property

    ' The executing command source (the action used to execute command)
    Public ReadOnly Property Source as CommandSource 
      Get
        Return CType(vaProxy.Command.Action(), CommandSource)
      End Get
    End Property

    ' The text found before a wildcard phrase (only when using wildcards. Ex: *rocket* & voice "This is a test rocketship" returns "This is a test")
    Public ReadOnly Property Before as String 
      Get
        Return vaProxy.Command.Before()
      End Get
    End Property

    ' The text found before a wildcard phrase (only when using wildcards. Ex: *rocket* & voice "This is a test rocketship" returns "ship")
    Public ReadOnly Property After as String 
      Get
        Return vaProxy.Command.After()
      End Get
    End Property

    ' The wildcard phrase itself (only when using wildcards. Ex: *rocket* returns "rocket")
    Public ReadOnly Property WildcardKey as String 
      Get
        Return vaProxy.Command.WildcardKey()
      End Get
    End Property

    ' The confidence of the voice recognition for the currently executing command (on a scale of 0-100)
    Public ReadOnly Property Confidence As Integer
      Get
        Return vaProxy.Command.Confidence()
      End Get
    End Property

    ' The minimum confidence of the voice recognition for the currently executing command (on a scale of 0-100)
    Public ReadOnly Property MinConfidence As Integer
      Get
        Return vaProxy.Command.MinConfidence()
      End Get
    End Property

    ' Is the currently executing command called by another command?
    Public ReadOnly Property IsSubCommand As Boolean
      Get
        Return vaProxy.Command.IsSubCommand()
      End Get
    End Property

    ' Is the currently executing command executed by a double-tap
    Public ReadOnly Property IsDoubleTapInvoked As Boolean
      Get
        Return vaProxy.Command.IsDoubleTapInvoked()
      End Get
    End Property

    ' Is the currently executing command executed by a long-press
    Public ReadOnly Property IsLongPressInvoked As Boolean
      Get
        Return vaProxy.Command.IsLongPressInvoked()
      End Get
    End Property
    
    ' The full value of the "When I Say" command input box
    Public ReadOnly Property WhenISay As String
      Get
        Return vaProxy.Command.WhenISay()
      End Get
    End Property

    ' Is the currently executing command executed with a listening override? ("computer, open door" vs. "open door")
    Public ReadOnly Property IsListeningOverride As Boolean
      Get
        Return vaProxy.Command.IsListeningOverride()
      End Get
    End Property

    ' Is the currently executing command a composite (prefix + suffix command)
    Public ReadOnly Property IsComposite As Boolean
      Get
        Return vaProxy.Command.IsComposite()
      End Get
    End Property

    ' If a composite command, returns the prefix part (empty string if not composite)
    Public ReadOnly Property PrefixPart As String
      Get
        Return vaProxy.Command.PrefixPart()
      End Get
    End Property
    
    ' If a composite command, returns the suffix part (empty string if not composite)
    Public ReadOnly Property SuffixPart As String
      Get
        Return vaProxy.Command.SuffixPart()
      End Get
    End Property
    
    ' If a composite command, returns the composite group if it is being used
    Public ReadOnly Property CompositeGroup As String
      Get
        Return vaProxy.Command.CompositeGroup()
      End Get
    End Property
    
    ' Returns category of executing command. If a composite, returns PrefixPart() & " " & SuffixPart()
    Public ReadOnly Property Category As String
      Get
        Return vaProxy.Command.Category()
      End Get
    End Property

    ' If a composite command, returns the prefix category (empty string if not composite)
    Public ReadOnly Property PrefixCategory As String
      Get
        Return vaProxy.Command.PrefixCategory()
      End Get
    End Property
    
    ' If a composite command, returns the suffix category (empty string if not composite)
    Public ReadOnly Property SuffixCategory As String
      Get
        Return vaProxy.Command.SuffixCategory()
      End Get
    End Property

    ' Is the currently executing command already executing in another instance?
    Public ReadOnly Property IsAlreadyExecuting As Boolean
      Get
        Return vaProxy.Command.AlreadyExecuting()
      End Get
    End Property

    ' Where a command originated from
    Public Enum CommandSource
      Spoken
      Keyboard
      Joystick
      Mouse
      Profile
      External
      Unrecognized
      ProfileUnloadChange
      ProfileUnloadClose
      DictationRecognized
      Plugin
      Other
    End Enum

  End Class
End Namespace