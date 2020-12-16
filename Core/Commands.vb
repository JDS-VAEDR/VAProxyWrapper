Namespace Global.JDS.VAProxyWrapper.Core
  Public Class Commands
    Private ReadOnly Property vaProxy As Object
    Private ReadOnly Property Parent As VAProxy
    
    Public Sub New(proxy As Object, vaParent As VAProxy)
      ' Dynamically typecast the proxy as one of the VoiceAttackInitProxyClass types
      Dim isProxyType As Type = proxy.GetType
      ' And save reference to the passed proxy and parent 
      _vaProxy = Convert.ChangeType(proxy, isProxyType)
      _Parent = vaParent

    End Sub

    ' (En|Dis)ables execution of the passed command by name (this session only)
    Overloads Public Property isSessionEnabled(CommandName As String) As Boolean
      Get
        Return vaProxy.Command.GetSessionEnabled(CommandName)
      End Get
      Set (isEnabled As Boolean)
        vaProxy.Command.SetSessionEnabled(CommandName, isEnabled)
      End Set
    End Property

    ' (En|Dis)ables the passed command by InternalID (this session only)
    Overloads Public Property isSetSessionEnabled(InternalID As Guid) As Boolean
      Get
        Return vaProxy.Command.GetSessionEnabled(InternalID)
      End Get
      Set (isEnabled As Boolean)
        vaProxy.Command.SetSessionEnabled(InternalID, isEnabled)
      End Set
    End Property

    ' (En|Dis)ables the passed Category in its entirety (this session only)
    Public Property isSessionEnabledByCategory(CategoryName As String) As Boolean
      Get
        Return vaProxy.Command.GetSessionEnabledByCategory(CategoryName)
      End Get
      Set (isEnabled As Boolean)
        vaProxy.Command.SetSessionEnabledByCategory(CategoryName, isEnabled)
      End Set
    End Property

    ' Runs the passed command by name if it exists
    Overloads Public Sub Execute(CommandName As String, Optional WaitForReturn As Boolean = False, Optional RunAsSubCommand As Boolean = False)
      vaProxy.Command.Execute(CommandName, WaitForReturn, RunAsSubCommand)
    End Sub

    ' Runs the passed command by internal id if it exists
    Overloads Public Sub Execute(InternalID As Guid, Optional WaitForReturn As Boolean = False, Optional RunAsSubCommand As Boolean = False)
      vaProxy.Command.Execute(InternalID, WaitForReturn, RunAsSubCommand)
    End Sub

    ' Runs the passed command by name (with every available optional argument) if it exists
    Overloads Public Sub Execute(CommandName As String, Optional WaitForReturn As Boolean = False, Optional RunAsSubCommand As Boolean = False,
                                 Optional CompletedAction As Action(Of Guid) = Nothing, Optional PassedText As String = Nothing,
                                 Optional PassedIntegers As String = Nothing, Optional PassedDecimals As String = Nothing,
                                 Optional PassedBooleans As String = Nothing, Optional PassedDates As String = Nothing)
      
      vaProxy.Command.Execute(CommandName, WaitForReturn, RunAsSubCommand, CompletedAction, PassedText, PassedIntegers, PassedDecimals, PassedBooleans, PassedDates)
    End Sub

    ' Runs the passed command by internal id (with every available optional arguments) if it exists
    Overloads Public Sub Execute(InternalID As Guid, Optional WaitForReturn As Boolean = False, Optional RunAsSubCommand As Boolean = False,
                                 Optional CompletedAction As Action(Of Guid?) = Nothing, Optional PassedText As String = Nothing,
                                 Optional PassedIntegers As String = Nothing, Optional PassedDecimals As String = Nothing,
                                 Optional PassedBooleans As String = Nothing, Optional PassedDates As String = Nothing)

      vaProxy.Command.Execute(InternalID, WaitForReturn, RunAsSubCommand, CompletedAction, PassedText, PassedIntegers, PassedDecimals, PassedBooleans, PassedDates)
    End Sub

    ' Checks if a command exists given the passed name
    Overloads Public ReadOnly Property Exists(CommandName As String) As Boolean
      Get
        Return vaProxy.Command.Exists(CommandName)
      End Get
    End Property

    ' Checks if a command exists given the passed internal id
    Overloads Public ReadOnly Property Exists(InternalID As Guid) As Boolean
      Get 
        Return vaProxy.Command.Exists(InternalID)
      End Get
    End Property

    ' Checks if a command is active given the passed internal id
    Overloads Public ReadOnly Property Active(InternalID As Guid) As Boolean
      Get 
        Return vaProxy.Command.Active(InternalID)
      End Get
    End Property

    ' Returns an int indicating the number of running instances of the passed command indicated by a spoken phrase
    Overloads Public ReadOnly Property ActiveCount(CommandPhrase As String) As Integer
      Get
        Return vaProxy.Command.ActiveCount(CommandPhrase)
      End Get
    End Property

    ' Returns an int indicating the number of running instances of the passed command indicated by a spoken phrase
    Overloads Public ReadOnly Property ActiveCount(InternalID As Guid) As Integer
      Get
        Return vaProxy.Command.ActiveCount(InternalID)
      End Get
    End Property

    ' Whether a category exists in the profile
    Public ReadOnly Property CategoryExists(CategoryName As String) As Boolean
      Get
        Return vaProxy.Command.CategoryExists(CategoryName)
      End Get
    End Property

    ' Number of seconds since last command was executed - spoken, keyboard, mouse click, joystick button press.
    ' Does NOT include subcommands, right-click executions, external command executions, etc
    Public ReadOnly Property TimeSinceLastUserExec() As Integer
      Get
        Return vaProxy.Command.LastUserExec()
      End Get
    End Property

    ' Total number of top-level commands executed in this session 
    Public ReadOnly Property ExecutionCount As Integer
      Get
        Return vaProxy.Command.ExecutionCount()
      End Get
    End Property

    ' Returns the most recent spoken command phrase (last executed command using voice)
    Public ReadOnly Property LastSpoken As String
      Get
        Return vaProxy.Command.LastSpoken()
      End Get
    End Property

    ' Returns the next most recently spoken command phrase (next-to-last executed command using voice)
    Public ReadOnly Property PreviousSpoken As String
      Get
        Return vaProxy.Command.PreviousSpoken()
      End Get
    End Property

    ' Returns a string of the spoken command at the passed index (zero indexed) 
    Public ReadOnly Property SpokenHistory(index As Integer) As String
      Get
        Return vaProxy.Command.SpokenHistory(index)
      End Get
    End Property

  End Class
End Namespace
