Imports VA = VoiceAttack

Namespace Global.JDS.VAProxyWrapper.Core
  Public Class EventHandlers
   
    Private ReadOnly Property VAProxy As VA.VoiceAttackInitProxyClass
    Private ReadOnly Property Parent As VAProxy
    
    ' Events we will fire internally when we receive the equal event through the VA Proxy 
    Public Event ProfileChanging(FromInternalID As Guid?, ToInternalID As Guid?, FromName As String, ToName As String)
    Public Event ProfileChanged(FromInternalID As Guid?, ToInternalID As Guid?, FromName As String, ToName As String)
    Public Event CommandExecuting(InternalID As Guid?)
    Public Event CommandExecuted(InternalID As Guid?)
    Public Event BooleanChanged(Name As String, bFrom As Boolean?, bTo As Boolean?, InternalID As Guid?)
    Public Event DateChanged(Name As String, dFrom As DateTime?, dTo As DateTime?, InternalID As Guid?)
    Public Event DecimalChanged(Name As String, FromValue As Decimal?, ToValue As Decimal?, InternalID As Guid?)
    Public Event IntegerChanged(Name As String, iFrom As Integer?, iTo As Integer?, InternalID As Guid?)
    Public Event TextChanged(Name As String, sFrom As String, sTo As String, InternalID As Guid?)
    
    Public Sub New(proxy As Object, vaParent As VAProxy)
      ' Dynamically typecast the proxy as one of the VoiceAttackInitProxyClass types
      Dim isProxyType As Type = proxy.GetType
      ' And save reference to the passed proxy and parent 
      _VAProxy = Convert.ChangeType(proxy, isProxyType)
      _Parent = vaParent
      
      ' Create our event handlers
      Create()
    End Sub
    
    Public Sub Unload()
      ' Remove our event handlers
      Remove()
    End Sub

    ' Create the event handlers for Voice Attack events
    Public Sub Create()

      ' Profile events
      AddHandler VAProxy.ProfileChanging, AddressOf onProfileChanging
      AddHandler VAProxy.ProfileChanged, AddressOf onProfileChanged

      ' Command events (MUST have the Command Advanced Option "Enable Proxy Command Events" enabled)
      AddHandler vaProxy.CommandExecuting, AddressOf onCommandExecuting
      AddHandler vaProxy.CommandExecuted, AddressOf onCommandExecuted

      ' Variable events (a variable MUST end with '#' for its event to fire)
      AddHandler vaProxy.BooleanVariableChanged, AddressOf onBooleanChanged
      AddHandler vaProxy.DateVariableChanged, AddressOf onDateChanged
      AddHandler vaProxy.DecimalVariableChanged, AddressOf onDecimalChanged
      AddHandler vaProxy.IntegerVariableChanged, AddressOf onIntegerChanged
      AddHandler vaProxy.TextVariableChanged, AddressOf onTextChanged

    End Sub
    
    ' Remove the handlers created in CreateHandlers 
    Public Sub Remove()
      RemoveHandler vaProxy.ProfileChanging, AddressOf onProfileChanging
      RemoveHandler vaProxy.ProfileChanged, AddressOf onProfileChanged

      RemoveHandler vaProxy.CommandExecuting, AddressOf onCommandExecuting
      RemoveHandler vaProxy.CommandExecuted, AddressOf onCommandExecuted

      RemoveHandler vaProxy.BooleanVariableChanged, AddressOf onBooleanChanged
      RemoveHandler vaProxy.DateVariableChanged, AddressOf onDateChanged
      RemoveHandler vaProxy.DecimalVariableChanged, AddressOf onDecimalChanged
      RemoveHandler vaProxy.IntegerVariableChanged, AddressOf onIntegerChanged
      RemoveHandler vaProxy.TextVariableChanged, AddressOf onTextChanged
      
    End Sub
    
    Public Sub Reset()
      ' Reset the handlers 
      Remove()
      Create()
    End Sub

    ' Called when Voice Attack sends a ProfileChanging event through the proxy
    Public Sub onProfileChanging(FromInternalID As Guid?, ToInternalID As Guid?, FromName As String, ToName As String)
      RaiseEvent ProfileChanging(FromInternalID, ToInternalID, FromName, ToName)
    End Sub

    Public Sub onProfileChanged(FromInternalID As Guid?, ToInternalID As Guid?, FromName As String, ToName As String)
      RaiseEvent ProfileChanged(FromInternalID, ToInternalID, FromName, ToName)
    End Sub

    '*********************
    ' Command execution
    Public Sub onCommandExecuting(InternalID As Guid?)
      RaiseEvent CommandExecuting(InternalID)
    End Sub

    Public Sub onCommandExecuted(InternalID As Guid?)
      RaiseEvent CommandExecuted(InternalID)
    End Sub

    '*********************
    ' Variable changes 
    Public Sub onBooleanChanged(Name As String, bFrom As Boolean?, bTo As Boolean?, InternalID As Guid?)
      RaiseEvent BooleanChanged(Name, bFrom, bTo, InternalID)
    End Sub

    Public Sub onDateChanged(Name As String, dFrom As DateTime?, dTo As DateTime?, InternalID As Guid?)
      RaiseEvent DateChanged(Name, dFrom, dTo, InternalID)
    End Sub
      
    Public Sub onDecimalChanged(Name As String, FromValue As Decimal?, ToValue As Decimal?, InternalID As Guid?)
      RaiseEvent DecimalChanged(Name, FromValue, ToValue, InternalID)
    End Sub

    Public Sub onIntegerChanged(Name As String, iFrom As Integer?, iTo As Integer?, InternalID As Guid?)
      RaiseEvent IntegerChanged(Name, iFrom, iTo, InternalID)
    End Sub

    Public Sub onTextChanged(Name As String, sFrom As String, sTo As String, InternalID As Guid?)
      RaiseEvent TextChanged(Name, sFrom, sTo, InternalID)
    End Sub

  End Class
End Namespace
