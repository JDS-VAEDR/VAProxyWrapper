Namespace Global.JDS.VAProxyWrapper.Core
  Public Class Variables
    Private ReadOnly Property vaProxy As Object
    Private ReadOnly Property Parent As VAProxy
    
    Public Sub New(proxy As Object, vaParent As VAProxy)
      ' Dynamically typecast the proxy as one of the VoiceAttackInitProxyClass types
      Dim isProxyType As Type = proxy.GetType
      ' And save reference to the passed proxy and parent 
      _vaProxy = Convert.ChangeType(proxy, isProxyType)
      _Parent = vaParent
    End Sub

    ' Sets a variable
    ' @typeparam T - Type of the variable
    ' @param Name - name of the variable
    ' @param value - the value of the variable
    Public Sub [Set](Of T)(name As String, value As T)
      Dim code As TypeCode = Convert.GetTypeCode(value)
      [Set](name, value, code)
    End Sub
    
    Public Sub [Set](name As String, value As Object, code As TypeCode)
      Select Case code
        ' Set a Boolean value
        Case TypeCode.Boolean
          SetBoolean(name, CBool(Convert.ChangeType(value, GetType(Boolean))))

        ' Set a DateTime value
        Case TypeCode.DateTime  
          SetDate(name, CDate(Convert.ChangeType(value, GetType(DateTime))))

        ' Set a Decimal value
        Case TypeCode.Decimal, TypeCode.Single, TypeCode.Double
          SetDecimal(name, CDec(Convert.ChangeType(value, GetType(Decimal))))

        ' Set an Int value
        Case TypeCode.Int32, TypeCode.UInt32, TypeCode.Int64, TypeCode.UInt64
          SetInt(name, CInt(Convert.ChangeType(value, GetType(Integer))))

        ' Set a SmallInt value (Short)
        Case TypeCode.Byte, TypeCode.Int16, TypeCode.UInt16, TypeCode.SByte
          SetSmallInt(name, CShort(Convert.ChangeType(value, GetType(Short))))

        ' Set a String value
        Case TypeCode.String, TypeCode.Char
          SetText(name, CStr(Convert.ChangeType(value, GetType(String))))

        ' Set an Object value
        Case TypeCode.Object
          Dim newCode = Convert.GetTypeCode(value)
          [Set](name, value, newCode)

        ' Type is empty, skip entirely & log
        Case TypeCode.Empty
          Parent.Log.Trace($"Skipping '{name}' variable - no type specified")
          
        ' Log warning if nothing above matched
        Case Else
          Parent.Log.Warning($"Could not set '{name}' to '{value}'. Type '{code.ToString()}' is not supported")

      End Select
    End Sub
    
    ' Gets a variable from the Voice Attack proxy
    Public Function [Get](Of T)(name As String) As T
      Dim code As TypeCode = Type.GetTypeCode(GetType(T))
      
      Select Case code
        Case TypeCode.Boolean
          Return CType(Convert.ChangeType(GetBoolean(name), GetType(T)), T)

        Case TypeCode.DateTime
          Return CType(Convert.ChangeType(GetDate(name), GetType(T)), T)

        Case TypeCode.Decimal, TypeCode.Single, TypeCode.Double
          Return CType(Convert.ChangeType(GetDecimal(name), GetType(T)), T)

        Case TypeCode.Int32, TypeCode.UInt32, TypeCode.Int64, TypeCode.UInt64
          Return CType(Convert.ChangeType(GetInt(name), GetType(T)), T)

        Case TypeCode.Byte, TypeCode.Int16, TypeCode.UInt16, TypeCode.SByte
          Return CType(Convert.ChangeType(GetSmallInt(name), GetType(T)), T)

        Case TypeCode.String, TypeCode.Char
          Return CType(Convert.ChangeType(GetText(name), GetType(T)), T)

        Case Else
          Parent.Log.Warning($"Could not retrieve '{name}'. Type '{code.ToString()}' is unsupported.")
          return CType(Nothing, T)

      End Select
    End Function

    ' Gets a boolean variable from the current profile through the Voice Attack proxy
    Private Function GetBoolean(name As String) As Boolean?
      Return vaProxy.GetBoolean(name)
    End Function
    
    ' Gets a date variable to the current profile through the Voice Attack proxy
    Private Function GetDate(name As String) As DateTime?
      Return vaProxy.GetDate(name)
    End Function
    
    ' Gets a decimal variable to the current profile through the Voice Attack proxy
    Private Function GetDecimal(name As String) As Decimal?
      Return vaProxy.GetDecimal(name)
    End Function
    
    ' Gets an integer variable to the current profile through the Voice Attack proxy
    Private Function GetInt(name As String) As Integer?
      Return vaProxy.GetInt(name)
    End Function
    
    ' Gets a small variable to the current profile through the Voice Attack proxy
    Private Function GetSmallInt(name As String) As Short?
      Return vaProxy.GetSmallInt(name)
    End Function
    
    ' Gets a text variable to the current profile through the Voice Attack proxy
    Private Function GetText(name As String) As String
      Return vaProxy.GetText(name)
    End Function
    
    ' Pushes a boolean variable to the current profile through the Voice Attack proxy
    Private Sub SetBoolean(name As String, value As Boolean?)
      Parent.Log.Trace(String.Format("Set '{0}' to '{1}'", $"{{BOOL:{name}}}", value))
      vaProxy.SetBoolean(name, value)
    End Sub
    
    ' Pushes a date variable to the current profile through the Voice Attack proxy
    Private Sub SetDate(name As String, value As DateTime?)
      Parent.Log.Trace(String.Format("Set '{0}' to '{1}'", $"{{DATE:{name}}}", value))
      vaProxy.SetDate(name, value)
    End Sub
    
    ' Pushes a decimal variable to the current profile through the Voice Attack proxy
    Private Sub SetDecimal(name As String, value As Decimal?)
      Parent.Log.Trace(String.Format("Set '{0}' to '{1}'", $"{{DEC:{name}}}", value))
      vaProxy.SetDecimal(name, value)
    End Sub

    ' Pushes an integer variable to the current profile through the Voice Attack proxy
    Private Sub SetInt(name As String, value As Integer?)
      Parent.Log.Trace(String.Format("Set '{0}' to '{1}'", $"{{INT:{name}}}", value))
      vaProxy.SetInt(name, value)
    End Sub

    ' Pushes a small variable to the current profile through the Voice Attack proxy
    Private Sub SetSmallInt(name As String, value As Short?)
      Parent.Log.Trace(String.Format("Set '{0}' to '{1}'", $"{{SMALL:{name}}}", value))
      vaProxy.SetSmallInt(name, value)
    End Sub

    ' Pushes a text variable to the current profile through the Voice Attack proxy
    Private Sub SetText(name As String, value As String)
      Parent.Log.Trace(String.Format("Set '{0}' to '{1}'", $"{{TXT:{name}}}", value))
      vaProxy.SetText(name, value)
    End Sub

  End Class
End Namespace
