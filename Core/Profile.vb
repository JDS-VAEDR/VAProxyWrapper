Namespace Global.JDS.VAProxyWrapper.Core
  
  ' This class is a wrapper for the command object of Voice Attack's proxy
  Public Class Profile
    Private ReadOnly Property VAProxy As Object
    Private ReadOnly Property Parent As VAProxy
    
    Public Sub New(proxy As Object, vaParent As VAProxy)
      ' Dynamically typecast the proxy as one of the VoiceAttackInitProxyClass types
      Dim isProxyType As Type = proxy.GetType
      ' And save reference to the passed proxy and parent 
      _vaProxy = Convert.ChangeType(proxy, isProxyType)
      _Parent = vaParent
    End Sub

    ' Returns the name of the currently active profile
    Public ReadOnly Property Name As String
      Get
        Return VAProxy.Profile.Name()
      End Get
    End Property

    ' Returns author flag internal id as a nullable Guid
    Public ReadOnly Property InternalID As Guid?
      Get
        Return VAProxy.Profile.InternalID()
      End Get
    End Property

    ' Returns flag for whether the profile exists based on passed name
    Overloads Public ReadOnly Property Exists(ProfileName As String) As Boolean
      Get
        Return VAProxy.Profile.Exists(ProfileName)
      End Get
    End Property

    ' Returns flag for whether the profile exists based on passed internal id
    Overloads Public ReadOnly Property Exists(ProfileInternalID As Guid) As Boolean
      Get
        Return VAProxy.Profile.Exists(ProfileInternalID)
      End Get
    End Property

    ' Returns author flag AuthorTag1
    Public ReadOnly Property AuthorTag1 As String
      Get
        Return VAProxy.Profile.AuthorTag1()
      End Get
    End Property

    ' Returns author flag AuthorTag2
    Public ReadOnly Property AuthorTag2 As String
      Get
        Return VAProxy.Profile.AuthorTag2()
      End Get
    End Property

    ' Returns author flag AuthorTag3
    Public ReadOnly Property AuthorTag3 As String
      Get
        Return VAProxy.Profile.AuthorTag3()
      End Get
    End Property

    ' Returns author flag AuthorID (nullable)
    Public ReadOnly Property AuthorID As Guid?
      Get
        Return VAProxy.Profile.AuthorID()
      End Get
    End Property

    ' Returns author flag ProductID (nullable)
    Public ReadOnly Property ProductID As Guid?
      Get
        Return VAProxy.Profile.ProductID()
      End Get
    End Property

    ' Returns the name of the previously loaded profile
    Public ReadOnly Property PreviousName As String
      Get
        Return VAProxy.Profile.PreviousName()
      End Get
    End Property

    ' Returns author flag internal id of the previously loaded profile (nullable)
    Public ReadOnly Property PreviousInternalID As Guid?
      Get
        Return VAProxy.Profile.PreviousInternalID()
      End Get
    End Property

    ' Returns author flag PreviousAuthorTag1
    Public ReadOnly Property PreviousAuthorTag1 As String
      Get
        Return VAProxy.Profile.PreviousAuthorTag1()
      End Get
    End Property

    ' Returns author flag PreviousAuthorTag2
    Public ReadOnly Property PreviousAuthorTag2 As String
      Get
        Return VAProxy.Profile.PreviousAuthorTag2()
      End Get
    End Property

    ' Returns author flag PreviousAuthorTag3
    Public ReadOnly Property PreviousAuthorTag3 As String
      Get
        Return VAProxy.Profile.PreviousAuthorTag3()
      End Get
    End Property

    ' Returns author flag PreviousAuthorID (nullable)
    Public ReadOnly Property PreviousAuthorID As Guid?
      Get
        Return VAProxy.Profile.PreviousAuthorID()
      End Get
    End Property

    ' Returns author flag PreviousProductID (nullable)
    Public ReadOnly Property PreviousProductID As Guid?
      Get
        Return VAProxy.Profile.PreviousProductID()
      End Get
    End Property

    ' Returns the name of the next loaded profile
    Public ReadOnly Property NextName As String
      Get
        Return VAProxy.Profile.NextName()
      End Get
    End Property

    ' Returns author flag internal id of the next loaded profile (nullable)
    Public ReadOnly Property NextInternalID As Guid?
      Get
        Return VAProxy.Profile.NextInternalID()
      End Get
    End Property

    ' Returns author flag NextAuthorTag1
    Public ReadOnly Property NextAuthorTag1 As String
      Get
        Return VAProxy.Profile.NextAuthorTag1()
      End Get
    End Property

    ' Returns author flag NextAuthorTag2
    Public ReadOnly Property NextAuthorTag2 As String
      Get
        Return VAProxy.Profile.NextAuthorTag2()
      End Get
    End Property

    ' Returns author flag NextAuthorTag3
    Public ReadOnly Property NextAuthorTag3 As String
      Get
        Return VAProxy.Profile.NextAuthorTag3()
      End Get
    End Property

    ' Returns author flag NextAuthorID (nullable)
    Public ReadOnly Property NextAuthorID As Guid?
      Get
        Return VAProxy.Profile.NextAuthorID()
      End Get
    End Property

    ' Returns author flag NextProductID (nullable)
    Public ReadOnly Property NextProductID As Guid?
      Get
        Return VAProxy.Profile.NextProductID()
      End Get
    End Property

    ' Returns the name of the next loaded profile
    Public ReadOnly Property History(index As Integer) As String
      Get
        Return VAProxy.Profile.History(index)
      End Get
    End Property

    ' Returns author flag internal id of the next loaded profile (nullable)
    Public ReadOnly Property HistoryInternalID(index As Integer) As Guid?
      Get
        Return VAProxy.Profile.HistoryInternalID(index)
      End Get
    End Property

    ' Returns author flag HistoryAuthorID (nullable)
    Public ReadOnly Property HistoryAuthorID(index As Integer) As Guid?
      Get
        Return VAProxy.Profile.HistoryAuthorID(index)
      End Get
    End Property

    ' Returns author flag HistoryAuthorTag1
    Public ReadOnly Property HistoryAuthorTag1 As String
      Get
        Return VAProxy.Profile.HistoryAuthorTag1()
      End Get
    End Property

    ' Returns author flag HistoryAuthorTag2
    Public ReadOnly Property HistoryAuthorTag2 As String
      Get
        Return VAProxy.Profile.HistoryAuthorTag2()
      End Get
    End Property

    ' Returns author flag HistoryAuthorTag3
    Public ReadOnly Property HistoryAuthorTag3 As String
      Get
        Return VAProxy.Profile.HistoryAuthorTag3()
      End Get
    End Property

    ' Returns author flag HistoryProductID (nullable) - UNDOCUMENTED
    Public ReadOnly Property HistoryProductID(index As Integer) As Guid?
      Get
        Return VAProxy.Profile.HistoryProductID()
      End Get
    End Property

    ' Attempts to switch to the profile named in ProfileName
    Overloads Public Sub SwitchTo(ProfileName As String)
      VAProxy.Profile.SwitchTo(ProfileName)
    End Sub

    ' Attempts to switch to the profile with internal ID identified in ProfileID
    Overloads Public Sub SwitchTo(ProfileID As Guid)
      VAProxy.Profile.SwitchTo(ProfileID)
    End Sub

    ' Reloads the currently active profile (stops execution prior to load)
    Public Sub Reset()
      VAProxy.Profile.Reset()
    End Sub

    ' Saves the currently active profile with all changes - UNDOCUMENTED (UNTESTED)
    Public Sub Save()
      VAProxy.Profile.Save()
    End Sub

  End Class
End Namespace
