Imports VA = VoiceAttack
Imports JDS.VAProxyWrapper.Core

Namespace Global.JDS.VAProxyWrapper.Interfaces
  Friend Interface IVoiceAttackProxy

    '*****************************************************************************
    '
    '     A Voice Attack Proxy Abstraction Layer
    '
    '*****************************************************************************
    ' Allowing us to use late-binding for the VAProxyWrapper and still have intellisense. 
    '
    ' The containers and values are listed in the order which they appear in the
    ' VA Help File PDF compiled for 1.8.7+ (API v4)
    '*****************************************************************************

    '*******************************
    '  Default required properties
    '*******************************
    ReadOnly Property ProxyType As String
    ReadOnly Property VAProxy As Object
    ReadOnly Property LogLevel As LogLevel
    ReadOnly Property LogPrefix As String

    ' These two are required calls from within the VA_Invoke1/VA_Exit1 methods
    ' -- This is so the proxy class objects that VA passes are typecast correctly
    ' -- for any processing the user might wish to do
    Sub Invoke(proxy As VA.VoiceAttackInvokeProxyClass)
    Sub [Exit](proxy As Object)
    
    '*******************************
    '    VAProxyWrapper Base Attributes
    '*******************************

    ' Returns the "Context" passed through the proxy to the plugin
    ReadOnly Property Context As String

    ' Returns the "Session State" provided by Voice Attack (we instantiate, so this isn't really necessary)
   ReadOnly Property SessionState As IReadOnlyDictionary(Of String, Object)

    ' Container for getting and setting Voice Attack variables
    ReadOnly Property Variables As Variables

    ' Collection of Voice Attack profile names
    ReadOnly Property ProfileNames As IReadOnlyCollection(Of String)

    ' Collection of **NON-NULL** Voice Attack profile internal IDs (GUIDs)
    ReadOnly Property ProfileInternalIDs As IReadOnlyCollection(Of Guid)

    ' Logging Container - includes WriteToLog(string val, string color), ClearLog()
    ReadOnly Property Log As Log

    ' Sets the Voice Attack window opacity
    Sub Opacity(percentage As Integer)

    ' Versions Container - includes ProxyVersion, VAVersion, IsRelease, IsTrial
    ReadOnly Property Versions As Versions

    ' Paths Container - directory locations (PluginPath(), InstallDir, SoundsDir, AppsDir, AssembliesDir)
    ReadOnly Property Paths As Paths

    ' "Stop all commands" executed/active?
    ReadOnly Property Stopped As Boolean

    ' Clears the "Stop all commands" action flag
    Sub ResetStopFlag()
    
    ' Container for all options information - includes PluginsEnabled, NestedTokensEnabled, AutoProfileSwitchingEnabled
    ReadOnly Property Options As Options

    ' Voice Attack's main window handle
    ReadOnly Property MainWindowHandle As IntPtr

    ' Closes the Voice Attack window
    Sub Close()

    '*******************************
    '   VAProxyWrapper Utility Attributes
    '*******************************

    ' Generates a collection of phrases based on a query 
    ' @param query - phrases query
    ' @param trimSpaces - whether the phrases should be trimmed
    ' @param lowercase - whether to lowercase the elements
    ' @remarks - "good [morning;day;night]" would generate "good morning", "good day", and "good night"
    Function ExtractPhrases(query As String, Optional trimSpaces As Boolean = False, Optional lowercase As Boolean = False) As IReadOnlyCollection(Of String)

    ' Shortcut to get token values from Voice Attack ({TOKEN})
    Function ParseTokens(token As String) As String

    ' Container for all speech information
    ReadOnly Property Speech As Speech
    
    ' ActiveWindow Container - everything about Voice Attack Window/Process, & CapsLockOn, NumLockOn, ScrollLockOn
    ReadOnly Property ActiveWindow As ActiveWindow

    '*******************************
    '   VAProxyWrapper Command Attributes
    '*******************************

    ' Container for all command information
    ReadOnly Property Commands As Commands
    
    ' Container for information about an executing command
    ReadOnly Property Command As Command

    '*******************************
    '   VAProxyWrapper Profile Attributes
    '*******************************

    ReadOnly Property Profile As Profile

    '*******************************
    '   VAProxyWrapper Queue Attributes
    '*******************************

    ReadOnly Property Queue As Queue

    '*******************************
    '    VAProxyWrapper Events & Handlers
    '*******************************

    ReadOnly Property EventHandlers As EventHandlers

  End Interface
End Namespace
