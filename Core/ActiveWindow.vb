Namespace Global.JDS.VAProxyWrapper.Core
  Public Class ActiveWindow
    Private ReadOnly Property vaProxy As Object
    Private ReadOnly Property Parent As VAProxy
    
    Public Sub New(proxy As Object, vaParent As VAProxy)
      ' Dynamically typecast the proxy as one of the VoiceAttackInitProxyClass types
      Dim isProxyType As Type = proxy.GetType
      ' And save reference to the passed proxy and parent 
      _vaProxy = Convert.ChangeType(proxy, isProxyType)
      _Parent = vaParent
    End Sub

    ' Returns the Active Window Title
    Public ReadOnly Property Title As String
      Get
        Return vaProxy.Utility.ActiveWindowTitle()
      End Get
    End Property

    ' Returns the Active Window Process Name
    Public ReadOnly Property ProcessName As String
      Get
        Return vaProxy.Utility.ActiveWindowProcessName()
      End Get
    End Property

    ' Returns the Active Window Process ID
    Public ReadOnly Property ProcessID As Integer
      Get
        Return vaProxy.Utility.ActiveProcessID()
      End Get
    End Property

'    ActiveWindowPath() - This string function returns the path of the active window's executable. 
'    ActiveWindowWidth() - Returns the active window's width as an integer.  Helps with resizing/moving. 
'    ActiveWindowHeight() - Returns the active window's height as an integer.  Helps with resizing/moving. 
'    ActiveWindowLeft() - Returns the active window's left (X coordinate) as an integer.  Helps with resizing/moving. 
'    ActiveWindowTop() - Returns the active window's top (Y coordinate) as an integer.  Helps with resizing/moving. 
'    ActiveWindowRight() - Returns the active window's right (left + width) as an integer.  Helps with resizing/moving. 
'    ActiveWindowBottom() - Returns the active window's bottom (top + height) as an integer.  Helps with resizing/moving. 
    
'    ProcessExists(String ProcessName) – This Boolean function returns true if a process with the name specified in textVariable exists.  Returns false if not.  Note that this can take wildcards (*).  For instance, if ProcessName contains '*notepad*' (without quotes), the search will look for a process that has a name that contains 'notepad'. 
'    ProcessCount(String ProcessName) – Returns an integer value of 1 or more if processes with the name specified in ProcessName exist.  Returns 0 if there are none.  Note that this can take wildcards (*).  For instance, if ProcessName contains '*notepad*' (without quotes), the search will look for processes that have a name that contains 'notepad'. 
'    ProcessForeground(String ProcessName) – This Boolean function returns true if a process with a main window with the title specified in ProcessName is the foreground window.  Returns false if not.  Note that this can take wildcards (*) for window titles that change.  For instance, if ProcessName contains '*notepad*' (without quotes), the search will look for a process that has a name that contains 'notepad'. 
'    ProcessMinimized(String ProcessName) – Returns a Boolean value of true if a process with a main window with the title specified in ProcessName is minimized.  Returns false if not.  Note that this can take wildcards (*) for window titles that change.  For instance, if ProcessName contains '*notepad*' (without quotes), the search will look for a process that has a name that contains 'notepad'. 
'    ProcessMaximized(String ProcessName) - Returns a Boolean value of true if a process with a main window with the title specified in ProcessName is maximized.  Returns false if not.  Note that this can take wildcards (*) for window titles that change.  For instance, if ProcessName contains '*notepad*' (without quotes), the search will look for a process that has a name that contains 'notepad'. 

'    WindowExists(String WindowName) - Returns a Boolean value of true if a window with the title specified in WindowName exists.  Returns false if not.  Note that this can take wildcards (*) for window titles that change.  For instance, if WindowName contains '*notepad*' (without quotes), the search will look for a window that has a title that contains 'notepad'. 
'    WindowForeground(String WindowName) - Returns a Boolean value of true if a window with the title specified in WindowName is the foreground window.  Returns false if not.  Note that this can take wildcards (*) for window titles that change.  For instance, if WindowName contains '*notepad*' (without quotes), the search will look for a window that has a title that contains 'notepad'. 
'    WindowMinimized(String WindowName) - Returns a Boolean value of true if a window with the title specified in WindowName is minimized.  Returns false if not.  Note that this can take wildcards (*) for window titles that change.  For instance, if textVariable contains '*notepad*' (without quotes), the search will look for a window that has a title that contains 'notepad'. 
'    WindowMaximized(String WindowName) - Returns a Boolean value of true if a window with the title specified in WindowName is maximized.  Returns false if not.  Note that this can take wildcards (*) for window titles that change.  For instance, if WindowName contains '*notepad*' (without quotes), the search will look for a window that has a title that contains 'notepad'. 
'    WindowCount(String WindowName) - Returns an integer value of 1 or more if windows with the title specified in WindowName exist.  Returns 0 if there are none.  Note that this can take wildcards (*) for window titles that change.  For instance, if WindowName contains '*notepad*' (without quotes), the search will look for windows that have a title that contains 'notepad'. 
'    WindowTitleUnderMouse() - Returns the title for the window that is currently located under the mouse as a string. 
'    WindowProcessUnderMouse() - Returns the process name for the window that is currently located under the mouse as a string. 

'    CommandTargetForeground() - Returns a Boolean value of true if the current command’s target is the foreground window.  Returns false if the target is not the foreground window or if the target does not exist. 
'    CommandTargetMinimized() - Returns a Boolean value of true if the current command’s target is minimized.  Returns false if the target is not minimized or if the target does not exist. 
'    CommandTargetMaximized() - Returns a Boolean value of true if the current command’s target is maximized.  Returns false if the target is not maximized or if the target does not exist. 

'    MousePositionScreenX() - Returns the X coordinate of the mouse position as it relates to the screen as an integer.  Zero would be the top-left corner of the screen. 
'    MousePositionScreenY() - Returns the Y coordinate of the mouse position as it relates to the screen as an integer.  Zero would be the top-left corner of the screen. 
'    MousePositionWindowX() - Returns the X coordinate of the mouse position as it relates to the active window as an integer.  Zero would be the top-left corner of the window. 
'    MousePositionWindowY() - Returns the Y coordinate of the mouse position as it relates to the active window as an integer.  Zero would be the top-left corner of the window. 

'    CapsLockOn() - Returns a Boolean value of true if the caps lock key is locked. 
'    NumLockOn() - Returns a Boolean value of true if the numlock key is locked. 
'    ScrollLockOn() - Returns a Boolean value of true if the scroll lock key is locked.  

  End Class
End Namespace