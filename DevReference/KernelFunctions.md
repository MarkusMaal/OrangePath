# Kernel functions
OrangePath offers various functions built right into the kernel. These are available in the `OPMain` module.

## Helper functions

### GetAppID(Shp As Shape) As String
This function should be used, when trying to get the ID of the current window. An App ID is used to separate different open application or even various instances of the same application. This should be invoked in a shape macro, like this:
```VB
	' Not a real macro, just an example, you can use any name
	Sub Anything(Shp As Shape)
		Dim AppID As String
		AppID = GetAppID(Shp)
	End Sub
```


### SetVar(Key As String, Value As String)
Allows you to set a global variable, which is visible to any macro by calling the CheckVars macro.

Usage:
```VB
	SetVar "Message", "Hello!"
```

### UnsetVar(Key As String)
Allows you to unset a global variable. It is good practice to clear any global variable once you don't
need it to avoid any conflicts between applications.

Usage:
```VB
	UnsetVar "Message"
```

### CheckVars(str As String)
Allows you to convert any variable references inside a string to their actual value.

You may use this to include the variable value inside a string like so...

```VB
	Dim FullMessage As String
	FullMessage = CheckVars("Your name is %name%!")
```

Or you may just get the value of the variable like so...

```VB
	Dim Name As String
	Name = CheckVars("%name%")
	' If the variable is unset, set it to be an empty string
	If Name = "%name%" Then Name = ""
```



### SaveSysConfig(Key As String, Value As String)
Allows you to save specific system configuration. This directly saves your value as a stream to `/System/Settings.cnf`. Difference from SetVar is that this is used by the system to save certain global settings (such as Autologin) and is not recommended to be used as a general purpose variable storage.

This example disables autologin:
```VB
	SaveSysConfig "Autologin", "Nobody"
```

### GetSysConfig(Name As String) As String
Returns a system configuration with a key specified. If access is denied or the file doesn't exist, a star `*` will be returned.

This example gets the user account name, which is automatically logged into (returns "Nobody" if autologin is disabled):
```VB
	GetSysConfig "Autologin"
```


### Pause (Length As LongPtr)
Pauses execution for the number of seconds specied.

### IsInArray (stringToBeFound As String, arr() As String) As Boolean
Returns True or False depending if a string exists in String array.

Usage:
```VB
	Dim inArray As Boolean
	inArray = IsInArray("Any text", stringArr)
```

### ShapeExists(oSl As Slide, ShapeName As String) As Boolean
Returns True or False depending if a shape exists on a specified slide.

Usage:
```VB
	Dim hasShape As Boolean
	hasShape = ShapeExists(Slide1, "WindowAppNotes:" & AppID)
```

### GroupItemExists(oSl As Shape, ShapeName As String) As Boolean
Returns True or False depending if a shape is a group inside a shape specified.

Usage:
```VB
	Dim inGroup As Boolean
	inGroup = ShapeExists(Slide1.Shapes("RegularApp:" & AppID), "ExampleAppTest:" & AppID)
```

### UpdateTime
Restarts a clock, which is not changing its value.

### InvertValue(Shp As Shape)
This is a shape macro. If invoked (on mouse click or hover), it changes the text of the shape to one of the following:
- If the text on shape is "True", it sets the textframe text to "False"
- Otherwise, it sets the textframe text to "True"


## Window management

### CreateNewWindow
Creates a new window with various parameters specified. Creates a taskbar icon and automatically moves windows
to new workspaces if more than 5 non-modal windows are in a workspace.

Specifying parameters:
```VB
	Slide2.Shapes("AppCreatingEvent").TextFrame.TextRange.Text = "<app_name>"
	Slide2.Shapes("AppID").TextFrame.TextRange.Text = "<app_id>"
```

Generally speaking, you don't need to mess with AppID and changing it to a lower value than it already is can cause
system instability. AppCreatingEvent allows you to specify which application or modal dialog to open. This function
alone does not allow you to specify startup scripts, see "Creating new apps" under DevTools documentation.

The function is called like so:
```VB
	CreateNewWindow
```

### CloseWindow(Shp As Shape)
This is a shape macro, meaning it is supposed to be attached to a shape. Running this macro will close
a window, which that shape is attached to. To use it, add it as an action to a shape, which is inside a
window.

### MovableWindow(Shp As Shape)
This is a shape macro. If added as a "On Hover" action to a shape, it will move the entire window.

### ResizingWindow(Shp As Shape)
This is a shape macro. If added as a "On Hover" action to a shape, it will resize the entire window.

### CheckSize(Shp As Shape)
This is a shape macro. If added as a "On Click" action on a shape, it will switch current workspace.

### CleanPopups
Closes all windows. Executed automatically on system restart.

### CheckUncheck
**TODO**: Move to UI components documentation

This is a shape macro. If addad as a "On Click" action on a checkbox shape, it will check/uncheck it.

### FocusWindow(AppID As String)
This cannot be used a shape macro. Focuses a window with an AppID specified.

Example usage:
```VB
	FocusWindow "1"
```

### MinimizeWindow(Shp As Shape)
This is a shape macro. If invoked on click or hover from a shape inside a window, it minimizes that window.

### MinimizeRestore(Shp As Shape)
This is a shape macro. If invoked from a shape (on click or hover) which has a specific AppID attached to its name (e.g. TaskIcon:17), it does one of the following:
- If the window corresponding to AppID is not minimized and not in focus, it focuses the window
- If the window corresponding to AppID is not minimized, but is in focus, it minimizes the window
- If the window corresponding to AppID is minimized, it restores the window and gives it focus

### PasteToGroup(Ref As Shape, Shp As Shape, Name As String, OffsetX As Integer, OffsetY As Integer, Sld As Slide, Optional Macro As String = "")
This macro allows you to paste a shape into the application window. Here's an examplanation of all the parameters, you may need to use:
- **Ref** is a reference shape, which an App ID is extracted from. If you know the AppID, you may use the whole window as a reference shape, e.g. `Slide1.Shapes("RegularApp:" & AppID)`
- **Shp** is the shape you want to paste into the window
- **Name** is a name, which you want to give to the newly pasted shape. The name must be in the format `<AnythingInPascalCase>App<AppName>:<AppID>`, e.g. `CustomShapeAppTest:13`
- **OffsetX** is the horizontal offset you want to use for the shape (from the top edge of the slide)
- **OffsetY** is the vertical offset you want to use for the shape (from the left edge of the slide)
- **Sld** is the slide where the window is located at, e.g. `Slide1`
- **Macro** is an optional parameter, it specifies which macro should be run when this pasted shape is clicked

Full example
```VB
	PasteToGroup Slide1.Shapes("RegularApp:" & AppID), "CustomShapeAppTest:" & AppID, Slide1.Shapes("RegularApp:" & AppID).Left + 20, Slide1.Shapes("RegularApp:" & AppID).Top + 50, Slide1
```

### EraseFromGroup(Ref As Shape, ShpName As String, Sld As Slide)
This macro allows you to delete a shape from the application window. Here's an explanation of all the parameters, you have to use:
- **Ref** is a reference shape, which an App ID is extracted from. If you know the AppID, you may use the whole window as a reference shape, e.g. `Slide1.Shapes("RegularApp:" & AppID)`
- **ShpName** is the name of the shape you wish to delete from the group. The name must be in the format `<AnythingInPascalCase>App<AppName>:<AppID>`, e.g. `CustomShapeAppTest:13`
- **Sld** is the slide where the window is located at, e.g. `Slide1`

Full example
```VB
	EraseFromGroup Slide1.Shapes("RegularApp:" & AppID), "CustomShapeAppTest:" & AppID, Slide1
```

## Power management

### Hibernate
This macro takes no arguments. If invoked, it hibernates the system, which saves the current state and
allows the user to come back later to resume exactly where they left off.

### Restart
This macro takes no arguments. If invoked, it restarts the system.

### RestartRecovery
This macro takes no arguments. If invoked, it restarts the system to recovery mode, which allows the user to
reset factory settings and apply system updates.

### Slide2Run
This macro executes while displaying the startup screen.

### SavePresentation
Saves the current presentation and takes no arguments. Does not work if system is in shut down mode.

Usage
```VB
	SavePresentation
```