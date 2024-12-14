# Event hooks

Sunlight OS provides a few hooks for detecting certain user actions, such as minimizing and restoring a window. Every subroutine that gets called this way has to be named in the format `App<AppName><HookName>`, e.g. `AppSettingsRestore`. In addition, every such routine takes `AppID As String` as the argument.

Example:

```VB
' Run this routine after restoring the window
Sub AppSettingsRestore(AppID As String)
    AppSettingsSwitchCat Slide1.Shapes("CatPersonaliseAppSettings:" & AppID)
End Sub
```

## Restore

This hook is triggered after the user restores the application window from a minimized state. This is handy, as by default, Sunlight OS makes every shape visible when you restore a window, so it may be neccessary to manually specify which shapes you want to make invisible again.

## Minimize

This hook is triggered after the application window gets minimized.

## Focus

This hook is triggered when the application window gets focus. The focus can be given if the user drags an inactive application window, when clicking on the taskbar label/icon or when triggering the `FocusWindow(AppID)` macro programmatically.

## SizeChanged

This hook is triggered after the user has finished resizing the window. This is handy when you want to make responsive interfaces.

## Custom hooks

A custom hook can be triggered by calling this macro:

```VB
	TryRunMacro <AppName As String>, <HookName As String>, <AppID As Integer>
```