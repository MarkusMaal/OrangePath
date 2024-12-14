# Useful shapes

There are some useful shapes in Sunlight OS you may want to interact with directly. Here's a list of them. All of them can be accessed from Slide1 object.

## AppID

This shape is hidden by default, but can be revealed by the user when enabling "Debug mode" in settings. Contains AppID for the most recently opened app.

## MoveEvent and ResizeEvent

These shapes either contain the text "False", "True" or "N/A" and are used specifically for detecting when the user is resizing or moving a window. Enable "Debug mode" in settings and see how their values change when moving or resizing windows. "N/A" value is displayed when the user attempts to resize or move windows when a modal dialog is open.

## AppCreatingEvent

This shape contains the name of the most recently opened app.

## Username

Name of the currently logged in user. For guest session, the value will be "Guest" and if nobody is logged in, the value will be "Nobody".

## RegularApp:<AppID>

This is a group that contains an entire app window. Useful to know about for performing bulk actions on shapes.

## TaskIcon:<AppID>

Label for a specific app on the taskbar.

## ITaskIcon:<AppID>

Icon for a specific app on the taskbar.

## TaskbarButtonSample

Sample of how a taskbar label should be sized and formatted.

## Clock

Displays current clock time. Use UpdateTime macro to refresh it.

## WaitPlease

This group contains the load indicator along with its label.

## AxTextBox

ActiveX Text Box control that is shown dynamically depending on what window is focused.

## AxComboBox

ActiveX Combo Box control that can be shown programmatically when required. Here's an example:

```VB
    Slide1.AxComboBox.Clear
    Slide1.AxComboBox.AddItem ("ComboBox item 1")
    Slide1.AxComboBox.AddItem ("ComboBox item 2")
    Slide1.AxComboBox.AddItem ("ComboBox item 3")
    Slide1.AxComboBox.AddItem ("ComboBox item 4")
    Slide1.AxComboBox.Visible = True
    ActivePresentation.SlideShowWindow.View.GotoSlide (4)
    Slide1.AxComboBox.Left = Slide1.Shapes("ComboBoxTriggerAppHello:" & AppID).Left
    Slide1.AxComboBox.Top = Slide1.Shapes("ComboBoxTriggerAppHello:" & AppID).Top
    Slide1.AxComboBox.Width = 200
    Slide1.AxComboBox.DropDown
    Slide1.AxComboBox.Visible = False
```
