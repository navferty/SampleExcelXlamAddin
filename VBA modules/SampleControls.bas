Attribute VB_Name = "SampleControls"
Option Explicit

Public Sub OnSplitButton1Click(rc As IRibbonControl)
    MsgBox "Split button 1 was clicked"
End Sub

Public Sub OnSplitButton2Click(rc As IRibbonControl)
    MsgBox "Split button 2 was clicked"
End Sub

Public Sub OnToggleButtonClick(rc As IRibbonControl, isButtonPressed As Boolean)
    MsgBox "Toggle button was toggled, button now is " & IIf(isButtonPressed, "pressed", "not pressed")
End Sub

Public Sub OnDropDownSelected(rc As IRibbonControl, selectedItemId As String, selectedItemIndex As Integer)
    MsgBox "DropDown was changed, selected item id is " & selectedItemId
End Sub

Public Sub OnComboBoxSelected(rc As IRibbonControl, comboBoxValue As String)
    MsgBox "Combo box was changed, value is " & comboBoxValue
End Sub

Public Sub GetMenuContent(rc As IRibbonControl, ByRef returnedVal)
    Dim xml As String

    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">" & _
          "<button id=""but1"" imageMso=""Help"" label=""Help"" onAction=""OnHelpPressed""/>" & _
          "<button id=""but2"" imageMso=""FindDialog"" label=""Find"" onAction=""OnFindPressed""/>" & _
          "</menu>"

    returnedVal = xml
End Sub

Public Sub OnCheckBoxToggled(rc As IRibbonControl, isButtonChecked As Boolean)
    MsgBox "Check box was toggled, value is " & IIf(isButtonChecked, "checked", "not checked")
End Sub

Public Sub OnHelpPressed(rc As IRibbonControl)
    MsgBox "Help button pressed"
End Sub

Public Sub OnFindPressed(rc As IRibbonControl)
    MsgBox "Find button pressed"
End Sub
