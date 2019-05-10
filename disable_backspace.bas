Attribute VB_Name = "NewMacros"
Public backspace_disabled As Boolean


Sub disable_backspace()

KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyB, wdKeyShift, wdKeyControl), KeyCategory:=wdKeyCategoryMacro, Command:="print_hello_world"
KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyBackspace), KeyCategory:=wdKeyCategoryMacro, Command:="empty_function"


End Sub


Public Sub print_hello_world()
Dim aKey As KeyBinding

For Each aKey In KeyBindings
   If aKey.KeyCode = BuildKeyCode(wdKeyBackspace) Then
    If backspace_disabled Then
    MsgBox ("disable_backspace")
    'KeyBindings.Add backspace_fake
    aKey.Rebind KeyCategory:=wdKeyCategoryMacro, Command:="empty_function"
    
    Else
    'FindKey(BuildKeyCode(wdKeyBackspace)).Execute
    MsgBox ("enable_backsapc")
    'MsgBox ("bye")
    aKey.Disable
    End If
   End If
   
   
   

Next aKey

backspace_disabled = Not backspace_disabled

End Sub

Public Sub empty_function()

End Sub






