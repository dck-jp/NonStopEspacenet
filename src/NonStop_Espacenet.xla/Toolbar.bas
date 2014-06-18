Attribute VB_Name = "Toolbar"
Option Explicit

Sub MakeToolBar()
    Dim myBar As CommandBar
    
    Dim myButton1 As CommandBarControl
    

    Set myBar = Application.CommandBars.Add( _
        Name:="NonStop_Espacenet", Position:=msoBarFloating)
    myBar.Visible = True

    '=========================================================
    Set myButton1 = myBar.Controls.Add( _
        Type:=msoControlButton, ID:=1)
    With myButton1
        .Style = msoButtonIconAndCaption
        .OnAction = "VAMIE_espacenet"
        .FaceId = 84
        .Caption = "spacenet"
    End With
End Sub


Sub RemoveToolBar()
    On Error Resume Next
        Application.CommandBars("NonStop_Espacenet").Delete
    On Error GoTo 0
End Sub
