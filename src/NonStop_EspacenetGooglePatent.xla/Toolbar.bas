Attribute VB_Name = "Toolbar"
Option Explicit

Sub MakeToolBar()
    Dim myBar As CommandBar:   Set myBar = Application.CommandBars.Add(Name:="NonStop_Espacenet", Position:=msoBarFloating)
    myBar.Visible = True

    '=========================================================
    Dim myButton1 As CommandBarControl: Set myButton1 = myBar.Controls.Add(Type:=msoControlButton, ID:=1)
    With myButton1
        .Style = msoButtonIconAndCaption
        .OnAction = "VAMIE_espacenet"
        .FaceId = 84
        .Caption = "spacenet"
    End With
    
    Dim myButton2 As CommandBarControl: Set myButton2 = myBar.Controls.Add(Type:=msoControlButton, ID:=2)
    With myButton2
        .Style = msoButtonIconAndCaption
        .OnAction = "VAMIE_googlePatent"
        .FaceId = 86
        .Caption = "ooglePatent"
    End With
    
End Sub


Sub RemoveToolBar()
    On Error Resume Next
        Application.CommandBars("NonStop_Espacenet").Delete
    On Error GoTo 0
End Sub
