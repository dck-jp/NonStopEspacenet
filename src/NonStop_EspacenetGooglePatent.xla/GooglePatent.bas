Attribute VB_Name = "GooglePatent"
Option Explicit

Sub VAMIE_googlePatent()
    Dim num As String
    Dim cell As Range: For Each cell In Selection
        num = cell
        num = FormatUSNum(num)
        Call OpenBrowser(num)
    Next
End Sub

Private Sub OpenBrowser(ByVal num As String)
On Error GoTo ErrorHandler
        Dim ie As VAMIE
        If Len(num) = 7 Then
            ' Patent Number の場合
            Set ie = New VAMIE
            Call ie.goto_url("http://www.google.com/advanced_patent_search")
            Call ie.type_val("as_pnum", num)
            Call ie.submit_click("btnG", "name")
        Else
            ' Application Number の場合
            Set ie = New VAMIE
            Call ie.goto_url("http://www.google.com/advanced_patent_search")
            Call ie.type_val("as_q", num)
            Call ie.submit_click("btnG", "name")
        End If
        
    Dim objWSH: Set objWSH = CreateObject("WScript.Shell")
    objWSH.Run ie.url, 1
    ie.Quit
ErrorHandler:
End Sub

Sub test()
    OpenBrowser "US2010-123"
End Sub

' US Patent/Application Numberだけを抜き出す
' ※ 多様な記載方法に対応するのはめんどくさいため手抜き
Private Function FormatUSNum(num)
    Dim temp As String
    
    temp = num
    temp = Replace(temp, " ", "")
    temp = Replace(temp, "　", "")
    temp = StrConv(temp, vbNarrow)
    temp = Replace(temp, "US", "")
    temp = Replace(temp, "B", "")
    temp = Replace(temp, "A5", "")
    temp = Replace(temp, "A4", "")
    temp = Replace(temp, "A3", "")
    temp = Replace(temp, "A2", "")
    temp = Replace(temp, "A1", "")
    temp = Replace(temp, "A", "")
    temp = Replace(temp, "-", "/")

    FormatUSNum = temp
End Function




