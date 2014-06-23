Attribute VB_Name = "Espacenet"
Option Explicit

Sub VAMIE_espacenet()
    Dim num As String
    
    Dim cell As Range: For Each cell In Selection
        num = cell
        num = FormatNum(num)
        Call OpenBrowser(num)
    Next cell
End Sub

Private Sub OpenBrowser(ByVal num As String)
On Error GoTo ErrorHandler
    Dim ie As VAMIE
    Set ie = New VAMIE
    'ie.Visible = False
    Call ie.goto_url("http://worldwide.espacenet.com/numberSearch?locale=en_EP")
    Call ie.type_val("cqlEditBox", num)
    Call ie.submit_click("Submit", "name")
    
    Dim objWSH: Set objWSH = CreateObject("WScript.Shell")
    objWSH.Run ie.url, 1
    ie.Quit
ErrorHandler:
End Sub

Sub test()
    OpenBrowser "JP2010/1234"
End Sub


' NationCode +  Patent/Application Numberを抜き出す
' ※ 多様な記載方法に対応するのはめんどくさいため手抜き
Private Function FormatNum(num)
    Dim temp As String
    
    temp = num
    temp = Replace(temp, " ", "")
    temp = Replace(temp, "　", "")
    temp = StrConv(temp, vbNarrow)
    
    temp = Replace(temp, "B", "")
    temp = Replace(temp, "A5", "")
    temp = Replace(temp, "A4", "")
    temp = Replace(temp, "A3", "")
    temp = Replace(temp, "A2", "")
    temp = Replace(temp, "A1", "")
    temp = Replace(temp, "A", "")
    temp = Replace(temp, "-", "")
    temp = Replace(temp, "/", "")

    FormatNum = temp
End Function



