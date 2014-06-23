VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VAMIE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' VAMIE (VBA Auto Mation for Internet Explorer)
'
' LastModified:
'   2014/1/8    + DOM�Z���N�^�g�� (ByName��ǉ�)
'
'   2014/1/7    * Vbaget�V���[�Y�Ƃ��ă��t�@�C��(���U���Ă����R�[�h���W��)
'                   + �v���p�e�B�̐���
'                   + ����R�}���h�̒ǉ�(goto_url_NoWait, Wait, DisableConfirmFunction)
'                   + 64bit�ւ̑Ή�
'
'   2012/6/20   - domselec()�̃G���[�����폜
'   2012/6/18   + domselec()�ɃG���[�����ǉ� �i�f�o�b�O���ɖ���c�����Â炢���߁j
'
' This Class Module:
'   First Release 2011/11/08
'   Created By D*isuke YAMAKWA
'
' SpecialThanks:
'   Excel VBA�̃}�N����, IE���������삵�悤(DOM�Z���N�^�֐���VBA�Ŏ���)
'   http://d.hatena.ne.jp/language_and_engineering/20090710/p1
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private ie As Object

Property Let Visible(x As Boolean)
    ie.Visible = x
End Property
Property Get Visible() As Boolean
    Visible = ie.Visible
End Property

Property Get Document() As Boolean
    Document = ie.Document
End Property

Property Get url() As String
    url = ie.LocationURL
End Property

Sub Class_Initialize()
    Set ie = CreateObject("InternetExplorer.Application")
    ie.Visible = True
End Sub

Sub goto_url(url)
    ie.Navigate url
    waitIE
End Sub

' waitIE�����ނƖ������[�v����悤�ȃy�[�W�΍�
Sub goto_url_NoWait(url)
    ie.Navigate url
End Sub

Sub Quit()
    ie.Quit
End Sub

' $
Function gid(dom_id)
    ' ���F��IE��getElementById��name���Q�Ƃ���
    Set gid = ie.Document.getElementById(dom_id)
End Function

' getElementsByTagName
Function gtn(tag_name)
    Set gtn = ie.Document.getElementsByName(tag_name)
End Function

' ���͂��܂�
Sub type_val(dom_id, val)
    gid(dom_id).value = val
    Sleep 100
End Sub

' ���M�{�^���⃊���N���N���b�N
Sub submit_click(dom_id, Optional type_ As String = "id")
    If type_ = "id" Then
    gid(dom_id).Click
    Else
    gtn(dom_id)(0).Click
    End If
    waitIE
End Sub

' �Ȉ�DOM�Z���N�^
Function domselec(arr)
    Dim parent_obj As Object: Set parent_obj = ie.Document
    Dim child_obj As Object
    Dim cur, continue_flag, dom_id, tag_name, index_num, name_

    cur = 0: continue_flag = True
    
    Do While continue_flag = True
        Select Case arr(cur):
            Case "id"
                dom_id = arr(cur + 1)
                Set child_obj = parent_obj.getElementById(dom_id)
                cur = cur + 2
            Case "tag"
                tag_name = arr(cur + 1)
                index_num = arr(cur + 2)
                Set child_obj = parent_obj.getElementsByTagName(tag_name)(index_num)
                cur = cur + 3
            Case "name"
                name_ = arr(cur + 1)
                index_num = arr(cur + 2)
                Set child_obj = parent_obj.getElementsByName(name_)(index_num)
                cur = cur + 3
        End Select
        
        Set parent_obj = child_obj
        
        If cur > UBound(arr) Then
            continue_flag = False
        End If
    Loop
    
    Set domselec = parent_obj
End Function

' �v�f���N���b�N���܂�
Sub ie_click(dom_id)
    gid(dom_id).Click
    Sleep 100
End Sub


' �`�F�b�N�{�b�N�X�̏�Ԃ��Z�b�g���܂�
Sub set_check_state(dom_id, checked_flag)
    ' ��]�ʂ�̃`�F�b�N��ԂłȂ���΃N���b�N
    If Not (gid(dom_id).Checked = checked_flag) Then
        ie_click dom_id
    End If
End Sub


' �Z���N�g�{�b�N�X�𕶌��x�[�X�őI�����܂�
Sub select_by_label(dom_id, label)
    If Len(label) < 1 Then
      Exit Sub
    End If
    
    Dim opts As Object
    Dim i As Integer
    Set opts = gid(dom_id).Options
    For i = 0 To opts.Length - 1
        ' text��������
        If opts(i).innerText = label Then
            opts(i).Selected = True
            Exit Sub
        End If
    Next i
    
End Sub

' ���W�I�{�^����l�x�[�X�őI�����܂�
Sub select_radio_by_val(post_name, value)
    If Len(value) < 1 Then
        Exit Sub
    End If
    
    Dim radios: Set radios = ie.Document.getElementsByName(post_name)
    Dim i: For i = 0 To radios.Length - 1
        If radios(i).value = CStr(value) Then
            radios(i).Click
            Sleep 100
        End If
    Next i

End Sub
' =======================================================
Sub waitIE()
    Do While ie.Busy = True Or ie.ReadyState <> 4
        DoEvents
    Loop
    
    Sleep 100
End Sub

Sub Wait(millisecond As Integer)
    Sleep millisecond
    DoEvents
End Sub

' =======================================================
Function GetIEVer()
    Dim FS: Set FS = CreateObject("Scripting.FileSystemObject")
    Dim hoge: hoge = Fix(val(FS.GetFileVersion(ie.FullName)))
    GetIEVer = hoge
End Function

'confirm()�Ăяo�����Ɋm�F�_�C�A���O��\�������Ȃ�
Sub DisableConfirmFunction()
    Dim ele: Set ele = ie.Document.createElement("SCRIPT")
    
    ele.Type = "text/javascript"
    ele.Text = "function confirm() { return true; }"
    
    Call ie.Document.body.appendChild(ele)
End Sub
