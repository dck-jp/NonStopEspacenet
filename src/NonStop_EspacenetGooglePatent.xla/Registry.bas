Attribute VB_Name = "Registry"
#If VBA7 Then
Private Declare PtrSafe Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
          (ByVal hKey As LongPtr, ByVal lpSubKey As String, ByVal ulOptions As LongPtr, _
          ByVal samDesired As LongPtr, phkResult As LongPtr) As LongPtr
          
Private Declare PtrSafe Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As LongPtr

Private Declare PtrSafe Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
          (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal Reserved As LongPtr, _
          ByVal dwType As LongPtr, ByVal lpData As Any, ByVal cbData As LongPtr) As LongPtr
          
Private Declare PtrSafe Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
          (ByVal hKey As LongPtr, ByVal lpValueName As String, ByVal lpReserved As LongPtr, _
          lpType As LongPtr, lpData As Any, lpcbData As LongPtr) As LongPtr
#Else
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
          (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
          ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
          (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, _
          ByVal dwType As Long, ByVal lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
          (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
          lpType As Long, lpData As Any, lpcbData As Long) As Long
#End If

Private Const ERROR_SUCCESS = 0
Private Const REG_SZ = 1
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_CURRENT_CONFIG = &H80000005

Public Function GetRegValue(lngRootKey As LongPtr, strSubKey As String, _
                    strName As String) As String
'�T�v ���W�X�g���̒l���擾����
'���� lngRootKey : ���W�X�g�����[�g�L�[
'     strSubKey  : ���W�X�g���T�u�L�[
'     strName    : ���O
'�Ԓl �擾�������W�X�g���̒l

  Dim lngRet As LongPtr
  Dim hWnd As LongPtr
  Dim strValue As String

  '�n���h�����J��
  'hWnd = Application.hWnd
  lngRet = RegOpenKeyEx(lngRootKey, strSubKey, 0, KEY_QUERY_VALUE, hWnd)
  '�󂯎��l�p�̃o�b�t�@���m��
  strValue = String(255, " ")
  '�l���擾
  lngRet = RegQueryValueEx(hWnd, strName, 0, 0, ByVal strValue, LenB(strValue))
  '�n���h�������
  RegCloseKey hWnd
  
  '�擾�����l����㑱��Null����菜��
  'strValue = Left(strValue, InStr(strValue, vbNullChar) - 1)
  '�擾�����l��Ԃ�l�ɐݒ�
  GetRegValue = strValue

End Function

