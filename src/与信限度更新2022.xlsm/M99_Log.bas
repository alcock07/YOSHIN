Attribute VB_Name = "M99_Log"
Option Explicit

Public Const APP_NAME = "�^�M���x�X�V2022"

Sub Open_Log(strCPN As String)
Dim strLOG As String
Dim boolA  As Boolean
    strLOG = Format(Now(), "yyyy/mm/dd") & " " & Format(Now(), "hh:mm:ss") & " -" & strCPN & "- " & APP_NAME & "�FStart"
    boolA = AddText("\\192.168.128.4\os\admin\alcock.Log", strLOG)
End Sub

Sub Close_Log(strCPN As String)
Dim strLOG As String
Dim boolA As Boolean
    strLOG = Format(Now(), "yyyy/mm/dd") & " " & Format(Now(), "hh:mm:ss") & " -" & strCPN & "- " & APP_NAME & "�FEnd"
    boolA = AddText("X:\admin\alcock.Log", strLOG)
End Sub

Public Function AddText(FName As String, txt As String) As Boolean
'=============================
'÷��̧�ْǉ�
'FName : �o��̧�ٖ�
'txt   : �o��÷��
'=============================
    Dim iFNW
    On Error Resume Next
    iFNW = FreeFile
    Open FName For Append As iFNW
        Print #iFNW, txt
    Close iFNW
End Function

Public Function WriteText(FName As String, txt As String) As Boolean
'==============================
'÷��̧�ُ�����
'FName : �o��̧�ٖ�
'txt   : �o��÷��
'==============================
    Dim iFNW
    On Error Resume Next
    iFNW = FreeFile
    Open FName For Output As iFNW
        Print #iFNW, txt
    Close iFNW
End Function


