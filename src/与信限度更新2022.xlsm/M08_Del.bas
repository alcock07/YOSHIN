Attribute VB_Name = "M08_Del"
Option Explicit

Sub Del_Data()

    Dim cnW  As New ADODB.Connection
    Dim rsA  As New ADODB.Recordset
    Dim strSQL  As String

    '�^�M���x�f�[�^����
    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW
    cnW.Open
    rsA.Open "�^�M���x�f�[�^", cnW, adOpenStatic, adLockPessimistic
    If rsA.EOF = False Then rsA.MoveFirst
    Do Until rsA.EOF
        '��������߃`�F�b�N
        If (rsA.Fields("���ߓ�") < rsA.Fields("�x����")) Or rsA.Fields("�T�C�N��") > "01" Then
            rsA.Fields("����") = "Y"
        End If
        '���ߓ������u��
        If rsA.Fields("���ߓ�") = "99" Then rsA.Fields("���ߓ�") = "��"
        '����������u��
        If rsA.Fields("�x����") = "99" Then rsA.Fields("�x����") = "��"
        rsA.Update
        rsA.MoveNext
    Loop
    
    rsA.Close
    
    '�v��̂ݍ폜
    strSQL = ""
    strSQL = strSQL & "   DELETE "
    strSQL = strSQL & "   FROM "
    strSQL = strSQL & "       �^�M���x�f�[�^ "
    strSQL = strSQL & "   WHERE "
    strSQL = strSQL & "       �^�M���x�f�[�^.���Ӑ�R�[�h > '0000000900000'"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    '�|�Y��Ѝ폜
    strSQL = ""
    strSQL = strSQL & "   DELETE "
    strSQL = strSQL & "   FROM"
    strSQL = strSQL & "       �^�M���x�f�[�^ "
    strSQL = strSQL & "   WHERE "
    strSQL = strSQL & "       �^�M���x�f�[�^.���Ӑ�R�[�h = '0000000210850'"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    '�O���[�v�ԍ폜
    strSQL = ""
    strSQL = strSQL & "   DELETE "
    strSQL = strSQL & "   FROM "
    strSQL = strSQL & "       �^�M���x�f�[�^ "
    strSQL = strSQL & "   WHERE "
    strSQL = strSQL & "       �^�M���x�f�[�^.GCODE = '0000000819001'"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    '���Ȃ��폜
    strSQL = ""
    strSQL = strSQL & "   DELETE "
    strSQL = strSQL & "   FROM "
    strSQL = strSQL & "       �^�M���x�f�[�^ "
    strSQL = strSQL & "   WHERE "
    strSQL = strSQL & "       [�^�M���x�f�[�^]![���|�c]+[�^�M���x�f�[�^]![��`��] <=0"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    '�^�M�ΏۊO�폜
    strSQL = ""
    strSQL = strSQL & "   DELETE "
    strSQL = strSQL & "   FROM "
    strSQL = strSQL & "       �^�M���x�f�[�^ "
    strSQL = strSQL & "   WHERE "
    strSQL = strSQL & "       �^�M���x�f�[�^.�^�M���x�z = 999999999999"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
Exit_DB:

    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If
    
End Sub

