Attribute VB_Name = "M06_Joken"
Option Explicit

Sub Set_JOKEN()

    Dim cnW     As New ADODB.Connection
    Dim rsA     As New ADODB.Recordset
    Dim rsM     As New ADODB.Recordset
    Dim strSQL  As String
    Dim strGCD  As String
    
    '�^�M���x�f�[�^�̎���������X�V
    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW
    cnW.Open
    rsA.Open "�^�M���x�f�[�^", cnW, adOpenStatic, adLockPessimistic
    If rsA.EOF = False Then
        rsA.MoveFirst
    End If
    rsA.MoveFirst
    Do Until rsA.EOF
        strCODE = rsA.Fields("���Ӑ�R�[�h") & ""
        If strCODE <> "" Then
            '���Ӑ�Ͻ��̎���������擾
            strSQL = ""
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & "     TOKSMEDD"
            strSQL = strSQL & "     , TOKKESCC"
            strSQL = strSQL & "     , TOKKESDD"
            strSQL = strSQL & "     , Trim(UKETEGST00)"
            strSQL = strSQL & "     , Trim(UKETEGST01)"
            strSQL = strSQL & "     , Trim(LMTCD) "
            strSQL = strSQL & " FROM "
            strSQL = strSQL & "     TOKMTA "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & "     TOKCD = '" & strCODE & "'"
            rsM.Open strSQL, cnW, adOpenStatic, adLockReadOnly
            If rsM.EOF = False Then
                rsA.Fields("���ߓ�") = rsM.Fields("TOKSMEDD")
                rsA.Fields("�T�C�N��") = rsM.Fields("TOKKESCC")
                rsA.Fields("�x����") = rsM.Fields("TOKKESDD")
                If rsM.Fields(3) & "" = "" Then 'UKETEGST00
                    rsA.Fields("�T�C�g") = rsM.Fields(4) 'UKETEGST01
                Else
                    rsA.Fields("�T�C�g") = rsM.Fields(3) 'UKETEGST00
                End If
                strGCD = rsM.Fields(5) & "" 'LMTCD
                If strGCD = "" Then
                    rsA.Fields("GCODE") = strCODE
                Else
                    rsA.Fields("GCODE") = strGCD
                End If
                rsA.Update
            End If
            rsM.Close
        End If
        rsA.MoveNext
    Loop
    
Exit_DB:
    
    If Not rsM Is Nothing Then
        If rsM.State = adStateOpen Then rsM.Close
        Set rsM = Nothing
    End If
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If
    
End Sub
