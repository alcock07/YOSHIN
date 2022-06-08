Attribute VB_Name = "M06_Joken"
Option Explicit

Sub Set_JOKEN()

    Dim cnW     As New ADODB.Connection
    Dim rsA     As New ADODB.Recordset
    Dim rsM     As New ADODB.Recordset
    Dim strSQL  As String
    Dim strGCD  As String
    
    '与信限度データの取引条件を更新
    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW
    cnW.Open
    rsA.Open "与信限度データ", cnW, adOpenStatic, adLockPessimistic
    If rsA.EOF = False Then
        rsA.MoveFirst
    End If
    rsA.MoveFirst
    Do Until rsA.EOF
        strCODE = rsA.Fields("得意先コード") & ""
        If strCODE <> "" Then
            '得意先ﾏｽﾀの取引条件を取得
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
                rsA.Fields("締め日") = rsM.Fields("TOKSMEDD")
                rsA.Fields("サイクル") = rsM.Fields("TOKKESCC")
                rsA.Fields("支払日") = rsM.Fields("TOKKESDD")
                If rsM.Fields(3) & "" = "" Then 'UKETEGST00
                    rsA.Fields("サイト") = rsM.Fields(4) 'UKETEGST01
                Else
                    rsA.Fields("サイト") = rsM.Fields(3) 'UKETEGST00
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
