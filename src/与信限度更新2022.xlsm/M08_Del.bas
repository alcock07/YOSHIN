Attribute VB_Name = "M08_Del"
Option Explicit

Sub Del_Data()

    Dim cnW  As New ADODB.Connection
    Dim rsA  As New ADODB.Recordset
    Dim strSQL  As String

    '与信限度データ整理
    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW
    cnW.Open
    rsA.Open "与信限度データ", cnW, adOpenStatic, adLockPessimistic
    If rsA.EOF = False Then rsA.MoveFirst
    Do Until rsA.EOF
        '回収日超過チェック
        If (rsA.Fields("締め日") < rsA.Fields("支払日")) Or rsA.Fields("サイクル") > "01" Then
            rsA.Fields("超過") = "Y"
        End If
        '締め日月末置換
        If rsA.Fields("締め日") = "99" Then rsA.Fields("締め日") = "末"
        '回収日月末置換
        If rsA.Fields("支払日") = "99" Then rsA.Fields("支払日") = "末"
        rsA.Update
        rsA.MoveNext
    Loop
    
    rsA.Close
    
    '計画のみ削除
    strSQL = ""
    strSQL = strSQL & "   DELETE "
    strSQL = strSQL & "   FROM "
    strSQL = strSQL & "       与信限度データ "
    strSQL = strSQL & "   WHERE "
    strSQL = strSQL & "       与信限度データ.得意先コード > '0000000900000'"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    '倒産会社削除
    strSQL = ""
    strSQL = strSQL & "   DELETE "
    strSQL = strSQL & "   FROM"
    strSQL = strSQL & "       与信限度データ "
    strSQL = strSQL & "   WHERE "
    strSQL = strSQL & "       与信限度データ.得意先コード = '0000000210850'"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    'グループ間削除
    strSQL = ""
    strSQL = strSQL & "   DELETE "
    strSQL = strSQL & "   FROM "
    strSQL = strSQL & "       与信限度データ "
    strSQL = strSQL & "   WHERE "
    strSQL = strSQL & "       与信限度データ.GCODE = '0000000819001'"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    '債権なし削除
    strSQL = ""
    strSQL = strSQL & "   DELETE "
    strSQL = strSQL & "   FROM "
    strSQL = strSQL & "       与信限度データ "
    strSQL = strSQL & "   WHERE "
    strSQL = strSQL & "       [与信限度データ]![売掛残]+[与信限度データ]![手形債権] <=0"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    '与信対象外削除
    strSQL = ""
    strSQL = strSQL & "   DELETE "
    strSQL = strSQL & "   FROM "
    strSQL = strSQL & "       与信限度データ "
    strSQL = strSQL & "   WHERE "
    strSQL = strSQL & "       与信限度データ.与信限度額 = 999999999999"
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

