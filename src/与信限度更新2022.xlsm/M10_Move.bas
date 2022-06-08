Attribute VB_Name = "M10_Move"
Option Explicit

Sub Move_Data()

'===================
'与信限度データ移動
'===================

    Dim cnW    As New ADODB.Connection
    Dim rsA    As New ADODB.Recordset
    Dim strSQL As String
    
    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW
    cnW.Open
    
    'データ参照DBクリア
    strSQL = ""
    strSQL = strSQL & "DELETE"
    strSQL = strSQL & "       FROM 与信限度Link"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    'データ移行
    strSQL = ""
    strSQL = strSQL & "INSERT INTO 与信限度Link"
    strSQL = strSQL & "       SELECT 与信限度データ.*"
    strSQL = strSQL & "       FROM 与信限度データ"
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

