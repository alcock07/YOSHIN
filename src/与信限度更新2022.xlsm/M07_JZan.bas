Attribute VB_Name = "M07_JZan"
Option Explicit

Private Const MYPROVIDERE = "Provider=SQLOLEDB;"
Private Const MYSERVER = "Data Source=192.168.128.9\SQLEXPRESS;"
Private Const USER = "User ID=sa;"
Private Const PSWD = "Password=ALCadmin!;"

Sub J_ZAN()

    Dim cnW     As New ADODB.Connection
    Dim cnJ     As New ADODB.Connection
    Dim rsA     As New ADODB.Recordset
    Dim rsJ     As New ADODB.Recordset
    Dim strSQL  As String
    Dim strCD   As String
    Dim strNT   As String
    
    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW
    cnW.Open
    strNT = "Initial Catalog=process_os;"
    cnJ.ConnectionString = MYPROVIDERE & MYSERVER & strNT & USER & PSWD
    cnJ.Open
    
    Dim SQLS1 As String
    SQLS1 = ""
    SQLS1 = SQLS1 & "   SELECT "
    SQLS1 = SQLS1 & "       得意先コード"
    SQLS1 = SQLS1 & "       , 受注残 "
    SQLS1 = SQLS1 & "   FROM "
    SQLS1 = SQLS1 & "       与信限度データ"
    rsA.Open SQLS1, cnW, adOpenStatic, adLockPessimistic
    If rsA.EOF Then GoTo Exit_DB
    
    '与信限度データの受注残を更新
    rsA.MoveFirst
    Do Until rsA.EOF
        strCD = rsA.Fields("得意先コード") & ""
        If strCD <> "" Then
            strSQL = ""
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & "     Sum(zankn) "
            strSQL = strSQL & " FROM "
            strSQL = strSQL & "     JUZTBZ_Hybrid "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & "     tokcd = '" & strCD & "'"
            rsJ.Open strSQL, cnJ, adOpenStatic, adLockReadOnly
            If rsJ.EOF = False Then
                rsA.Fields("受注残") = rsJ.Fields(0)
                rsA.Update
            End If
            rsJ.Close
        End If
        rsA.MoveNext
    Loop
    
Exit_DB:
    
    If Not rsJ Is Nothing Then
        If rsJ.State = adStateOpen Then rsJ.Close
        Set rsJ = Nothing
    End If
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not cnJ Is Nothing Then
        If cnJ.State = adStateOpen Then cnJ.Close
        Set cnJ = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If
    
End Sub
