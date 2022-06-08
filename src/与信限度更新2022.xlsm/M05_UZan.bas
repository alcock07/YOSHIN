Attribute VB_Name = "M05_UZan"
Option Explicit

Sub ZAN_Change()
    
    '得意先ﾏｽﾀの売掛残を取得して
    '売掛残以降の売掛金額を売掛ｻﾏﾘから取得
    
    Dim cnW     As ADODB.Connection
    Dim rsS     As ADODB.Recordset
    Dim rsA     As ADODB.Recordset
    Dim rsM     As ADODB.Recordset
    Dim strSQL  As String
    Dim strZ    As String
    Dim lngZ    As Double
    Dim strDate As String
    Dim dNOW    As Date
    Dim lngYY   As Long
    Dim lngMM   As Long
        
    '当月末生成
    dNOW = Date
    lngYY = Strings.Format(dNOW, "yyyy")
    lngMM = Strings.Format(dNOW, "mm")
    lngMM = lngMM + 1
    If lngMM = 13 Then
        lngMM = 1
        lngYY = lngYY + 1
    End If
    dNOW = Strings.Format(lngYY, "0000") & "/" & Strings.Format(lngMM, "00") & "/01"
    strDate = Strings.Format(dNOW - 1, "yyyymmdd")
    
    'データベース設定
    Set cnW = New ADODB.Connection
    Set rsS = New ADODB.Recordset
    Set rsA = New ADODB.Recordset '与信限度データ
    Set rsM = New ADODB.Recordset
    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW
    cnW.Open
    rsA.Open "与信限度データ", cnW, adOpenStatic, adLockPessimistic
    
    If rsA.EOF Then
        GoTo Exit_DB
    End If
    
    '与信限度データの売掛残を更新
    rsA.MoveFirst
    Do Until rsA.EOF
        strCODE = rsA.Fields("得意先コード") & ""
        If strCODE <> "" Then
            '得意先ﾏｽﾀの売掛残を取得
            strSQL = ""
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & "     TOKCD"
            strSQL = strSQL & "     , SMAZANDT"
            strSQL = strSQL & "     , SMAZANKN "
            strSQL = strSQL & " FROM "
            strSQL = strSQL & "     TOKMTA "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & "     TOKCD = '" & strCODE & "'"
            rsM.Open strSQL, cnW, adOpenStatic, adLockReadOnly 'TOKMTA
            If rsM.EOF Then
                strZ = ""
                lngZ = 0
            Else
                strZ = rsM.Fields("SMAZANDT")
                lngZ = rsM.Fields("SMAZANKN")
            End If
            rsM.Close
            '得意先ﾏｽﾀの売掛残以降の売掛金額を取得
            strSQL = ""
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & "     SMADT"
            strSQL = strSQL & "     , Sum([SMAURIKN00]"
            strSQL = strSQL & "             +[SMAURIKN01]"
            strSQL = strSQL & "             +[SMAURIKN02]"
            strSQL = strSQL & "             +[SMAURIKN03]"
            strSQL = strSQL & "             +[SMAURIKN04]"
            strSQL = strSQL & "             +[SMAURIKN05]"
            strSQL = strSQL & "             +[SMAURIKN06]"
            strSQL = strSQL & "             +[SMAURIKN07]"
            strSQL = strSQL & "             +[SMAURIKN08]"
            strSQL = strSQL & "             +[SMAURIKN09]"
            strSQL = strSQL & "             +[SMAUZEKN]"
            strSQL = strSQL & "             +[SMAUZKKN]"
            strSQL = strSQL & "             -[SMANYUKN00]"
            strSQL = strSQL & "             -[SMANYUKN01]"
            strSQL = strSQL & "             -[SMANYUKN02]"
            strSQL = strSQL & "             -[SMANYUKN03]"
            strSQL = strSQL & "             -[SMANYUKN04]"
            strSQL = strSQL & "             -[SMANYUKN05]"
            strSQL = strSQL & "             -[SMANYUKN06]"
            strSQL = strSQL & "             -[SMANYUKN07]"
            strSQL = strSQL & "             -[SMANYUKN08]"
            strSQL = strSQL & "             -[SMANYUKN09]) "
            strSQL = strSQL & " FROM "
            strSQL = strSQL & "     TOKSMA "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & "     TOKCD = '" & strCODE & "'"
            strSQL = strSQL & " GROUP BY "
            strSQL = strSQL & "     SMADT "
            strSQL = strSQL & " HAVING "
            strSQL = strSQL & "     SMADT > '" & strZ & "' "
            strSQL = strSQL & "     And SMADT <='" & strDate & "'"
            strSQL = strSQL & " ORDER BY "
            strSQL = strSQL & "     SMADT DESC"
            rsS.Open strSQL, cnW, adOpenStatic, adLockReadOnly 'TOKSMA
            If rsS.EOF = False Then
                rsS.MoveFirst
                strZ = rsS.Fields("SMADT")
                Do Until rsS.EOF
                    lngZ = lngZ + rsS.Fields(1)
                    rsS.MoveNext
                Loop
            End If
            rsS.Close
        End If
        rsA.Fields("売掛残") = lngZ
        rsA.Update
        rsA.MoveNext
    Loop
    
Exit_DB:
    
    If Not rsM Is Nothing Then
        If rsM.State = adStateOpen Then rsM.Close
        Set rsM = Nothing
    End If
    If Not rsS Is Nothing Then
        If rsS.State = adStateOpen Then rsS.Close
        Set rsS = Nothing
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


