Attribute VB_Name = "M03_JISSEKI"
Option Explicit

Sub Set_Data()
    
    Dim cnW    As New ADODB.Connection
    Dim rsS    As New ADODB.Recordset
    Dim rsM    As New ADODB.Recordset
    Dim rsA    As New ADODB.Recordset
    Dim strSQL As String
    Dim lng12  As Long     '月判定
    Dim lngYD  As Long     '年度算出用
    Dim strTK  As String   '当期
    Dim strZK  As String   '前期
    Dim lngM   As Long
    Dim lng3H  As Long
    Dim lngTH  As Long
    Dim lngZH  As Long
    Dim lngW   As Long
        
    '期の算出
    lng12 = CLng(Format(Now(), "mm"))
    lngYD = CLng(Format(Now(), "yyyy"))
    If lng12 > 10 Then
        lngYD = lngYD + 1
    End If
    strTK = CStr(lngYD)
    strZK = CStr(lngYD - 1)
        
    '得意先マスタから与信限度データ作成（当月&前月実績追加）
    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW
    cnW.Open
    rsA.Open "与信限度データ", cnW, adOpenStatic, adLockPessimistic '与信限度データテーブル
        
    '得意先マスタ(TOKMTA)
    strSQL = ""
    strSQL = strSQL & "   SELECT "
    strSQL = strSQL & "       TOKCD"
    strSQL = strSQL & "       , TOKNMA"
    strSQL = strSQL & "       , TOKNMB"
    strSQL = strSQL & "       , TOKRN"
    strSQL = strSQL & "       , TOKBMNCD"
    strSQL = strSQL & "       , TANCD"
    strSQL = strSQL & "       , SMAZANDT"
    strSQL = strSQL & "       , SMAZANKN"
    strSQL = strSQL & "       , LMTKN "
    strSQL = strSQL & "   FROM "
    strSQL = strSQL & "       TOKMTA "
    strSQL = strSQL & "   WHERE "
    strSQL = strSQL & "       TOKRN <> 'X' "
    strSQL = strSQL & "       AND DATKB = '1'"
    strSQL = strSQL & "   ORDER BY "
    strSQL = strSQL & "       TOKBMNCD"
    strSQL = strSQL & "       , TANCD"
    
    rsM.Open strSQL, cnW, adOpenStatic, adLockReadOnly
    rsM.MoveFirst
    lngW = 0
    cnW.BeginTrans
    Do Until rsM.EOF
        strCODE = rsM.Fields(0) & ""
        lngTH = 0: lngZH = 0: lng3H = 0
        If strCODE <> "" Then
            rsA.AddNew
            rsA.Fields("得意先コード") = strCODE
            rsA.Fields("得意先名") = Trim(rsM.Fields("TOKNMA")) & " " & Trim(rsM.Fields("TOKNMB"))  '得意先名
            rsA.Fields("与信限度額") = CDbl(rsM.Fields("LMTKN"))                                    '与信限度額
            rsA.Fields("部門コード") = rsM.Fields("TOKBMNCD")                                       '部門コード
            rsA.Fields("担当者コード") = rsM.Fields("TANCD")                                        '担当者コード
            '営業実績にアクセスして実績データ追加 ==========
            strSQL = ""
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & "     MONTH"
            strSQL = strSQL & "     , Sum(UDNKN) as UDNKN "
            strSQL = strSQL & " FROM "
            strSQL = strSQL & "     実績当期 "
            strSQL = strSQL & " GROUP BY "
            strSQL = strSQL & "     TOKCD"
            strSQL = strSQL & "     , YEARD"
            strSQL = strSQL & "     , MONTH "
            strSQL = strSQL & " HAVING "
            strSQL = strSQL & "     TOKCD = '" & strCODE & "'"
            strSQL = strSQL & "     AND YEARD = '" & strTK & "'"
            rsS.Open strSQL, cnW, adOpenStatic, adLockReadOnly '得意先毎売上抽出
            
            If rsS.EOF = False Then
                Do Until rsS.EOF
                    Select Case rsS.Fields("MONTH")
                        Case "10"
                            rsA.Fields("実績10") = rsS.Fields(1)
                        Case "11"
                            rsA.Fields("実績11") = rsS.Fields(1)
                        Case "12"
                            rsA.Fields("実績12") = rsS.Fields(1)
                        Case "01"
                            rsA.Fields("実績01") = rsS.Fields(1)
                        Case "02"
                            rsA.Fields("実績02") = rsS.Fields(1)
                        Case "03"
                            rsA.Fields("実績03") = rsS.Fields(1)
                        Case "04"
                            rsA.Fields("実績04") = rsS.Fields(1)
                        Case "05"
                            rsA.Fields("実績05") = rsS.Fields(1)
                        Case "06"
                            rsA.Fields("実績06") = rsS.Fields(1)
                        Case "07"
                            rsA.Fields("実績07") = rsS.Fields(1)
                        Case "08"
                            rsA.Fields("実績08") = rsS.Fields(1)
                        Case "09"
                            rsA.Fields("実績09") = rsS.Fields(1)
                    End Select
                    rsS.MoveNext
                Loop
            End If
            rsS.Close
            
            '当期平均計算
            If lng12 > 0 And lng12 < 10 Then  '1～9月
                lngTH = rsA.Fields(3) + rsA.Fields(4) + rsA.Fields(5)
                For lngM = 1 To lng12
                    lngTH = lngTH + rsA.Fields(lngM + 5)
                Next lngM
                rsA.Fields("当期平均") = lngTH / lng12 + 3
                lng3H = rsA.Fields(lng12 + 2) + rsA.Fields(lng12 + 3) + rsA.Fields(lng12 + 4) '3ヵ月平均
            ElseIf lng12 = 10 Then
                rsA.Fields("当期平均") = 0
            ElseIf lng12 = 11 Then
                rsA.Fields("当期平均") = rsA.Fields(3)
                lng3H = rsA.Fields(3) '3ヵ月平均
            ElseIf lng12 = 12 Then
                lngTH = rsA.Fields(3) + rsA.Fields(4)
                rsA.Fields("当期平均") = lngTH / 2
                lng3H = rsA.Fields(3) + rsA.Fields(4) '3ヵ月平均
            End If
            
            '前期データ取得
            strSQL = ""
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & "     MONTH"
            strSQL = strSQL & "     , Sum(UDNKN) "
            strSQL = strSQL & " FROM "
            strSQL = strSQL & "     実績前期 "
            strSQL = strSQL & " GROUP BY "
            strSQL = strSQL & "     TOKCD"
            strSQL = strSQL & "     , YEARD"
            strSQL = strSQL & "     , MONTH "
            strSQL = strSQL & " HAVING "
            strSQL = strSQL & "     TOKCD = '" & strCODE & " '"
            strSQL = strSQL & "     AND YEARD = '" & strZK & "'"
            rsS.Open strSQL, cnW, adOpenStatic, adLockReadOnly
            
            If rsS.EOF = False Then
                Do Until rsS.EOF
                    lngZH = lngZH + rsS.Fields(1)
                    If lng12 = 10 And (rsS.Fields("MONTH") = "07" Or rsS.Fields("MONTH") = "08" Or rsS.Fields("MONTH") = "09") Then
                        lng3H = lng3H + rsS.Fields(1)
                    End If
                    If lng12 = 11 And (rsS.Fields("MONTH") = "08" Or rsS.Fields("MONTH") = "09") Then
                        lng3H = lng3H + rsS.Fields(1)
                    End If
                    If lng12 = 12 And (rsS.Fields("MONTH") = "09") Then
                        lng3H = lng3H + rsS.Fields(1)
                    End If
                    rsS.MoveNext
                Loop
            End If
            rsS.Close
            rsA.Fields("前期平均") = lngZH / 12
            rsA.Fields("3ヶ月平均") = lng3H / 3
            rsA.Update
            lngW = lngW + 1
            If lngW > 5000 Then
                cnW.CommitTrans
                cnW.BeginTrans
                lngW = 0
            End If
        End If
        rsM.MoveNext
    Loop
    cnW.CommitTrans
    
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

