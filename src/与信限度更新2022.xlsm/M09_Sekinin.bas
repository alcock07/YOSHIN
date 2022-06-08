Attribute VB_Name = "M09_Sekinin"
Option Explicit

Sub Sekinin_Data()
    
    Dim cnW      As New ADODB.Connection
    Dim rsA      As New ADODB.Recordset
    Dim rsB      As New ADODB.Recordset
    Dim rsX      As New ADODB.Recordset
    Dim strSQL   As String
    Dim strNCD   As String
    Dim strBmn   As String
    Dim dblG     As Double
    Dim dblT     As Double
    Dim dblF(4)  As Double
    Dim dblFT(4) As Double
    Dim strF(4)  As String

    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW
    cnW.Open
    
    '請求ﾏｽﾀで更新
    strSQL = ""
    strSQL = strSQL & "   UPDATE "
    strSQL = strSQL & "       与信限度データ "
    strSQL = strSQL & "       INNER JOIN 請求先マスタ ON 与信限度データ.GCODE = 請求先マスタ.請求先ｺｰﾄﾞ "
    strSQL = strSQL & "   SET "
    strSQL = strSQL & "       与信限度データ.GNAME = Trim([請求先マスタ]![グループ名])"
    strSQL = strSQL & "       , 与信限度データ.評点 = [請求先マスタ]![TDBPT]"
    strSQL = strSQL & "       , 与信限度データ.決算日 = [請求先マスタ]![TDBDT]"
    strSQL = strSQL & "       , 与信限度データ.保険 = [請求先マスタ]![HOKEN]"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    '部門区分で更新
    strSQL = ""
    strSQL = strSQL & "   UPDATE "
    strSQL = strSQL & "       与信限度データ "
    strSQL = strSQL & "       INNER JOIN 部門区分 ON 与信限度データ.担当者コード = 部門区分.担当者ｺｰﾄﾞ8 "
    strSQL = strSQL & "   SET "
    strSQL = strSQL & "       与信限度データ.支店 = Left(部門区分!支店,2)"
    strSQL = strSQL & "       , 与信限度データ.部門名 = [部門区分]![部門名]"
    strSQL = strSQL & "       , 与信限度データ.担当者名 = [部門区分]![担当者略称]"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    '責任部門更新
    strSQL = ""
    strSQL = strSQL & "   DELETE "
    strSQL = strSQL & "   FROM "
    strSQL = strSQL & "       責任部門"
    rsX.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    rsX.Open "責任部門", cnW, adOpenStatic, adLockPessimistic
    
    'Gｺｰﾄﾞ,担当者ｺｰﾄﾞごとの債権額
    strSQL = ""
    strSQL = strSQL & "   SELECT "
    strSQL = strSQL & "       GCODE"
    strSQL = strSQL & "       , 担当者コード"
    strSQL = strSQL & "       , Sum(売掛残)"
    strSQL = strSQL & "       , Sum(手形債権) "
    strSQL = strSQL & "   FROM "
    strSQL = strSQL & "       与信限度データ "
    strSQL = strSQL & "   GROUP BY "
    strSQL = strSQL & "       GCODE"
    strSQL = strSQL & "       , 担当者コード "
    strSQL = strSQL & "   ORDER BY "
    strSQL = strSQL & "       GCODE"
    rsA.Open strSQL, cnW, adOpenStatic, adLockReadOnly
    
    If rsA.EOF = False Then rsA.MoveFirst
    dblG = 0
    dblT = 0
    Erase dblF, strF, dblFT
    Do Until rsA.EOF
        If strNCD <> rsA.Fields(0) Then
            '債権があれば責任部門作成
            If strNCD <> "" Then
                rsX.AddNew
                rsX.Fields(0) = strNCD
                If dblG = 0 Then
                    '部門の債権額が総額と同じ場合は責任部門とする
                    If dblFT(0) > (dblT * 0.8) Then
                        rsX.Fields(1) = "大阪"
                        rsX.Fields(3) = "OS"
                    ElseIf dblFT(1) > (dblT * 0.8) Then
                        rsX.Fields(1) = "東京"
                        rsX.Fields(3) = "TK"
                    ElseIf dblFT(2) > (dblT * 0.8) Then
                        rsX.Fields(1) = "本部"
                        rsX.Fields(3) = "HB"
                    ElseIf dblFT(3) > (dblT * 0.8) Then
                        rsX.Fields(1) = "関東"
                        rsX.Fields(3) = "KA"
                    ElseIf dblFT(4) > (dblT * 0.8) Then
                        rsX.Fields(1) = "東海"
                        rsX.Fields(3) = "TA"
                    Else
                        rsX.Fields(1) = "ｸﾞﾙｰﾌﾟ"
                        rsX.Fields(3) = "GR"
                    End If
                    If rsX.Fields(3) = "GR" Then
                        rsX.Fields(2) = ""
                    Else
                        strSQL = ""
                        strSQL = strSQL & " SELECT "
                        strSQL = strSQL & "     担当者コード"
                        strSQL = strSQL & "     , First(担当者名)"
                        strSQL = strSQL & "     , Sum(売掛残)"
                        strSQL = strSQL & "     , Sum(手形債権) "
                        strSQL = strSQL & " FROM "
                        strSQL = strSQL & "     与信限度データ "
                        strSQL = strSQL & " WHERE "
                        strSQL = strSQL & "     支店 = '" & rsX.Fields(1) & "'"
                        strSQL = strSQL & "     AND GCODE = '" & strNCD & "'"
                        strSQL = strSQL & " GROUP BY "
                        strSQL = strSQL & "     担当者コード "
                        strSQL = strSQL & " ORDER BY "
                        strSQL = strSQL & "     Sum(売掛残) DESC"
                        rsB.Open strSQL, cnW, adOpenStatic, adLockReadOnly
                        If rsB.EOF = False Then
                            rsB.MoveFirst
                            If rsB.Fields(3) > (dblT * 0.8) Then
                                rsX.Fields(2) = rsB.Fields(1)
                            Else
                                rsX.Fields(2) = ""
                            End If
                        Else
                            rsX.Fields(2) = ""
                        End If
                        rsB.Close
                    End If
                Else
                    '部門の債権額が総額と同じ場合は責任部門とする
                    If dblF(0) > (dblG * 0.8) Then
                        rsX.Fields(1) = "大阪"
                        rsX.Fields(3) = "OS"
                    ElseIf dblF(1) > (dblG * 0.8) Then
                        rsX.Fields(1) = "東京"
                        rsX.Fields(3) = "TK"
                    ElseIf dblF(2) > (dblG * 0.8) Then
                        rsX.Fields(1) = "本部"
                        rsX.Fields(3) = "HB"
                    ElseIf dblF(3) > (dblG * 0.8) Then
                        rsX.Fields(1) = "関東"
                        rsX.Fields(3) = "KA"
                    ElseIf dblF(4) > (dblG * 0.8) Then
                        rsX.Fields(1) = "東海"
                        rsX.Fields(3) = "TA"
                    Else
                        rsX.Fields(1) = "ｸﾞﾙｰﾌﾟ"
                        rsX.Fields(3) = "GR"
                    End If
                    If rsX.Fields(3) = "GR" Then
                        rsX.Fields(2) = ""
                    Else
                        strSQL = ""
                        strSQL = strSQL & " SELECT "
                        strSQL = strSQL & "     担当者コード"
                        strSQL = strSQL & "     , First(担当者名)"
                        strSQL = strSQL & "     , Sum(売掛残)"
                        strSQL = strSQL & "     , Sum(手形債権) "
                        strSQL = strSQL & " FROM "
                        strSQL = strSQL & "     与信限度データ "
                        strSQL = strSQL & " WHERE "
                        strSQL = strSQL & "     支店 ='" & rsX.Fields(1) & "'"
                        strSQL = strSQL & "     AND GCODE = '" & strNCD & "'"
                        strSQL = strSQL & " GROUP BY "
                        strSQL = strSQL & "     担当者コード "
                        strSQL = strSQL & " ORDER BY "
                        strSQL = strSQL & "     Sum(売掛残) DESC"
                        rsB.Open strSQL, cnW, adOpenStatic, adLockReadOnly
                        If rsB.EOF = False Then
                            rsB.MoveFirst
                            If rsB.Fields(2) > (dblG * 0.8) Then
                                rsX.Fields(2) = rsB.Fields(1)
                            Else
                                rsX.Fields(2) = ""
                            End If
                        Else
                            rsX.Fields(2) = ""
                        End If
                        rsB.Close
                    End If
                End If
                rsX.Update
            End If
            strNCD = rsA.Fields(0)
            dblG = 0
            dblT = 0
            Erase dblF, strF, dblFT
        End If
        
        dblG = dblG + rsA.Fields(2) 'ｸﾞﾙｰﾌﾟ計に合算(売掛)
        dblT = dblT + rsA.Fields(3) 'ｸﾞﾙｰﾌﾟ計に合算(手形)
        strBmn = Mid(rsA.Fields(1), 5, 2) '担当者ｺｰﾄﾞの上2桁を判定
        If strBmn = "01" Then
            dblF(0) = dblF(0) + rsA.Fields(2)
            dblFT(0) = dblFT(0) + rsA.Fields(3)
            If strF(0) = "" Then strF(0) = rsA.Fields(1)
        ElseIf strBmn = "02" Then
            dblF(1) = dblF(1) + rsA.Fields(2)
            dblFT(1) = dblFT(1) + rsA.Fields(3)
            If strF(1) = "" Then strF(1) = rsA.Fields(1)
        ElseIf strBmn = "07" Then
            dblF(3) = dblF(3) + rsA.Fields(2)
            dblFT(3) = dblFT(3) + rsA.Fields(3)
            If strF(3) = "" Then strF(3) = rsA.Fields(1)
        ElseIf strBmn = "08" Then
            dblF(4) = dblF(4) + rsA.Fields(2)
            dblFT(4) = dblFT(4) + rsA.Fields(3)
            If strF(4) = "" Then strF(4) = rsA.Fields(1)
        Else
            dblF(2) = dblF(2) + rsA.Fields(2)
            dblFT(2) = dblFT(2) + rsA.Fields(3)
            If strF(2) = "" Then strF(2) = rsA.Fields(1)
        End If
        rsA.MoveNext
    Loop

    rsX.AddNew
    rsX.Fields(0) = strNCD
    If dblG = 0 Then
        '部門の債権額が総額と同じ場合は責任部門とする
        If dblFT(0) > (dblT * 0.8) Then
            rsX.Fields(1) = "大阪"
            rsX.Fields(3) = "OS"
        ElseIf dblFT(1) > (dblT * 0.8) Then
            rsX.Fields(1) = "東京"
            rsX.Fields(3) = "TK"
        ElseIf dblFT(2) > (dblT * 0.8) Then
            rsX.Fields(1) = "本部"
            rsX.Fields(3) = "HB"
        ElseIf dblFT(3) > (dblT * 0.8) Then
            rsX.Fields(1) = "関東"
            rsX.Fields(3) = "KA"
        ElseIf dblFT(4) > (dblT * 0.8) Then
            rsX.Fields(1) = "東海"
            rsX.Fields(3) = "TA"
        Else
            rsX.Fields(1) = "ｸﾞﾙｰﾌﾟ"
            rsX.Fields(3) = "GR"
        End If
        If rsX.Fields(3) = "GR" Then
            rsX.Fields(2) = ""
        Else
            strSQL = ""
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & "     担当者コード"
            strSQL = strSQL & "     , First(担当者名)"
            strSQL = strSQL & "     , Sum(売掛残)"
            strSQL = strSQL & "     , Sum(手形債権) "
            strSQL = strSQL & " FROM "
            strSQL = strSQL & "     与信限度データ "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & "     支店 = '" & rsX.Fields(1) & "'"
            strSQL = strSQL & "     AND GCODE = '" & strNCD & "'"
            strSQL = strSQL & " GROUP BY "
            strSQL = strSQL & "     担当者コード "
            strSQL = strSQL & " ORDER BY "
            strSQL = strSQL & "     Sum(売掛残) DESC"
            
            rsB.Open strSQL, cnW, adOpenStatic, adLockReadOnly
            If rsB.EOF = False Then
                rsB.MoveFirst
                If rsB.Fields(3) > (dblT * 0.8) Then
                    rsX.Fields(2) = rsB.Fields(1)
                Else
                    rsX.Fields(2) = ""
                End If
             Else
                rsX.Fields(2) = ""
            End If
            rsB.Close
        End If
    Else
        '部門の債権額が総額と同じ場合は責任部門とする
        If dblF(0) > (dblG * 0.8) Then
            rsX.Fields(1) = "大阪"
            rsX.Fields(3) = "OS"
        ElseIf dblF(1) > (dblG * 0.8) Then
            rsX.Fields(1) = "東京"
            rsX.Fields(3) = "TK"
        ElseIf dblF(2) > (dblG * 0.8) Then
            rsX.Fields(1) = "本部"
            rsX.Fields(3) = "HB"
        ElseIf dblF(3) > (dblG * 0.8) Then
            rsX.Fields(1) = "関東"
            rsX.Fields(3) = "KA"
        ElseIf dblF(4) > (dblG * 0.8) Then
            rsX.Fields(1) = "東海"
            rsX.Fields(3) = "TA"
        Else
            rsX.Fields(1) = "ｸﾞﾙｰﾌﾟ"
            rsX.Fields(3) = "GR"
        End If
        If rsX.Fields(3) = "GR" Then
            rsX.Fields(2) = ""
        Else
            strSQL = ""
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & "     担当者コード"
            strSQL = strSQL & "     , First(担当者名)"
            strSQL = strSQL & "     , Sum(売掛残)"
            strSQL = strSQL & "     , Sum(手形債権) "
            strSQL = strSQL & " FROM "
            strSQL = strSQL & "     与信限度データ "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & "     支店 = '" & rsX.Fields(1) & "'"
            strSQL = strSQL & "     AND GCODE = '" & strNCD & "'"
            strSQL = strSQL & " GROUP BY "
            strSQL = strSQL & "     担当者コード "
            strSQL = strSQL & " ORDER BY "
            strSQL = strSQL & "     Sum(売掛残) DESC"
            rsB.Open strSQL, cnW, adOpenStatic, adLockReadOnly
            If rsB.EOF = False Then
                rsB.MoveFirst
                If rsB.Fields(2) > (dblG * 0.8) Then
                    rsX.Fields(2) = rsB.Fields(1)
                Else
                    rsX.Fields(2) = ""
                End If
            Else
                rsX.Fields(2) = ""
            End If
            rsB.Close
        End If
    End If
    rsX.Update
    
    rsA.Close
    
    '責任部門更新
    strSQL = ""
    strSQL = strSQL & "   UPDATE "
    strSQL = strSQL & "       与信限度データ "
    strSQL = strSQL & "       INNER JOIN 責任部門 ON 与信限度データ.GCODE = 責任部門.CODE "
    strSQL = strSQL & "   SET "
    strSQL = strSQL & "       与信限度データ.責任部門 = [責任部門]![SBMN]"
    strSQL = strSQL & "       , 与信限度データ.担当者名 = [責任部門]![STAN]"
    strSQL = strSQL & "       , 与信限度データ.G区分 = [責任部門]![GKBN]"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
Exit_DB:
    
    If Not rsA Is Nothing Then
        If rsA.State = adStateOpen Then rsA.Close
        Set rsA = Nothing
    End If
    If Not rsB Is Nothing Then
        If rsB.State = adStateOpen Then rsB.Close
        Set rsB = Nothing
    End If
    If Not rsX Is Nothing Then
        If rsX.State = adStateOpen Then rsX.Close
        Set rsX = Nothing
    End If
    If Not cnW Is Nothing Then
        If cnW.State = adStateOpen Then cnW.Close
        Set cnW = Nothing
    End If
    
End Sub
