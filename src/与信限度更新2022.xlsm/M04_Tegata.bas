Attribute VB_Name = "M04_Tegata"
Option Explicit

Sub Tegata_Add()

    Dim cnW     As New ADODB.Connection
    Dim rsT     As New ADODB.Recordset
    Dim rsA     As New ADODB.Recordset
    Dim strSQL  As String
    Dim strDate As String
    Dim dNOW    As Date
    Dim lngT    As Double
    Dim lngM    As Double
    
    '前日生成
    dNOW = DateTime.Date - 1
    strDate = Strings.Format(dNOW, "yyyymmdd")
    
    'データベース設定
    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW
    cnW.Open
    rsA.Open "与信限度データ", cnW, adOpenStatic, adLockPessimistic
    If rsA.EOF Then
        GoTo Exit_DB
    End If
    
    '与信限度データに手形債権追加
    rsA.MoveFirst
    Do Until rsA.EOF
        strCODE = rsA.Fields(0) & ""
        If strCODE <> "" Then
            '手形期日が当月以降の手形データのみ集計
            strSQL = ""
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & "     NYUKN"
            strSQL = strSQL & "     , DKBID "
            strSQL = strSQL & " FROM "
            strSQL = strSQL & "     手形明細 "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & "     TOKCD = '" & strCODE & "'"
            strSQL = strSQL & "     AND TEGDT > '" & strDate & "'"
            rsT.Open strSQL, cnW, adOpenStatic, adLockReadOnly
            If rsT.EOF Then
                rsA.Fields("手形債権") = 0
                rsA.Update
            Else
                rsT.MoveFirst
                lngT = 0: lngM = 0
                Do Until rsT.EOF
                    If IsNull(rsT.Fields(0)) = False Then
                        If rsT.Fields(1) = "03" Then
                            lngT = lngT + rsT.Fields(0)
                        Else
                            lngM = lngM + rsT.Fields(0)
                        End If
                        rsT.MoveNext
                    End If
                Loop
                rsA.Fields("手形債権") = lngT
                rsA.Fields("廻り手形") = lngM
                rsA.Update
            End If
            rsT.Close
        End If
        rsA.MoveNext
    Loop

Exit_DB:


    If Not rsT Is Nothing Then
        If rsT.State = adStateOpen Then rsT.Close
        Set rsT = Nothing
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

