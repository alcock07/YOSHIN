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
    
    'ŋÏ―ĀÅXV
    strSQL = ""
    strSQL = strSQL & "   UPDATE "
    strSQL = strSQL & "       ^MĀxf[^ "
    strSQL = strSQL & "       INNER JOIN ŋæ}X^ ON ^MĀxf[^.GCODE = ŋæ}X^.ŋæš°ÄÞ "
    strSQL = strSQL & "   SET "
    strSQL = strSQL & "       ^MĀxf[^.GNAME = Trim([ŋæ}X^]![O[vž])"
    strSQL = strSQL & "       , ^MĀxf[^.]_ = [ŋæ}X^]![TDBPT]"
    strSQL = strSQL & "       , ^MĀxf[^.Zú = [ŋæ}X^]![TDBDT]"
    strSQL = strSQL & "       , ^MĀxf[^.ÛŊ = [ŋæ}X^]![HOKEN]"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    'åæŠÅXV
    strSQL = ""
    strSQL = strSQL & "   UPDATE "
    strSQL = strSQL & "       ^MĀxf[^ "
    strSQL = strSQL & "       INNER JOIN åæŠ ON ^MĀxf[^.SŌR[h = åæŠ.SŌš°ÄÞ8 "
    strSQL = strSQL & "   SET "
    strSQL = strSQL & "       ^MĀxf[^.xX = Left(åæŠ!xX,2)"
    strSQL = strSQL & "       , ^MĀxf[^.åž = [åæŠ]![åž]"
    strSQL = strSQL & "       , ^MĀxf[^.SŌž = [åæŠ]![SŌŠĖ]"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    'ÓCåXV
    strSQL = ""
    strSQL = strSQL & "   DELETE "
    strSQL = strSQL & "   FROM "
    strSQL = strSQL & "       ÓCå"
    rsX.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    rsX.Open "ÓCå", cnW, adOpenStatic, adLockPessimistic
    
    'Gš°ÄÞ,SŌš°ÄÞēÆĖÂ z
    strSQL = ""
    strSQL = strSQL & "   SELECT "
    strSQL = strSQL & "       GCODE"
    strSQL = strSQL & "       , SŌR[h"
    strSQL = strSQL & "       , Sum(|c)"
    strSQL = strSQL & "       , Sum(č`Â ) "
    strSQL = strSQL & "   FROM "
    strSQL = strSQL & "       ^MĀxf[^ "
    strSQL = strSQL & "   GROUP BY "
    strSQL = strSQL & "       GCODE"
    strSQL = strSQL & "       , SŌR[h "
    strSQL = strSQL & "   ORDER BY "
    strSQL = strSQL & "       GCODE"
    rsA.Open strSQL, cnW, adOpenStatic, adLockReadOnly
    
    If rsA.EOF = False Then rsA.MoveFirst
    dblG = 0
    dblT = 0
    Erase dblF, strF, dblFT
    Do Until rsA.EOF
        If strNCD <> rsA.Fields(0) Then
            'Â Š ęÎÓCåėŽ
            If strNCD <> "" Then
                rsX.AddNew
                rsX.Fields(0) = strNCD
                If dblG = 0 Then
                    'åĖÂ zŠzÆŊķęÍÓCåÆ·é
                    If dblFT(0) > (dblT * 0.8) Then
                        rsX.Fields(1) = "åã"
                        rsX.Fields(3) = "OS"
                    ElseIf dblFT(1) > (dblT * 0.8) Then
                        rsX.Fields(1) = ""
                        rsX.Fields(3) = "TK"
                    ElseIf dblFT(2) > (dblT * 0.8) Then
                        rsX.Fields(1) = "{"
                        rsX.Fields(3) = "HB"
                    ElseIf dblFT(3) > (dblT * 0.8) Then
                        rsX.Fields(1) = "Ö"
                        rsX.Fields(3) = "KA"
                    ElseIf dblFT(4) > (dblT * 0.8) Then
                        rsX.Fields(1) = "C"
                        rsX.Fields(3) = "TA"
                    Else
                        rsX.Fields(1) = "ļÞŲ°Ėß"
                        rsX.Fields(3) = "GR"
                    End If
                    If rsX.Fields(3) = "GR" Then
                        rsX.Fields(2) = ""
                    Else
                        strSQL = ""
                        strSQL = strSQL & " SELECT "
                        strSQL = strSQL & "     SŌR[h"
                        strSQL = strSQL & "     , First(SŌž)"
                        strSQL = strSQL & "     , Sum(|c)"
                        strSQL = strSQL & "     , Sum(č`Â ) "
                        strSQL = strSQL & " FROM "
                        strSQL = strSQL & "     ^MĀxf[^ "
                        strSQL = strSQL & " WHERE "
                        strSQL = strSQL & "     xX = '" & rsX.Fields(1) & "'"
                        strSQL = strSQL & "     AND GCODE = '" & strNCD & "'"
                        strSQL = strSQL & " GROUP BY "
                        strSQL = strSQL & "     SŌR[h "
                        strSQL = strSQL & " ORDER BY "
                        strSQL = strSQL & "     Sum(|c) DESC"
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
                    'åĖÂ zŠzÆŊķęÍÓCåÆ·é
                    If dblF(0) > (dblG * 0.8) Then
                        rsX.Fields(1) = "åã"
                        rsX.Fields(3) = "OS"
                    ElseIf dblF(1) > (dblG * 0.8) Then
                        rsX.Fields(1) = ""
                        rsX.Fields(3) = "TK"
                    ElseIf dblF(2) > (dblG * 0.8) Then
                        rsX.Fields(1) = "{"
                        rsX.Fields(3) = "HB"
                    ElseIf dblF(3) > (dblG * 0.8) Then
                        rsX.Fields(1) = "Ö"
                        rsX.Fields(3) = "KA"
                    ElseIf dblF(4) > (dblG * 0.8) Then
                        rsX.Fields(1) = "C"
                        rsX.Fields(3) = "TA"
                    Else
                        rsX.Fields(1) = "ļÞŲ°Ėß"
                        rsX.Fields(3) = "GR"
                    End If
                    If rsX.Fields(3) = "GR" Then
                        rsX.Fields(2) = ""
                    Else
                        strSQL = ""
                        strSQL = strSQL & " SELECT "
                        strSQL = strSQL & "     SŌR[h"
                        strSQL = strSQL & "     , First(SŌž)"
                        strSQL = strSQL & "     , Sum(|c)"
                        strSQL = strSQL & "     , Sum(č`Â ) "
                        strSQL = strSQL & " FROM "
                        strSQL = strSQL & "     ^MĀxf[^ "
                        strSQL = strSQL & " WHERE "
                        strSQL = strSQL & "     xX ='" & rsX.Fields(1) & "'"
                        strSQL = strSQL & "     AND GCODE = '" & strNCD & "'"
                        strSQL = strSQL & " GROUP BY "
                        strSQL = strSQL & "     SŌR[h "
                        strSQL = strSQL & " ORDER BY "
                        strSQL = strSQL & "     Sum(|c) DESC"
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
        
        dblG = dblG + rsA.Fields(2) 'ļÞŲ°ĖßvÉZ(|)
        dblT = dblT + rsA.Fields(3) 'ļÞŲ°ĖßvÉZ(č`)
        strBmn = Mid(rsA.Fields(1), 5, 2) 'SŌš°ÄÞĖã2ðŧč
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
        'åĖÂ zŠzÆŊķęÍÓCåÆ·é
        If dblFT(0) > (dblT * 0.8) Then
            rsX.Fields(1) = "åã"
            rsX.Fields(3) = "OS"
        ElseIf dblFT(1) > (dblT * 0.8) Then
            rsX.Fields(1) = ""
            rsX.Fields(3) = "TK"
        ElseIf dblFT(2) > (dblT * 0.8) Then
            rsX.Fields(1) = "{"
            rsX.Fields(3) = "HB"
        ElseIf dblFT(3) > (dblT * 0.8) Then
            rsX.Fields(1) = "Ö"
            rsX.Fields(3) = "KA"
        ElseIf dblFT(4) > (dblT * 0.8) Then
            rsX.Fields(1) = "C"
            rsX.Fields(3) = "TA"
        Else
            rsX.Fields(1) = "ļÞŲ°Ėß"
            rsX.Fields(3) = "GR"
        End If
        If rsX.Fields(3) = "GR" Then
            rsX.Fields(2) = ""
        Else
            strSQL = ""
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & "     SŌR[h"
            strSQL = strSQL & "     , First(SŌž)"
            strSQL = strSQL & "     , Sum(|c)"
            strSQL = strSQL & "     , Sum(č`Â ) "
            strSQL = strSQL & " FROM "
            strSQL = strSQL & "     ^MĀxf[^ "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & "     xX = '" & rsX.Fields(1) & "'"
            strSQL = strSQL & "     AND GCODE = '" & strNCD & "'"
            strSQL = strSQL & " GROUP BY "
            strSQL = strSQL & "     SŌR[h "
            strSQL = strSQL & " ORDER BY "
            strSQL = strSQL & "     Sum(|c) DESC"
            
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
        'åĖÂ zŠzÆŊķęÍÓCåÆ·é
        If dblF(0) > (dblG * 0.8) Then
            rsX.Fields(1) = "åã"
            rsX.Fields(3) = "OS"
        ElseIf dblF(1) > (dblG * 0.8) Then
            rsX.Fields(1) = ""
            rsX.Fields(3) = "TK"
        ElseIf dblF(2) > (dblG * 0.8) Then
            rsX.Fields(1) = "{"
            rsX.Fields(3) = "HB"
        ElseIf dblF(3) > (dblG * 0.8) Then
            rsX.Fields(1) = "Ö"
            rsX.Fields(3) = "KA"
        ElseIf dblF(4) > (dblG * 0.8) Then
            rsX.Fields(1) = "C"
            rsX.Fields(3) = "TA"
        Else
            rsX.Fields(1) = "ļÞŲ°Ėß"
            rsX.Fields(3) = "GR"
        End If
        If rsX.Fields(3) = "GR" Then
            rsX.Fields(2) = ""
        Else
            strSQL = ""
            strSQL = strSQL & " SELECT "
            strSQL = strSQL & "     SŌR[h"
            strSQL = strSQL & "     , First(SŌž)"
            strSQL = strSQL & "     , Sum(|c)"
            strSQL = strSQL & "     , Sum(č`Â ) "
            strSQL = strSQL & " FROM "
            strSQL = strSQL & "     ^MĀxf[^ "
            strSQL = strSQL & " WHERE "
            strSQL = strSQL & "     xX = '" & rsX.Fields(1) & "'"
            strSQL = strSQL & "     AND GCODE = '" & strNCD & "'"
            strSQL = strSQL & " GROUP BY "
            strSQL = strSQL & "     SŌR[h "
            strSQL = strSQL & " ORDER BY "
            strSQL = strSQL & "     Sum(|c) DESC"
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
    
    'ÓCåXV
    strSQL = ""
    strSQL = strSQL & "   UPDATE "
    strSQL = strSQL & "       ^MĀxf[^ "
    strSQL = strSQL & "       INNER JOIN ÓCå ON ^MĀxf[^.GCODE = ÓCå.CODE "
    strSQL = strSQL & "   SET "
    strSQL = strSQL & "       ^MĀxf[^.ÓCå = [ÓCå]![SBMN]"
    strSQL = strSQL & "       , ^MĀxf[^.SŌž = [ÓCå]![STAN]"
    strSQL = strSQL & "       , ^MĀxf[^.GæŠ = [ÓCå]![GKBN]"
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
