Attribute VB_Name = "M02_Data"
Option Explicit

Public Const dbW As String = "\\192.168.128.4\hb\sys\売掛\売上債権\与信限度W.accdb"
Public strCODE   As String
Private strDate  As String
Private DateA    As Date
Private lngYY    As Long
Private lngMM    As Long
Private lngZZ    As Long

'===== 概略 =====
'得意先の債権データを計算する、毎日更新する
'データは営業実績、債権残、受取手形、受注残
'使用テーブル：Oracle（UDNTRA,TOKMTA,TOKSMA）
'              SQLServer（JUZTBZ_Hybrid）
'              Access（実績当期,実績前期）

Sub Proc_Main()
    
    Sheets("Sheet1").Range("A1").Select
    DoEvents
    
    Call Reset_Table   ' テーブル初期化(UDNTRA->手形明細)
    Call Set_Data      ' 得意先マスタから与信限度データ作成＆営業実績から実績データ追加
    Call Tegata_Add    ' 手形債権追加
    Call ZAN_Change    ' 売掛残を更新
    Call Set_JOKEN     ' 取引条件を更新
    Call J_ZAN         ' 受注残追加
    Call Del_Data      ' 不要データ削除
    Call Sekinin_Data  ' 責任部門更新
    Call Move_Data     ' データ参照DB更新
        
End Sub

Sub Reset_Table()

    Dim cnW  As New ADODB.Connection
    Dim rsA  As New ADODB.Recordset
    Dim strSQL As String
    Dim intC   As Long
    
    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW
    cnW.Open
    
    '与信限度データクリア
    strSQL = ""
    strSQL = strSQL & "DELETE "
    strSQL = strSQL & "       FROM 与信限度データ"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    '手形明細データクリア
    strSQL = ""
    strSQL = strSQL & "DELETE "
    strSQL = strSQL & "       FROM 手形明細"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    'UDNTRA->手形明細
    Dim SQLA1 As String
    strSQL = ""
    strSQL = strSQL & "INSERT INTO "
    strSQL = strSQL & "            手形明細 ( "
    strSQL = strSQL & "                       TOKCD,"
    strSQL = strSQL & "                       TEGDT,"
    strSQL = strSQL & "                       NYUKN,"
    strSQL = strSQL & "                       LINCMA,"
    strSQL = strSQL & "                       DKBID,"
    strSQL = strSQL & "                       DENNO"
    strSQL = strSQL & "                      ) "
    strSQL = strSQL & "         SELECT UDNTRA.TOKCD,"
    strSQL = strSQL & "                UDNTRA.TEGDT,"
    strSQL = strSQL & "                UDNTRA.NYUKN,"
    strSQL = strSQL & "                UDNTRA.LINCMA,"
    strSQL = strSQL & "                UDNTRA.DKBID,"
    strSQL = strSQL & "                UDNTRA.UDNNO"
    strSQL = strSQL & "         FROM UDNTRA"
    strSQL = strSQL & "              WHERE (UDNTRA.DKBID = '03' Or UDNTRA.DKBID = '08')" '取引区分コード(03:手形,08廻手形)
    strSQL = strSQL & "              And UDNTRA.DATKB = '1' "                            '伝票削除区分(1:使用中,9:削除)
    strSQL = strSQL & "              And UDNTRA.DENKB = '8' "                            '伝票区分(8:入金)
    strSQL = strSQL & "         ORDER BY UDNTRA.TOKCD"
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
