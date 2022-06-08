Attribute VB_Name = "M02_Data"
Option Explicit

Public Const dbW As String = "\\192.168.128.4\hb\sys\���|\�����\�^�M���xW.accdb"
Public strCODE   As String
Private strDate  As String
Private DateA    As Date
Private lngYY    As Long
Private lngMM    As Long
Private lngZZ    As Long

'===== �T�� =====
'���Ӑ�̍��f�[�^���v�Z����A�����X�V����
'�f�[�^�͉c�Ǝ��сA���c�A����`�A�󒍎c
'�g�p�e�[�u���FOracle�iUDNTRA,TOKMTA,TOKSMA�j
'              SQLServer�iJUZTBZ_Hybrid�j
'              Access�i���ѓ���,���ёO���j

Sub Proc_Main()
    
    Sheets("Sheet1").Range("A1").Select
    DoEvents
    
    Call Reset_Table   ' �e�[�u��������(UDNTRA->��`����)
    Call Set_Data      ' ���Ӑ�}�X�^����^�M���x�f�[�^�쐬���c�Ǝ��т�����уf�[�^�ǉ�
    Call Tegata_Add    ' ��`���ǉ�
    Call ZAN_Change    ' ���|�c���X�V
    Call Set_JOKEN     ' ����������X�V
    Call J_ZAN         ' �󒍎c�ǉ�
    Call Del_Data      ' �s�v�f�[�^�폜
    Call Sekinin_Data  ' �ӔC����X�V
    Call Move_Data     ' �f�[�^�Q��DB�X�V
        
End Sub

Sub Reset_Table()

    Dim cnW  As New ADODB.Connection
    Dim rsA  As New ADODB.Recordset
    Dim strSQL As String
    Dim intC   As Long
    
    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW
    cnW.Open
    
    '�^�M���x�f�[�^�N���A
    strSQL = ""
    strSQL = strSQL & "DELETE "
    strSQL = strSQL & "       FROM �^�M���x�f�[�^"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    '��`���׃f�[�^�N���A
    strSQL = ""
    strSQL = strSQL & "DELETE "
    strSQL = strSQL & "       FROM ��`����"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    'UDNTRA->��`����
    Dim SQLA1 As String
    strSQL = ""
    strSQL = strSQL & "INSERT INTO "
    strSQL = strSQL & "            ��`���� ( "
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
    strSQL = strSQL & "              WHERE (UDNTRA.DKBID = '03' Or UDNTRA.DKBID = '08')" '����敪�R�[�h(03:��`,08���`)
    strSQL = strSQL & "              And UDNTRA.DATKB = '1' "                            '�`�[�폜�敪(1:�g�p��,9:�폜)
    strSQL = strSQL & "              And UDNTRA.DENKB = '8' "                            '�`�[�敪(8:����)
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
