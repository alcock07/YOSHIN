Attribute VB_Name = "M10_Move"
Option Explicit

Sub Move_Data()

'===================
'�^�M���x�f�[�^�ړ�
'===================

    Dim cnW    As New ADODB.Connection
    Dim rsA    As New ADODB.Recordset
    Dim strSQL As String
    
    cnW.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" & dbW
    cnW.Open
    
    '�f�[�^�Q��DB�N���A
    strSQL = ""
    strSQL = strSQL & "DELETE"
    strSQL = strSQL & "       FROM �^�M���xLink"
    rsA.Open strSQL, cnW, adOpenStatic, adLockPessimistic
    
    '�f�[�^�ڍs
    strSQL = ""
    strSQL = strSQL & "INSERT INTO �^�M���xLink"
    strSQL = strSQL & "       SELECT �^�M���x�f�[�^.*"
    strSQL = strSQL & "       FROM �^�M���x�f�[�^"
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

