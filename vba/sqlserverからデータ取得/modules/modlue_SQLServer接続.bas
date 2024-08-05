Attribute VB_Name = "modlue_SQLServer�ڑ�"
Option Explicit

Function SQLServer�ڑ��J�n( _
    ByVal server As String _
    , ByVal method As String _
    , ByVal userName As String _
    , ByVal password As String _
    , ByVal dbName As String _
) As Object
    
On Error GoTo TAG_ERROR
    
    '-------------------------------
    ' �ϐ���`
    '-------------------------------
    Dim sqlserverObject As ADODB.Connection
    Set sqlserverObject = New ADODB.Connection
    Dim connectionStr As String
        
    
    '-------------------------------
    ' �ڑ��̎��s
    '-------------------------------
    '--- �ڑ�������̍\�z ---
    connectionStr = _
        "Provider=SQLOLEDB;" & _
        "Data Source=" & server & ";" & _
        "Initial Catalog=" & dbName & ";"
    
    If method = "windows" Then
        connectionStr = connectionStr & _
            "Trusted_Connection=Yes;"
    ElseIf method = "sqlserver" Then
        connectionStr = connectionStr & _
            "user id=" & userName & ";" & _
            "password=" & password & ";"
    End If
    
    connectionStr = connectionStr & "Connect Timeout=3;"
        
    '--- �ڑ��J�n ---
    sqlserverObject.Open connectionStr
    MsgBox _
        "DB�̐ڑ��ɐ������܂����B"
        
    
    '-------------------------------
    ' return
    '-------------------------------
    Set SQLServer�ڑ��J�n = sqlserverObject
    
    Exit Function
    
TAG_ERROR:
    
    MsgBox _
        "SQLServer�̐ڑ��Ɏ��s���܂����B" _
        , vbCritical _
        , "�G���[�FSQLServer�ڑ��J�n"
    
    Set SQLServer�ڑ��J�n = New ADODB.Connection
    
End Function


Function SQLServer�X�N���v�g���s(ByVal execSQL As String)
    
    '-------------------------------
    ' �ϐ���`
    '-------------------------------
    Set PUB_sqlserverRecordSet = CreateObject("ADODB.RecordSet")
    
    '-------------------------------
    ' ���s�O���f
    '-------------------------------
    If Len(execSQL) = 0 Then
        MsgBox "SQL���s�����񒷂�0�ł��B", vbCritical, "���s�s�FSQLServer�X�N���v�g���s"
        Exit Function
    End If
    
    
On Error GoTo TAG_ERROR
    '-------------------------------
    ' ���s
    '-------------------------------
    PUB_sqlserverRecordSet.Open _
        execSQL _
        , PUB_sqlserverConnection _
        , adOpenStatic, adLockReadOnly, adCmdText
    
    Exit Function
    
TAG_ERROR:
    MsgBox _
        "�G���[�����B���e���m�F���ĉ������B" & vbCrLf & Err.Description _
        , vbCritical _
        , "�G���[�FSQLServer�X�N���v�g���s"
    
End Function


Function Select���ʕ`��(ByVal addressDrawStart As String) As String
    '-------------------------------
    ' Excel���ʕ`��
    '-------------------------------
    Dim roopHeader As Integer
    Dim headerString As String

On Error GoTo TAG_ERROR
    
    '���R�[�h�Z�b�g�֐擪��
    PUB_sqlserverRecordSet.MoveFirst
   
    '-------------------------------
    ' Excel���ʕ`��
    '-------------------------------
    ' �擾SQL���
    Range(addressDrawStart).Value = "�擾SQL:"
    Range(addressDrawStart).Offset(0, 1).Value = PUB_sqlserverRecordSet.Source
    
    '�w�b�_�����o��
    For roopHeader = 1 To PUB_sqlserverRecordSet.Fields.Count
        ' �����o��������̊m��
        Select Case PUB_sqlserverRecordSet.Fields(roopHeader - 1).Name
            Case ""
                headerString = "Column" & roopHeader
            Case Else
                headerString = PUB_sqlserverRecordSet.Fields(roopHeader - 1).Name
        End Select
        
        ' �w�b�_������̋L��
        Range(addressDrawStart).Offset(1, roopHeader - 1) = headerString
    Next
    
    '�f�[�^�����o��
    Range(addressDrawStart).Offset(2, 0).CopyFromRecordset PUB_sqlserverRecordSet
    
    
    '-------------------------------
    ' �߂�l�̎Z�o�F�`��ŏI�A�h���X�Ԓn
    '-------------------------------
    Select���ʕ`�� = Range(addressDrawStart).Offset(PUB_sqlserverRecordSet.RecordCount + 1, 0).Address
    

    Exit Function
    
TAG_ERROR:
    MsgBox _
        "�G���[�����B���e���m�F���ĉ������B" & vbCrLf & Err.Description _
        , vbCritical _
        , "�G���[�FSQLServer�X�N���v�g���s"
        
    Select���ʕ`�� = addressDrawStart
        
End Function


Function SQLServer�ڑ��I��()
    If PUB_sqlserverConnection = "" Then
        Exit Function
    End If

    If MsgBox("�����R�l�N�V������j�����܂����H", vbOKCancel, "�㏑���m�F") = vbCancel Then
        Exit Function
    End If
    
    PUB_sqlserverConnection.Close
    Set PUB_sqlserverConnection = Nothing
    
    PUB_sqlserverRecordSet.Close
    Set PUB_sqlserverRecordSet = Nothing
    
    MsgBox "�j�����܂����B"
    
End Function



