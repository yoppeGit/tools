Attribute VB_Name = "module_�{�^������"
Option Explicit

Sub �ڑ����o�^()
        
    '------------------------------
    ' �����R�l�N�V�����L���m�F
    '------------------------------
    If PUB_sqlserverConnection <> "" Then
        If MsgBox("�����R�l�N�V�������㏑�����܂����H", vbOKCancel, "�㏑���m�F") = vbCancel Then
            Exit Sub
        End If
    End If
    
    Set PUB_sqlserverConnection = Nothing
    form_�ڑ����o�^.Show
    
    
End Sub


Sub SQL���s()
    '------------------------------
    ' �ϐ���`
    '------------------------------
    Dim selectedSQLFile As String
    Dim execSQLStatement() As String
    Dim sqlParts As Variant
    Dim addressDrawStart As String
    addressDrawStart = "B13"
    Dim addressDrawEnd As String
    
    '------------------------------
    ' �R�l�N�V�����m�F
    '------------------------------
    If PUB_sqlserverConnection = "" Then
        MsgBox "�ڑ�����o�^���Ă��������B"
        Call �ڑ����o�^
        Exit Sub
    End If
    
    '------------------------------
    ' SQL�X�N���v�g�I��
    '------------------------------
    ChDir ThisWorkbook.Path
    selectedSQLFile = Application.GetOpenFilename("SQL�X�N���v�g,*.sql")
    
    '--- �I���`�F�b�N ---
    If selectedSQLFile = "False" Then
        MsgBox "�t�@�C���I��������܂���ł����B", vbOKOnly, "���s�s��"
        Exit Sub
    End If
    
    
On Error GoTo TAG_ERROR

    '------------------------------
    ' SQL�X�N���v�g������擾
    '------------------------------
    execSQLStatement = SQL�X�N���v�g�ǂݍ���(selectedSQLFile)
    
    '------------------------------
    ' SQL�X�N���v�g���s�E���ʕ`��
    '------------------------------
    For Each sqlParts In execSQLStatement
        If Len(sqlParts) = 0 Then
            GoTo Continue
        End If
        
        Call SQLServer�X�N���v�g���s(sqlParts)
        
        addressDrawEnd = Select���ʕ`��(addressDrawStart)
        
        addressDrawStart = Range(addressDrawEnd).Offset(2, 0).Address

Continue:
    Next sqlParts
    
    '------------------------------
    ' �I������
    '------------------------------
    MsgBox _
        "���s�������܂����B" _
        , vbOKOnly _
        , "���s�����FSQL���s"
    
    Exit Sub
    
TAG_ERROR:
    MsgBox _
        "���s���ɃG���[���������܂����B���e���m�F���Ă��������B" & vbCrLf & vbCrLf & _
        Err.Description _
        , vbCritical _
        , "�G���[�FSQL���s"
    
End Sub


Sub �ڑ����j��()
    Call SQLServer�ڑ��I��
End Sub



Sub �`�惊�Z�b�g()
    Rows("13:" & Cells.Rows.Count).Delete
End Sub

