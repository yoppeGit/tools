VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_�ڑ����o�^ 
   Caption         =   "�ڑ����o�^"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7950
   OleObjectBlob   =   "form_�ڑ����o�^.frx":0000
   StartUpPosition =   1  '�I�[�i�[ �t�H�[���̒���
End
Attribute VB_Name = "form_�ڑ����o�^"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub buttonRegist_Click()
    '------------------------------
    ' �ϐ���`
    '------------------------------
    Dim checkErrMsg As String
    Dim sqlserverObject As New ADODB.Connection
    Dim method As String

    '------------------------------
    ' ���͒l�`�F�b�N
    '------------------------------
    If textboxServer.Value = "" Then
        checkErrMsg = checkErrMsg & "�T�[�o����͂��Ă��������B" & vbCrLf
    End If
    
    If radioMethodSQLServer.Value = True Then
        If textboxUserName.Value = "" Then
            checkErrMsg = checkErrMsg & "���[�U������͂��Ă��������B" & vbCrLf
        End If
        If textboxPassword.Value = "" Then
            checkErrMsg = checkErrMsg & "�p�X���[�h����͂��Ă��������B" & vbCrLf
        End If
        
    End If
    
    If textboxDBName.Value = "" Then
        checkErrMsg = checkErrMsg & "DB������͂��Ă��������B" & vbCrLf
    End If
    
    If checkErrMsg <> "" Then
        MsgBox _
            checkErrMsg _
            , vbCritical _
            , "���͒l�G���["
        Exit Sub
    End If
    
    
    '------------------------------
    ' �e�X�g�ڑ�
    '------------------------------
    If radioMethodWindows.Value = True Then
        method = "windows"
    ElseIf radioMethodSQLServer.Value = True Then
        method = "sqlserver"
    End If
    
    
    Set PUB_sqlserverConnection = SQLServer�ڑ��J�n( _
            textboxServer.Value _
            , method _
            , textboxUserName.Value _
            , textboxPassword.Value _
            , textboxDBName.Value _
        )
        
    If PUB_sqlserverConnection = "" Then
        Exit Sub
    End If
    
    
    Unload form_�ڑ����o�^
       
End Sub


Private Sub radioMethodSQLServer_Click()
    textboxUserName.Enabled = True
    textboxPassword.Enabled = True
    
    textboxUserName.BackColor = &HFFFFFF
    textboxPassword.BackColor = &HFFFFFF
End Sub

Private Sub radioMethodWindows_Click()
    textboxUserName.Enabled = False
    textboxPassword.Enabled = False
        
    textboxUserName.BackColor = &HE0E0E0
    textboxPassword.BackColor = &HE0E0E0
End Sub
