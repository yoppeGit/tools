VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} form_接続情報登録 
   Caption         =   "接続情報登録"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7950
   OleObjectBlob   =   "form_接続情報登録.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "form_接続情報登録"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub buttonRegist_Click()
    '------------------------------
    ' 変数定義
    '------------------------------
    Dim checkErrMsg As String
    Dim sqlserverObject As New ADODB.Connection
    Dim method As String

    '------------------------------
    ' 入力値チェック
    '------------------------------
    If textboxServer.Value = "" Then
        checkErrMsg = checkErrMsg & "サーバを入力してください。" & vbCrLf
    End If
    
    If radioMethodSQLServer.Value = True Then
        If textboxUserName.Value = "" Then
            checkErrMsg = checkErrMsg & "ユーザ名を入力してください。" & vbCrLf
        End If
        If textboxPassword.Value = "" Then
            checkErrMsg = checkErrMsg & "パスワードを入力してください。" & vbCrLf
        End If
        
    End If
    
    If textboxDBName.Value = "" Then
        checkErrMsg = checkErrMsg & "DB名を入力してください。" & vbCrLf
    End If
    
    If checkErrMsg <> "" Then
        MsgBox _
            checkErrMsg _
            , vbCritical _
            , "入力値エラー"
        Exit Sub
    End If
    
    
    '------------------------------
    ' テスト接続
    '------------------------------
    If radioMethodWindows.Value = True Then
        method = "windows"
    ElseIf radioMethodSQLServer.Value = True Then
        method = "sqlserver"
    End If
    
    
    Set PUB_sqlserverConnection = SQLServer接続開始( _
            textboxServer.Value _
            , method _
            , textboxUserName.Value _
            , textboxPassword.Value _
            , textboxDBName.Value _
        )
        
    If PUB_sqlserverConnection = "" Then
        Exit Sub
    End If
    
    
    Unload form_接続情報登録
       
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
