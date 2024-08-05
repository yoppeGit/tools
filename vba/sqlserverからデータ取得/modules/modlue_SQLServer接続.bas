Attribute VB_Name = "modlue_SQLServer接続"
Option Explicit

Function SQLServer接続開始( _
    ByVal server As String _
    , ByVal method As String _
    , ByVal userName As String _
    , ByVal password As String _
    , ByVal dbName As String _
) As Object
    
On Error GoTo TAG_ERROR
    
    '-------------------------------
    ' 変数定義
    '-------------------------------
    Dim sqlserverObject As ADODB.Connection
    Set sqlserverObject = New ADODB.Connection
    Dim connectionStr As String
        
    
    '-------------------------------
    ' 接続の実行
    '-------------------------------
    '--- 接続文字列の構築 ---
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
        
    '--- 接続開始 ---
    sqlserverObject.Open connectionStr
    MsgBox _
        "DBの接続に成功しました。"
        
    
    '-------------------------------
    ' return
    '-------------------------------
    Set SQLServer接続開始 = sqlserverObject
    
    Exit Function
    
TAG_ERROR:
    
    MsgBox _
        "SQLServerの接続に失敗しました。" _
        , vbCritical _
        , "エラー：SQLServer接続開始"
    
    Set SQLServer接続開始 = New ADODB.Connection
    
End Function


Function SQLServerスクリプト実行(ByVal execSQL As String)
    
    '-------------------------------
    ' 変数定義
    '-------------------------------
    Set PUB_sqlserverRecordSet = CreateObject("ADODB.RecordSet")
    
    '-------------------------------
    ' 実行前判断
    '-------------------------------
    If Len(execSQL) = 0 Then
        MsgBox "SQL実行文字列長が0です。", vbCritical, "実行不可：SQLServerスクリプト実行"
        Exit Function
    End If
    
    
On Error GoTo TAG_ERROR
    '-------------------------------
    ' 実行
    '-------------------------------
    PUB_sqlserverRecordSet.Open _
        execSQL _
        , PUB_sqlserverConnection _
        , adOpenStatic, adLockReadOnly, adCmdText
    
    Exit Function
    
TAG_ERROR:
    MsgBox _
        "エラー発生。内容を確認して下さい。" & vbCrLf & Err.Description _
        , vbCritical _
        , "エラー：SQLServerスクリプト実行"
    
End Function


Function Select結果描画(ByVal addressDrawStart As String) As String
    '-------------------------------
    ' Excel結果描画
    '-------------------------------
    Dim roopHeader As Integer
    Dim headerString As String

On Error GoTo TAG_ERROR
    
    'レコードセットへ先頭へ
    PUB_sqlserverRecordSet.MoveFirst
   
    '-------------------------------
    ' Excel結果描画
    '-------------------------------
    ' 取得SQL情報
    Range(addressDrawStart).Value = "取得SQL:"
    Range(addressDrawStart).Offset(0, 1).Value = PUB_sqlserverRecordSet.Source
    
    'ヘッダ書き出し
    For roopHeader = 1 To PUB_sqlserverRecordSet.Fields.Count
        ' 書き出し文字列の確定
        Select Case PUB_sqlserverRecordSet.Fields(roopHeader - 1).Name
            Case ""
                headerString = "Column" & roopHeader
            Case Else
                headerString = PUB_sqlserverRecordSet.Fields(roopHeader - 1).Name
        End Select
        
        ' ヘッダ文字列の記載
        Range(addressDrawStart).Offset(1, roopHeader - 1) = headerString
    Next
    
    'データ書き出し
    Range(addressDrawStart).Offset(2, 0).CopyFromRecordset PUB_sqlserverRecordSet
    
    
    '-------------------------------
    ' 戻り値の算出：描画最終アドレス番地
    '-------------------------------
    Select結果描画 = Range(addressDrawStart).Offset(PUB_sqlserverRecordSet.RecordCount + 1, 0).Address
    

    Exit Function
    
TAG_ERROR:
    MsgBox _
        "エラー発生。内容を確認して下さい。" & vbCrLf & Err.Description _
        , vbCritical _
        , "エラー：SQLServerスクリプト実行"
        
    Select結果描画 = addressDrawStart
        
End Function


Function SQLServer接続終了()
    If PUB_sqlserverConnection = "" Then
        Exit Function
    End If

    If MsgBox("既存コネクションを破棄しますか？", vbOKCancel, "上書き確認") = vbCancel Then
        Exit Function
    End If
    
    PUB_sqlserverConnection.Close
    Set PUB_sqlserverConnection = Nothing
    
    PUB_sqlserverRecordSet.Close
    Set PUB_sqlserverRecordSet = Nothing
    
    MsgBox "破棄しました。"
    
End Function



