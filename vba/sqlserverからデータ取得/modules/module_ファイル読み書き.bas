Attribute VB_Name = "module_�t�@�C���ǂݏ���"
Option Explicit

'=============================================================
' �X�N���v�g�ǂݍ���
' �I���p�X�ɑ΂��X�N���v�g��ǂݍ��ݕ������Ԃ�
'=============================================================
Function SQL�X�N���v�g�ǂݍ���(ByVal filePath As String) As String()
    
    '----------------------------------------
    ' �ϐ�
    '----------------------------------------
    Dim statementDelimiter As String
    statementDelimiter = ";"
    Dim readText As String
    Dim returnStr() As String
    Dim objFile As Object
    
On Error GoTo TAG_ERROR
    
    '----------------------------------------
    ' �t�@�C����������
    '----------------------------------------
    With CreateObject("Scripting.FileSystemObject")
        Set objFile = .OpenTextFile(filePath)
        
        With objFile
            Do Until .AtEndOfStream
                readText = readText & .ReadLine
            Loop
        End With
        
    End With
    
    returnStr = Split(readText, statementDelimiter)
    
    SQL�X�N���v�g�ǂݍ��� = returnStr
    
    Exit Function
    
    
TAG_ERROR:
    
    MsgBox _
        "�G���[�������܂����B" & vbCrLf & Err.Description _
        , vbCritical _
        , "�֐��F�X�N���v�g����"
        
End Function


'=============================================================
' �X�N���v�g����
' �I���p�X�ɑ΂��X�N���v�g���㏑���ۑ�����
'=============================================================
Function �X�N���v�g�㏑���ۑ�(ByVal filePath As String, ByVal writeStr As String, ByVal charCode As String)
    
    '----------------------------------------
    ' �ϐ�
    '----------------------------------------
    
    
    
On Error GoTo TAG_ERROR
    
    '----------------------------------------
    ' �t�@�C����������
    '----------------------------------------
    With CreateObject("ADODB.Stream")
        .Charset = charCode
        .Open
        
        .WriteText writeStr
        
        .SaveToFile filePath, 2
        
        .Close
    End With
    
    MsgBox "�ȉ��t�@�C���ւ̏������݂��������܂����B" & vbCr & filePath
    
    Exit Function
    
    
TAG_ERROR:
    
    MsgBox "�G���[�������܂����B", vbCritical, "�֐��F�X�N���v�g����"
    
End Function

