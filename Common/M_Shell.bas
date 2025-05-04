Attribute VB_Name = "M_Shell"
Option Explicit
'##############################################################################
' 外部プログラム処理
'##############################################################################
' 参照設定          |   Windows Script Host Object Model
'------------------------------------------------------------------------------
' 参照モジュール    |   M_String
'------------------------------------------------------------------------------
' 共通バージョン    |   ―
'------------------------------------------------------------------------------

'==============================================================================
' 公開定義
'==============================================================================

'==============================================================================
' 内部定義
'==============================================================================

'==============================================================================
' 公開処理
'==============================================================================
' 実行結果取得
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 一括返却
'------------------------------------------------------------------------------
Public Function F_Shell_GetResult( _
        ByRef aRtn As String, _
        ByVal aCmd As String) As Boolean
    Dim wkRet As Boolean: wkRet = False
    Dim wkRtn As String
    
    Dim wkShell As New WshShell
    Dim wkExec As Object
    
    '引数チェック
    If aCmd = "" Then
        Exit Function
    End If
    
    On Error GoTo PROC_ERROR
    'コマンド実行
    Set wkExec = wkShell.Exec(aCmd)
    
    wkRtn = wkExec.StdOut.ReadAll
    On Error GoTo 0

    If wkRtn <> "" Then
        aRtn = wkRtn
        wkRet = True
    End If

PROC_ERROR:
    F_Shell_GetResult = wkRet
End Function

'------------------------------------------------------------------------------
' 配列返却
'------------------------------------------------------------------------------
Public Function F_Shell_GetResultArray( _
        ByRef aRtnAry As Variant, _
        ByVal aCmd As String) As Boolean
    Dim wkRtnAry As Variant
    Dim wkRtnStr As String
    
    Dim wkTmpAry As Variant, wkTmp As Variant
    Dim wkTmpAry2 As Variant, wkTmp2 As Variant
    
    'コマンドの結果がNGであれば終了
    If F_Shell_GetResult(wkRtnStr, aCmd) <> True Then
        Exit Function
    End If
    
    'CRLF区切りで分割
    If M_String.F_String_GetSplit(wkTmpAry, wkRtnStr, vbCrLf) = True Then
        For Each wkTmp In wkTmpAry
            '空行でなければ追加
            If wkTmp <> "" Then
                wkRtnAry = M_Common.F_ReturnArrayAdd(wkRtnAry, wkTmp)
            End If
        Next wkTmp
    End If
    
    If IsArray(wkRtnAry) = True Then
        aRtnAry = wkRtnAry
        F_Shell_GetResultArray = True
    End If
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 実行
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_Shell_Run( _
        ByVal aCmd As String, _
        Optional ByVal aWindowStyle As Long = 0, _
        Optional ByVal aWaitOnReturn As Boolean = False) As Boolean
    Dim wkRet As Boolean: wkRet = False
    Dim wkRtn As String
    
    Dim wkShell As New WshShell
    Dim wkCmdRet As Integer
    
    '引数チェック
    If aCmd = "" Then
        Exit Function
    End If
    
    On Error GoTo PROC_ERROR
    'コマンド実行
    wkCmdRet = wkShell.Run(aCmd, WindowStyle:=aWindowStyle, WaitOnReturn:=aWaitOnReturn)
    If aWaitOnReturn = True Then
        '正常終了している場合は正常を返却
        If wkCmdRet = 0 Then
            wkRet = True
        End If
    Else
        '一旦正常とする
        wkRet = True
    End If
    On Error GoTo 0
    
PROC_ERROR:
    F_Shell_Run = wkRet
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 指定プログラム実行
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_Shell_OpenFile( _
        ByVal aFilePath As String, _
        ByVal aExePath As String, _
        Optional ByVal aExeArg As String = "", _
        Optional ByVal aWindowStyle As VbAppWinStyle = vbNormalFocus) As Boolean
    Dim wkRet As Boolean: wkRet = False
    Dim wkCmd As String
    
    '引数チェック
    If Dir(aFilePath) = "" Or Dir(aExePath) = "" Then
        Exit Function
    End If
    
    wkCmd = aExePath
    wkCmd = M_String.F_String_ReturnAdd(wkCmd, aExeArg, aDlmt:=" ")
    wkCmd = M_String.F_String_ReturnAdd(wkCmd, aFilePath, aDlmt:=" ")
    
    Shell wkCmd, aWindowStyle
    
    F_Shell_OpenFile = True
End Function
