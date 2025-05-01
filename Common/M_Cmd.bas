Attribute VB_Name = "M_Cmd"
Option Explicit
'##############################################################################
' コマンド処理
'##############################################################################
' 参照設定          |   Windows Script Host Object Model
'------------------------------------------------------------------------------
' 参照モジュール    |   M_String
'------------------------------------------------------------------------------

'==============================================================================
' 公開定義
'==============================================================================
' 定数定義
'------------------------------------------------------------------------------
Public Enum E_CMD_IDX_GREP_INF
    E_CMD_IDX_GREP_INF_NONE = D_IDX_START - 1
    E_CMD_IDX_GREP_INF_RESULT
    E_CMD_IDX_GREP_INF_FULLPATH
    E_CMD_IDX_GREP_INF_RLTPATH
    E_CMD_IDX_GREP_INF_FILE
    E_CMD_IDX_GREP_INF_LINE
    E_CMD_IDX_GREP_INF_OFFSET
    E_CMD_IDX_GREP_INF_SOURCE
    E_CMD_IDX_GREP_INF_MAX
    E_CMD_IDX_GREP_INF_EEND = E_CMD_IDX_GREP_INF_MAX - 1
End Enum

'------------------------------------------------------------------------------
' 構造体定義
'------------------------------------------------------------------------------
Public Type T_CMD_ARG_GREP_INF
    '検索文字列
    SrchPtn As String
    
    '単語チェック指定
    ChkWordFlg As Boolean
    ChkWordPtn As String
    ChkWordSpec As E_STRING_SPEC
    
    FullPath As String
    ExtSpec As String
End Type

'==============================================================================
' 内部定義
'==============================================================================
' 定数定義
'------------------------------------------------------------------------------
Private Enum PE_CMD_POS_GREPRET
    PE_CMD_POS_GREPRET_RLTPATH = 0
    PE_CMD_POS_GREPRET_LINE
    PE_CMD_POS_GREPRET_OFFSET
    PE_CMD_POS_GREPRET_SOURCE
    PE_CMD_POS_GREPRET_MAX
    PE_CMD_POS_GREPRET_EEND = PE_CMD_POS_GREPRET_MAX - 1
End Enum

'==============================================================================
' 公開処理
'==============================================================================
' コマンド実行結果取得
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' コマンド実行結果取得
'------------------------------------------------------------------------------
Public Function F_Cmd_GetCommandResult( _
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
    Set wkExec = wkShell.Exec("%ComSpec% /c " & aCmd)
    
    wkRtn = wkExec.StdOut.ReadAll
    On Error GoTo 0

    If wkRtn <> "" Then
        aRtn = wkRtn
        wkRet = True
    End If

PROC_ERROR:
    F_Cmd_GetCommandResult = wkRet
End Function

'------------------------------------------------------------------------------
' コマンド実行結果配列取得
'------------------------------------------------------------------------------
Public Function F_Cmd_GetCommandResultArray( _
        ByRef aRtnAry As Variant, _
        ByVal aCmd As String) As Boolean
    Dim wkRtnAry As Variant
    Dim wkRtnStr As String
    
    Dim wkTmpAry As Variant, wkTmp As Variant
    Dim wkTmpAry2 As Variant, wkTmp2 As Variant
    
    'コマンドの結果がNGであれば終了
    If F_Cmd_GetCommandResult(wkRtnStr, aCmd) <> True Then
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
        F_Cmd_GetCommandResultArray = True
    End If
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' コマンド実行
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_Cmd_RunCommand( _
        ByVal aCmd As String, _
        Optional ByVal aWindowStype As Long = 0, _
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
    wkCmdRet = wkShell.Run(aCmd, WindowStyle:=aWindowStype, WaitOnReturn:=aWaitOnReturn)
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
    F_Cmd_RunCommand = wkRet
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Grep結果取得
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 引数情報初期化
'------------------------------------------------------------------------------
Public Property Get G_Cmd_InitArgGrepInf() As T_CMD_ARG_GREP_INF
    With G_Cmd_InitArgGrepInf
        .SrchPtn = ""
        
        .ChkWordFlg = False
        .ChkWordPtn = D_STRING_MATCH_CHECKWORD
        .ChkWordSpec = E_STRING_SPEC_MATCH_ALL
        
        .FullPath = ""
        .ExtSpec = ""
    End With
End Property

'------------------------------------------------------------------------------
' Grep結果取得（引数情報指定）
'------------------------------------------------------------------------------
Public Function F_Cmd_GetTextGrepResult_Inf( _
        ByRef aRtnAryAry As Variant, _
        ByRef aArgInf As T_CMD_ARG_GREP_INF) As Boolean
    Dim wkRet As Boolean
    Dim wkRtnAryAry As Variant
    
    Dim wkGrepWd As String
    Dim wkGrepPath As String
    Dim wkAddInf As T_STRING_ARG_ADD_INF: wkAddInf = M_String.G_String_InitArgAddInf
    Dim wkSrchPtn As String
    Dim wkChkPtn As String
    
    Dim wkGrepRetAry As Variant, wkGrepRet As Variant
    Dim wkGrepRetInf(D_IDX_START To E_CMD_IDX_GREP_INF_EEND) As Variant
    Dim wkRltGrepRet As String
    Dim wkFullPath As String
    Dim wkSrc As String
    Dim wkSrchInf As T_STRING_ARG_SEARCH_INF: wkSrchInf = M_String.G_String_InitArgSrchInf
    
    Dim wkCmd As String
    Dim wkExtSpecAry As Variant, wkExtSpec As Variant
    
    Dim wkTmpStr As String
    Dim wkTmpAry As Variant, wkTmp As Variant
    Dim wkTmpFlg As Boolean
    Dim wkTmpCnt As Long
    
    With aArgInf
        '引数チェック
        If .SrchPtn = "" Or .FullPath = "" Or Dir(.FullPath, vbDirectory) = "" Or .ExtSpec = "" Then
            Exit Function
        End If
        
        'Grepパス作成（一旦全ファイル対象で検索）
        With wkAddInf
            .Target = aArgInf.FullPath
            .Add = "*"
            .Dlmt = "\"
            .DlmtChkFlg = True
        End With
        wkGrepPath = M_String.F_String_ReturnAdd_Inf(wkAddInf)
        
        'Grep実施
        wkSrchPtn = PF_Cmd_ReturnWildCardRegExp2FindStr(.SrchPtn)
        wkCmd = "findstr /r /s /n /p /o """ & wkSrchPtn & """ """ & wkGrepPath
        If F_Cmd_GetCommandResultArray(wkGrepRetAry, wkCmd) <> True Then
            Exit Function
        End If
        
        '拡張子指定を分割
        If M_String.F_String_GetSplitExtension(wkExtSpecAry, .ExtSpec) <> True Then
            wkExtSpecAry = M_Common.F_ReturnArrayAdd(wkExtSpecAry, "*")
        End If
        
        '初期化
        With wkSrchInf
            .SrchPtn = aArgInf.SrchPtn
            '単語検索ありの場合は単語チェック追加
            If aArgInf.ChkWordFlg = True Then
                .ChkPtn = D_STRING_MATCH_CHECKWORD
                .ChkSpec = aArgInf.ChkWordSpec Or E_STRING_SPEC_MATCH_WORD
            End If
        End With
        
        'Grep結果分ループ
        For Each wkGrepRet In wkGrepRetAry
            wkRltGrepRet = M_String.F_String_ReturnDelete(wkGrepRet, .FullPath, E_STRING_SPEC_POS_START)
            wkRltGrepRet = M_String.F_String_ReturnDelete(wkRltGrepRet, "\", E_STRING_SPEC_POS_START)
            
            If M_String.F_String_GetSplit(wkTmpAry, wkRltGrepRet, ":", aIncChkFlg:=True) <> True Then
                '分割できなかった場合は無視
            ElseIf UBound(wkTmpAry) < PE_CMD_POS_GREPRET_EEND Then
                '配列数が少なかった場合は無視
            Else
                'ソース確認
                wkSrc = ""
                For wkTmpCnt = PE_CMD_POS_GREPRET_SOURCE To UBound(wkTmpAry)
                    wkSrc = M_String.F_String_ReturnAdd(wkSrc, wkTmpAry(wkTmpCnt), aDlmt:=":")
                Next wkTmpCnt
                
                '単語チェックあり時は単語チェック実施
                If .ChkWordFlg = True Then
                    wkSrchInf.Target = wkSrc
                    wkTmpFlg = M_String.F_String_Check_Inf(wkSrchInf)
                Else
                    wkTmpFlg = True
                End If
                If wkTmpFlg = True Then
                    wkFullPath = .FullPath & "\" & wkTmpAry(PE_CMD_POS_GREPRET_RLTPATH)
                    
                    '拡張子検索
                    wkTmpFlg = False
                    For Each wkExtSpec In wkExtSpecAry
                        If wkTmpAry(PE_CMD_POS_GREPRET_RLTPATH) Like wkExtSpec Then
                            wkTmpFlg = True
                            Exit For
                        End If
                    Next wkExtSpec
                    
                    '拡張子が一致した場合はGrep結果を生成して戻り値に登録
                    If wkTmpFlg = True Then
                        '情報クリア
                        Erase wkGrepRetInf
                        
                        wkGrepRetInf(E_CMD_IDX_GREP_INF_RESULT) = wkGrepRet
                        wkGrepRetInf(E_CMD_IDX_GREP_INF_RLTPATH) = wkTmpAry(PE_CMD_POS_GREPRET_RLTPATH)
                        wkGrepRetInf(E_CMD_IDX_GREP_INF_FULLPATH) = wkFullPath
                        wkGrepRetInf(E_CMD_IDX_GREP_INF_FILE) = M_String.F_String_ReturnDeleteStr(wkGrepRetInf(E_CMD_IDX_GREP_INF_RLTPATH), "\", aDelPosSpec:=E_STRING_SPEC_POS_START)
                        wkGrepRetInf(E_CMD_IDX_GREP_INF_LINE) = Val(wkTmpAry(PE_CMD_POS_GREPRET_LINE))
                        wkGrepRetInf(E_CMD_IDX_GREP_INF_OFFSET) = Val(wkTmpAry(PE_CMD_POS_GREPRET_OFFSET))
                        wkGrepRetInf(E_CMD_IDX_GREP_INF_SOURCE) = wkSrc
                        
                        wkRtnAryAry = M_Common.F_ReturnArrayAdd(wkRtnAryAry, wkGrepRetInf)
                    End If
                End If
            End If
        Next wkGrepRet
    End With
    
    If IsArray(wkRtnAryAry) = True Then
        aRtnAryAry = wkRtnAryAry
        F_Cmd_GetTextGrepResult_Inf = True
    End If
End Function

'------------------------------------------------------------------------------
' Grep結果取得（引数指定）
'------------------------------------------------------------------------------
Public Function F_Cmd_GetTextGrepResult( _
        ByRef aRtnAryAry As Variant, _
        ByVal aSrchPtn As String, _
        ByVal aFullPath As String, _
        Optional ByVal aExtSpec As String = "*", _
        Optional ByVal aChkWordFlg As Boolean = True) As Boolean
    Dim wkArgInf As T_CMD_ARG_GREP_INF: wkArgInf = G_Cmd_InitArgGrepInf()
    
    With wkArgInf
        .SrchPtn = aSrchPtn
        .FullPath = aFullPath
        .ExtSpec = aExtSpec
        .ChkWordFlg = aChkWordFlg
    End With
    
    F_Cmd_GetTextGrepResult = F_Cmd_GetTextGrepResult_Inf(aRtnAryAry, wkArgInf)
End Function

'==============================================================================
' 内部処理
'==============================================================================
' ワイルドカード変換
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' RepExp→findstr
'------------------------------------------------------------------------------
Private Function PF_Cmd_ReturnWildCardRegExp2FindStr( _
        ByVal aSrchPtn As String) As String
    Dim wkRtn As String: wkRtn = aSrchPtn
    Dim wkTmpAry As Variant, wkTmp As Variant
    
    If M_String.F_String_GetSplit(wkTmpAry, wkRtn, "|", aIncChkFlg:=True) = True Then
        wkRtn = ""
        
        For Each wkTmp In wkTmpAry
            If wkTmp <> "" Then
                wkRtn = wkRtn & " " & wkTmp
            End If
        Next wkTmp
    End If
    
    PF_Cmd_ReturnWildCardRegExp2FindStr = wkRtn
End Function
