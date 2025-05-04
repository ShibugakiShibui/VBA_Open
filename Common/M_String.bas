Attribute VB_Name = "M_String"
Option Explicit
'##############################################################################
' 文字列処理
'##############################################################################
' 参照設定          |   Microsoft VBScript Regular Expressions 5.5
'------------------------------------------------------------------------------
' 参照モジュール    |   ―
'------------------------------------------------------------------------------
' 共通バージョン    |   250501
'------------------------------------------------------------------------------

'==============================================================================
' 公開定義
'==============================================================================
' 定数定義
'------------------------------------------------------------------------------
Public Enum E_STRING_SPEC
    E_STRING_SPEC_NONE = &H0
    E_STRING_SPEC_POS_START = &H1
    E_STRING_SPEC_POS_END = &H2
    E_STRING_SPEC_POS_MID = &H4
    E_STRING_SPEC_POS_BOTH = E_STRING_SPEC_POS_START Or E_STRING_SPEC_POS_END
    E_STRING_SPEC_MATCH_START = &H10
    E_STRING_SPEC_MATCH_END = &H20
    E_STRING_SPEC_MATCH_MID = &H40
    E_STRING_SPEC_MATCH_WORD = &H80
    E_STRING_SPEC_MATCH_ALL = E_STRING_SPEC_MATCH_START Or E_STRING_SPEC_MATCH_END
    E_STRING_SPEC_MATCH_MASK = E_STRING_SPEC_MATCH_START Or E_STRING_SPEC_MATCH_END Or E_STRING_SPEC_MATCH_MID
End Enum

Public Enum E_STRING_IDX_SRCH_INF
    E_STRING_IDX_SRCH_INF_NONE = D_IDX_START - 1
    E_STRING_IDX_SRCH_INF_POS_START
    E_STRING_IDX_SRCH_INF_LENGTH
    E_STRING_IDX_SRCH_INF_MAX
    E_STRING_IDX_SRCH_INF_EEND = E_STRING_IDX_SRCH_INF_MAX - 1
End Enum

Public Const D_STRING_MATCH_CHECKWORD As String = "[^A-Za-z0-9_]"

Public Const D_STRING_DLMT_EXTENSION As String = ";"

'------------------------------------------------------------------------------
' 構造体定義
'------------------------------------------------------------------------------
Public Type T_STRING_ARG_ADD_INF
    '文字列
    Target As String
    
    '区切り
    Dlmt As String
    DlmtChkFlg As Boolean
    
    '追加
    Add As String
    AddChkFlg As Boolean
    AddSpec As E_STRING_SPEC
    
    '対象外
    Excluded As String
End Type

Public Type T_STRING_ARG_SEARCH_INF
    '文字列
    Target As String
    
    '検索指定
    Search As String
    SrchSpec As E_STRING_SPEC
    SrchPtn As String
    
    '一致指定
    ChkPtn As String
    ChkSpec As E_STRING_SPEC
    ChkPtnOfs As Long
    
    '検索位置指定
    SttPos As Long
    EndPos As Long
    Length As Long
    
    '取得位置指定
    GetIdx As Long
End Type

Public Type T_STRING_ARG_GET_INF
    SrchInf As T_STRING_ARG_SEARCH_INF
    
    SttStr As String
    EndStr As String
    
    AddBefFlg As Boolean
    AddSrchFlg As Boolean
End Type

Public Type T_STRING_ARG_DEL_INF
    SrchInf As T_STRING_ARG_SEARCH_INF
    
    DelPosSpec As E_STRING_SPEC
    AddDelFlg As Boolean
End Type

'==============================================================================
' 内部定義
'==============================================================================

'==============================================================================
' 公開処理
'==============================================================================
' 文字列追加
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 引数情報初期化
'------------------------------------------------------------------------------
Public Property Get G_String_InitArgAddInf() As T_STRING_ARG_ADD_INF
    Dim wkInf As T_STRING_ARG_ADD_INF
    
    With wkInf
        .Target = ""
        
        .Dlmt = ""
        .DlmtChkFlg = True
        
        .Add = ""
        .AddChkFlg = False
        .AddSpec = E_STRING_SPEC_POS_END
        
        .Excluded = ""
    End With
    
    G_String_InitArgAddInf = wkInf
End Property

'------------------------------------------------------------------------------
' 文字列追加（引数情報指定）
'------------------------------------------------------------------------------
Public Function F_String_ReturnAdd_Inf( _
        ByRef aArgInf As T_STRING_ARG_ADD_INF) As String
    Dim wkRtn As String
    
    Dim wkAddChkFlg As Boolean
    Dim wkTmpStr As String
    
    With aArgInf
        '初期化
        wkRtn = .Target
        wkAddChkFlg = .AddChkFlg
        
        '追加ありの場合
        If .Add <> "" Then
            '区切り文字接続
            If wkRtn = "" Then
                '追加元が無い場合は無視
            ElseIf PF_String_GetAdd_Sub(wkRtn, .Dlmt, .AddSpec, .DlmtChkFlg, .Excluded) = True Then
                '区切り追加した場合は追加チェックを有効化
                wkAddChkFlg = True
            End If
            
            '追加文字接続
            PF_String_GetAdd_Sub wkRtn, .Add, .AddSpec, wkAddChkFlg, .Excluded
        End If
    End With
    
    F_String_ReturnAdd_Inf = wkRtn
End Function

' サブルーチン
Private Function PF_String_GetAdd_Sub( _
        ByRef aRtn As String, _
        ByVal aAdd As String, ByVal aAddSpec As E_STRING_SPEC, _
        ByVal aAddChkFlg As Boolean, ByVal aExcluded As String) As Boolean
    Dim wkRet As Boolean: wkRet = True
    
    Dim wkChkStr As String
    Dim wkExcStr As String
    
    If aAdd <> "" Then
        If aAddChkFlg = True Then
            '追加位置が後方指定の場合
            If aAddSpec = E_STRING_SPEC_POS_END Then
                wkChkStr = Right(aRtn, Len(aAdd))
                If aExcluded <> "" Then
                    wkExcStr = Right(aRtn, Len(aExcluded))
                End If
            '追加位置が前方指定の場合
            Else
                wkChkStr = Left(aRtn, Len(aAdd))
                If aExcluded <> "" Then
                    wkExcStr = Left(aRtn, Len(aExcluded))
                End If
            End If
            
            '接続位置に文字がある場合
            If StrComp(wkChkStr, aAdd, vbBinaryCompare) = 0 Then
                '対象外文字が無い場合は追加対象外
                If aExcluded = "" Then
                    wkRet = False
                '対象外文字と不一致の場合は追加対象外
                ElseIf StrComp(wkExcStr, aExcluded, vbBinaryCompare) <> 0 Then
                    wkRet = False
                End If
            End If
        End If
        
        '追加ありの場合
        If wkRet = True Then
            '追加位置指定に従って文字列追加
            If aAddSpec = E_STRING_SPEC_POS_END Then
                aRtn = aRtn & aAdd
            Else
                aRtn = aAdd & aRtn
            End If
        End If
    End If
    
    PF_String_GetAdd_Sub = wkRet
End Function

'------------------------------------------------------------------------------
' 文字列追加（引数指定）
'------------------------------------------------------------------------------
Public Function F_String_ReturnAdd( _
        ByVal aTarget As String, ByVal aAdd As String, _
        Optional ByVal aDlmt As String = "", _
        Optional ByVal aAddSpec As E_STRING_SPEC = E_STRING_SPEC_POS_END, _
        Optional ByVal aExcluded As String = "") As String
    Dim wkArgInf As T_STRING_ARG_ADD_INF: wkArgInf = G_String_InitArgAddInf()
    
    With wkArgInf
        .Target = aTarget
        .Dlmt = aDlmt
        .Add = aAdd
        .AddSpec = aAddSpec
        .Excluded = aExcluded
    End With
    
    F_String_ReturnAdd = F_String_ReturnAdd_Inf(wkArgInf)
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 検索情報配列取得
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 引数情報初期化
'------------------------------------------------------------------------------
Public Property Get G_String_InitArgSrchInf() As T_STRING_ARG_SEARCH_INF
    Dim wkInf As T_STRING_ARG_SEARCH_INF
    
    With wkInf
        .Target = ""
        
        .Search = ""
        .SrchSpec = E_STRING_SPEC_MATCH_MID
        .SrchPtn = ""
        
        .ChkPtn = ""
        .ChkSpec = E_STRING_SPEC_MATCH_ALL
        .ChkPtnOfs = 1
        
        .SttPos = D_POS_START
        .EndPos = D_POS_END
        .Length = D_POS_END
        
        .GetIdx = D_IDX_ALL
    End With
    
    G_String_InitArgSrchInf = wkInf
End Property

'------------------------------------------------------------------------------
' 検索情報配列取得（引数情報指定）
'------------------------------------------------------------------------------
Public Function F_String_GetSearchInfArray_Inf( _
        ByRef aRtnAryAry As Variant, _
        ByRef aArgInf As T_STRING_ARG_SEARCH_INF) As Boolean
    Dim wkRtnAryAry As Variant
    Dim wkRtnAry(D_IDX_START To E_STRING_IDX_SRCH_INF_EEND) As Variant
    Dim wkAddFlg As Boolean
    
    Dim wkArgInf As T_STRING_ARG_SEARCH_INF: wkArgInf = aArgInf
    Dim wkSttPos As Long, wkEndPos As Long
    Dim wkTgtLen As Long
    
    Dim wkSrchPtn As String
    Dim wkSrchRegExp As RegExp
    Dim wkMatch As Variant
    Dim wkMatchCnt As Long
    
    Dim wkChkPtn As String
    Dim wkChkRegExp As RegExp
    Dim wkChkStt As Long, wkChkEnd As Long
    
    Dim wkGetIdx As Long
    Dim wkRtnIdx As Long
    
    On Error GoTo PROC_ERROR
    With aArgInf
        '引数チェック
        If .Target = "" Or (.Search = "" And .SrchPtn = "") Or .SttPos < D_POS_START Then
            Exit Function
        End If
        
        '終端位置調整
        wkSttPos = .SttPos
        wkEndPos = .EndPos
        wkTgtLen = Len(.Target)
        If PF_String_GetPosEndAdjust(wkEndPos, Len(.Target), wkSttPos, .Length) <> True Then
            Exit Function
        End If
        
        '検索パターン保持
        wkSrchPtn = .SrchPtn
        If wkSrchPtn = "" Then
            '検索パターンを作成
            wkSrchPtn = PF_String_ReturnSearchPattern(.Search, .SrchSpec)
        End If
        
        '検索パターン設定、検索実施
        Set wkSrchRegExp = New RegExp
        With wkSrchRegExp
            .IgnoreCase = False
            .Global = True
            .Pattern = wkSrchPtn
        End With
        Set wkMatch = wkSrchRegExp.Execute(Mid(.Target, wkSttPos, (wkEndPos - wkSttPos + 1)))
        '検索結果が無ければ終了
        If wkMatch.Count <= 0 Then
            Exit Function
        End If
        
        '一致パターンがある場合
        If .ChkPtn <> "" And .ChkPtnOfs > 0 Then
            'チェック指定がある場合
            If M_Common.F_CheckBitOn(.ChkSpec, E_STRING_SPEC_MATCH_MASK) = True Then
                '一致パターン作成
                If M_Common.F_CheckBitOn(.ChkSpec, E_STRING_SPEC_MATCH_WORD) = True Then
                    wkChkPtn = PF_String_ReturnCheckPatternWord(wkSrchPtn, .SrchSpec, .ChkPtn, .ChkSpec)
                Else
                    wkChkPtn = PF_String_ReturnCheckPattern(wkSrchPtn, .SrchSpec, .ChkPtn, .ChkSpec)
                End If
            End If
            
            'チェックパターンがある場合、チェックパターン設定
            If wkChkPtn <> "" Then
                'チェック検索設定
                Set wkChkRegExp = New RegExp
                With wkChkRegExp
                    .IgnoreCase = False
                    .Global = False
                    .Pattern = wkChkPtn
                End With
            End If
        End If
        
        '検索ヒット数分、位置情報を抽出
        wkGetIdx = D_IDX_START
        wkRtnIdx = D_IDX_START
        For wkMatchCnt = 0 To wkMatch.Count - 1
            '初期化
            wkAddFlg = True
            wkRtnAry(E_STRING_IDX_SRCH_INF_POS_START) = wkSttPos + wkMatch.Item(wkMatchCnt).FirstIndex
            wkRtnAry(E_STRING_IDX_SRCH_INF_LENGTH) = wkMatch.Item(wkMatchCnt).Length
                
            If Not wkChkRegExp Is Nothing Then
                'チェック位置調整
                wkChkStt = wkRtnAry(E_STRING_IDX_SRCH_INF_POS_START)
                wkChkEnd = wkChkStt + wkRtnAry(E_STRING_IDX_SRCH_INF_LENGTH) - 1
    
                '開始位置調整
                wkChkStt = wkChkStt - .ChkPtnOfs
                If wkChkStt < D_POS_START Then
                    wkChkStt = D_POS_START
                End If
                '終了位置調整
                wkChkEnd = wkChkEnd + .ChkPtnOfs
                If wkChkEnd > wkTgtLen Then
                    wkChkEnd = wkTgtLen
                End If
                    
                wkAddFlg = wkChkRegExp.Test(Mid(.Target, wkChkStt, (wkChkEnd - wkChkStt + 1)))
            End If
                
            If wkAddFlg = True Then
                '取得インデックスが設定なしまたは取得インデックスが一致の場合、戻り引数に設定
                If .GetIdx < D_IDX_START Or wkGetIdx = .GetIdx Then
                    '取得カウンタが全取得の場合は設定位置を調整
                    If .GetIdx = D_IDX_ALL Then
                        wkRtnIdx = wkGetIdx
                    End If
                    '取得結果を配列に登録
                    wkRtnAryAry = M_Common.F_ReturnArrayAdd(wkRtnAryAry, wkRtnAry, aIdx:=wkRtnIdx)
                    
                    'インデックスが一致の場合は取得完了のためループ終了
                    If wkGetIdx = .GetIdx Then
                        Exit For
                    End If
                End If
                
                '取得カウンタ更新
                wkGetIdx = wkGetIdx + 1
            End If
        Next wkMatchCnt
    End With
    On Error GoTo 0
    
    '指定パターンが見つかった場合
    If IsArray(wkRtnAryAry) = True Then
        aRtnAryAry = wkRtnAryAry
        F_String_GetSearchInfArray_Inf = True
    End If
    
PROC_ERROR:
    '何もしない
End Function

'------------------------------------------------------------------------------
' 検索情報配列取得（引数指定）
'------------------------------------------------------------------------------
Public Function F_String_GetSearchInfArray( _
        ByRef aRtnAryAry As Variant, _
        ByVal aTarget As String, _
        Optional ByVal aSearch As String = "", Optional ByVal aSrchSpec As E_STRING_SPEC = E_STRING_SPEC_MATCH_MID, Optional ByVal aSrchPtn As String = "") As Boolean
    Dim wkArgInf As T_STRING_ARG_SEARCH_INF: wkArgInf = G_String_InitArgSrchInf()
    
    With wkArgInf
        .Target = aTarget
        .Search = aSearch
        .SrchSpec = aSrchSpec
        .SrchPtn = aSrchPtn
    End With
    
    F_String_GetSearchInfArray = F_String_GetSearchInfArray_Inf(aRtnAryAry, wkArgInf)
End Function

'------------------------------------------------------------------------------
' 単語検索情報配列取得（引数指定）
'------------------------------------------------------------------------------
Public Function F_String_GetSearchInfArrayWord( _
        ByRef aRtnAryAry As Variant, _
        ByVal aTarget As String, _
        Optional ByVal aSearch As String = "", Optional ByVal aSrchSpec As E_STRING_SPEC = E_STRING_SPEC_MATCH_MID, Optional ByVal aSrchPtn As String = "", _
        Optional ByVal aChkPtn As String = D_STRING_MATCH_CHECKWORD, Optional ByVal aChkSpec As E_STRING_SPEC = E_STRING_SPEC_MATCH_ALL) As Boolean
    Dim wkArgInf As T_STRING_ARG_SEARCH_INF: wkArgInf = G_String_InitArgSrchInf()
    
    '引数情報に設定
    With wkArgInf
        .Target = aTarget
        .Search = aSearch
        .SrchSpec = aSrchSpec
        
        .ChkPtn = aChkPtn
        .ChkSpec = aChkSpec Or E_STRING_SPEC_MATCH_WORD
    End With
    
    F_String_GetSearchInfArrayWord = F_String_GetSearchInfArray_Inf(aRtnAryAry, wkArgInf)
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 文字列存在チェック
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 文字列存在チェック（引数情報指定）
'------------------------------------------------------------------------------
Public Function F_String_Check_Inf( _
        ByRef aArgInf As T_STRING_ARG_SEARCH_INF) As Boolean
    Dim wkTmpAryAry As Variant
    
    F_String_Check_Inf = F_String_GetSearchInfArray_Inf(wkTmpAryAry, aArgInf)
End Function

'------------------------------------------------------------------------------
' 文字列存在チェック（引数指定）
'------------------------------------------------------------------------------
Public Function F_String_Check( _
        ByVal aTarget As String, _
        Optional ByVal aSearch As String = "", Optional ByVal aSrchSpec As E_STRING_SPEC = E_STRING_SPEC_MATCH_MID, Optional ByVal aSrchPtn As String = "") As Boolean
    Dim wkArgInf As T_STRING_ARG_SEARCH_INF: wkArgInf = G_String_InitArgSrchInf()
    
    With wkArgInf
        .Target = aTarget
        
        .Search = aSearch
        .SrchSpec = aSrchSpec
        .SrchPtn = aSrchPtn
    End With
    
    F_String_Check = F_String_Check_Inf(wkArgInf)
End Function

'------------------------------------------------------------------------------
' 単語存在チェック（引数指定）
'------------------------------------------------------------------------------
Public Function F_String_CheckWord( _
        ByVal aTarget As String, _
        Optional ByVal aSearch As String = "", Optional ByVal aSrchSpec As E_STRING_SPEC = E_STRING_SPEC_MATCH_MID, Optional ByVal aSrchPtn As String = "", _
        Optional ByVal aChkPtn As String = D_STRING_MATCH_CHECKWORD, Optional ByVal aChkSpec As E_STRING_SPEC = E_STRING_SPEC_MATCH_ALL) As Boolean
    Dim wkTmpAryAry As Variant
    
    F_String_CheckWord = F_String_GetSearchInfArrayWord(wkTmpAryAry, aTarget, aSearch:=aSearch, aSrchSpec:=aSrchSpec, aSrchPtn:=aSrchPtn, _
                                                        aChkPtn:=aChkPtn, aChkSpec:=aChkSpec)

End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 文字列分割
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_String_GetSplit( _
        ByRef aRtnAry As Variant, _
        ByVal aTarget As String, ByVal aDlmt As String, _
        Optional ByVal aIncChkFlg As Boolean = False) As Boolean
    Dim wkRtnAry As Variant
    
    '引数チェック
    If aTarget = "" Or aDlmt = "" Then
        Exit Function
    End If
    
    '区切りで分割
    wkRtnAry = Split(aTarget, aDlmt)
    
    '区切り包含チェックありの場合、区切り無しはNGで終了
    If aIncChkFlg = True Then
        If UBound(wkRtnAry) <= LBound(wkRtnAry) Then
            Exit Function
        End If
    End If
    
    aRtnAry = wkRtnAry
    F_String_GetSplit = True
End Function

'------------------------------------------------------------------------------
' 拡張子指定分割
'------------------------------------------------------------------------------
Public Function F_String_GetSplitExtension( _
        ByRef aRtnAry As Variant, _
        ByVal aExtSpec As String) As Boolean
    Dim wkRtnAry As Variant
    Dim wkTmpAry As Variant, wkTmp As Variant
    
    '文字列分割（引数チェック兼用）
    If F_String_GetSplit(wkTmpAry, aExtSpec, D_STRING_DLMT_EXTENSION) <> True Then
        Exit Function
    End If
    
    '分割全てループ
    For Each wkTmp In wkTmpAry
        '空白でなければ追加
        wkTmp = Trim(wkTmp)
        If wkTmp <> "" Then
            wkRtnAry = M_Common.F_ReturnArrayAdd(wkTmpAry, wkTmp)
        End If
    Next wkTmp
    
    If IsArray(wkRtnAry) = True Then
        aRtnAry = wkRtnAry
        F_String_GetSplitExtension = True
    End If
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 指定文字列間取得
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 引数情報初期化
'------------------------------------------------------------------------------
Public Property Get G_String_InitArgGetInf() As T_STRING_ARG_GET_INF
    Dim wkInf As T_STRING_ARG_GET_INF
    
    With wkInf
        .SrchInf = G_String_InitArgSrchInf
        
        .SttStr = ""
        .EndStr = ""
        
        .AddBefFlg = False
        .AddSrchFlg = False
    End With
    
    G_String_InitArgGetInf = wkInf
End Property

'------------------------------------------------------------------------------
' 指定文字列間取得（引数情報指定）
'------------------------------------------------------------------------------
Public Function F_String_GetMidStr_Inf( _
        ByRef aRtn As String, _
        ByRef aArgInf As T_STRING_ARG_GET_INF) As Boolean
    Dim wkRet As Boolean
    Dim wkRtn As String
    
    Dim wkSttInfAry As Variant, wkSttInf As Variant
    Dim wkEndInfAry As Variant, wkEndInf As Variant
    Dim wkArgSrchInf As T_STRING_ARG_SEARCH_INF
    
    Dim wkSttCnt As Long, wkSttNow As Long
    Dim wkChkCnt As Long
    
    Dim wkGetSttInf As Variant
    Dim wkGetEndInf As Variant
    Dim wkGetSttPos As Long, wkGetEndPos As Long
    
    With aArgInf
        wkArgSrchInf = aArgInf.SrchInf
        
        wkArgSrchInf.Search = .SttStr
        '開始位置取得（引数チェック兼用）
        If F_String_GetSearchInfArray_Inf(wkSttInfAry, wkArgSrchInf) <> True Then
            Exit Function
        End If
        
        wkArgSrchInf.Search = .EndStr
        '終了位置取得（引数チェック兼用）
        If F_String_GetSearchInfArray_Inf(wkEndInfAry, wkArgSrchInf) <> True Then
            Exit Function
        End If
    End With
    
    '終了位置で全ループ
    wkSttNow = LBound(wkSttInfAry) - 1
    For Each wkEndInf In wkEndInfAry
        '開始位置でループ
        For wkSttCnt = wkSttNow + 1 To UBound(wkSttInfAry)
            wkSttInf = wkSttInfAry(wkSttCnt)
            
            '開始位置が終了位置以上の場合はループ終了
            If wkSttInf(E_STRING_IDX_SRCH_INF_POS_START) >= wkEndInf(E_STRING_IDX_SRCH_INF_POS_START) Then
                Exit For
            End If
            
            '各種更新
            wkSttNow = wkSttCnt
            wkChkCnt = wkChkCnt + 1
            
            '初回の場合は開始位置保持
            If IsArray(wkGetSttInf) <> True Then
                wkGetSttInf = wkSttInf
            End If
        Next wkSttCnt
        
        '開始位置が見つかっている場合は終了位置チェック
        If IsArray(wkGetSttInf) = True Then
            wkChkCnt = wkChkCnt - 1
            
            '整合が取れた場合はループ終了
            If wkChkCnt <= 0 Then
                wkGetEndInf = wkEndInf
                wkRet = True
                Exit For
            End If
        End If
    Next wkEndInf
    
    '取得あり時は文字列返却
    If wkRet = True Then
        With aArgInf
            '開始位置調整
            '検索前文字列追加指定ありの場合
            If .AddBefFlg = True Then
                wkGetSttPos = .SrchInf.SttPos
            '検索文字列追加指定なしの場合
            ElseIf .AddSrchFlg <> True Then
                wkGetSttPos = wkGetSttInf(E_STRING_IDX_SRCH_INF_POS_START) + wkGetSttInf(E_STRING_IDX_SRCH_INF_LENGTH)
            End If
        
            '終了位置調整
            '検索前文字列追加指定あり、または検索文字列追加指定ありの場合
            If .AddBefFlg = True Or .AddSrchFlg = True Then
                wkGetEndPos = wkGetEndInf(E_STRING_IDX_SRCH_INF_POS_START) - wkGetEndInf(E_STRING_IDX_SRCH_INF_LENGTH) + 1
            '文字列追加指定なしの場合
            Else
                wkGetEndPos = wkGetEndInf(E_STRING_IDX_SRCH_INF_POS_START) - 1
            End If
            
            If wkGetSttPos <= wkGetEndPos Then
                wkRtn = Mid(.SrchInf.Target, wkGetSttPos, (wkGetEndPos - wkGetSttPos + 1))
            End If
        End With
        
        If wkRtn <> "" Then
            aRtn = wkRtn
            F_String_GetMidStr_Inf = True
        End If
    End If
End Function

'------------------------------------------------------------------------------
' 指定文字列間取得（引数指定）
'------------------------------------------------------------------------------
Public Function F_String_GetMidStr( _
        ByRef aRtn As String, _
        ByVal aTarget As String, _
        ByVal aSttStr As String, ByVal aEndStr As String, _
        Optional ByVal aAddBefFlg As Boolean = False, _
        Optional ByVal aAddSrchFlg As Boolean = False) As Boolean
    Dim wkArgInf As T_STRING_ARG_GET_INF: wkArgInf = G_String_InitArgGetInf()
    
    With wkArgInf
        With .SrchInf
            .Target = aTarget
            .GetIdx = D_IDX_START
        End With
        
        .SttStr = aSttStr
        .EndStr = aEndStr
        .AddBefFlg = aAddBefFlg
        .AddSrchFlg = aAddSrchFlg
    End With
    
    F_String_GetMidStr = F_String_GetMidStr_Inf(aRtn, wkArgInf)
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 指定文字列削除
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_String_ReturnDelete( _
        ByVal aTarget As String, _
        ByVal aDelete As String, ByVal aDelSpec As E_STRING_SPEC) As String
    Dim wkRtn As String: wkRtn = aTarget
    Dim wkLen As Long
    Dim wkDelLen As Long
    
    '文字列と削除文字がある場合
    If aTarget <> "" And aDelete <> "" Then
        '中間位置削除の場合
        If M_Common.F_CheckBitOn(aDelSpec, E_STRING_SPEC_POS_MID) = True Then
            wkRtn = Replace(aTarget, aDelete, "")
        Else
            wkLen = Len(aTarget)
            wkDelLen = Len(aDelete)
            
            '開始位置削除指定の場合
            If M_Common.F_CheckBitOn(aDelSpec, E_STRING_SPEC_POS_START) = True Then
                '開始位置に削除文字がある間ループ
                Do While StrComp(Left(wkRtn, wkDelLen), aDelete, vbBinaryCompare) = 0
                    wkLen = wkLen - wkDelLen
                    wkRtn = Right(wkRtn, wkLen)
                Loop
            End If
            
            '終了位置削除指定の場合
            If M_Common.F_CheckBitOn(aDelSpec, E_STRING_SPEC_POS_END) = True Then
                '終了位置に削除文字がある間ループ
                Do While StrComp(Right(wkRtn, wkDelLen), aDelete, vbBinaryCompare) = 0
                    wkLen = wkLen - wkDelLen
                    wkRtn = Left(wkRtn, wkLen)
                Loop
            End If
        End If
    End If
    
    F_String_ReturnDelete = wkRtn
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 指定文字列以前、以降削除
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 引数情報初期化
'------------------------------------------------------------------------------
Public Property Get G_String_InitArgDelInf() As T_STRING_ARG_DEL_INF
    Dim wkInf As T_STRING_ARG_DEL_INF
    
    With G_String_InitArgDelInf
        .SrchInf = G_String_InitArgSrchInf
        
        .DelPosSpec = E_STRING_SPEC_POS_START
        .AddDelFlg = False
    End With
    
    G_String_InitArgDelInf = wkInf
End Property

'------------------------------------------------------------------------------
' 指定文字列以降、以前削除（引数情報指定）
'------------------------------------------------------------------------------
Public Function F_String_ReturnDeleteStr_Inf( _
        ByRef aArgInf As T_STRING_ARG_DEL_INF) As String
    Dim wkRtn As String
    
    Dim wkArgInf As T_STRING_ARG_DEL_INF: wkArgInf = aArgInf
    Dim wkInfAryAry As Variant, wkInfAry As Variant
    Dim wkDelStt As Long, wkDelEnd As Long
    Dim wkStrLen As Long
    
    '文字列位置取得
    With wkArgInf
        wkRtn = .SrchInf.Target
        
        '検索文字列作成
        If F_String_GetSearchInfArray_Inf(wkInfAryAry, .SrchInf) <> True Then
            '検索文字が見つからなかった場合は無視
        Else
            wkStrLen = Len(.SrchInf.Target)
            
            '開始～文字列位置まで削除
            If M_Common.F_CheckBitOn(.DelPosSpec, E_STRING_SPEC_POS_START) = True Then
                '削除開始位置を設定
                wkDelStt = .SrchInf.SttPos
                
                '削除終了位置を設定
                wkDelEnd = wkInfAryAry(LBound(wkInfAryAry))(E_STRING_IDX_SRCH_INF_POS_START)
                '削除文字列追加指定ありの場合
                If .AddDelFlg = True Then
                    wkDelEnd = wkDelEnd - 1
                '削除文字列追加指定なしの場合
                Else
                    wkDelEnd = wkDelEnd + wkInfAryAry(LBound(wkInfAryAry))(E_STRING_IDX_SRCH_INF_LENGTH) - 1
                End If
            '文字列位置～終了まで削除
            Else
                '削除開始位置を設定
                wkDelStt = wkInfAryAry(LBound(wkInfAryAry))(E_STRING_IDX_SRCH_INF_POS_START)
                '削除文字列追加指定ありの場合
                If .AddDelFlg = True And wkDelStt > D_POS_START Then
                    wkDelStt = wkDelStt + wkInfAryAry(LBound(wkInfAryAry))(E_STRING_IDX_SRCH_INF_LENGTH)
                End If
                
                '削除終了位置を設定
                wkDelEnd = .SrchInf.EndPos
                If wkDelEnd < D_POS_START Then
                    PF_String_GetPosEndAdjust wkDelEnd, wkStrLen, .SrchInf.SttPos, .SrchInf.Length
                End If
            End If
            
            '削除位置に問題ない場合、削除実施
            If wkDelStt <= wkDelEnd Then
                wkRtn = ""
                If wkDelStt > 1 Then
                    wkRtn = Left(.SrchInf.Target, wkDelStt - 1)
                End If
                If wkDelEnd < wkStrLen Then
                    wkRtn = wkRtn & Right(.SrchInf.Target, (wkStrLen - wkDelEnd))
                End If
            End If
        End If
    End With
    
    F_String_ReturnDeleteStr_Inf = wkRtn
End Function

'------------------------------------------------------------------------------
' 指定文字列以降、以前削除（引数指定）
'------------------------------------------------------------------------------
Public Function F_String_ReturnDeleteStr( _
        ByVal aTarget As String, _
        ByVal aDelete As String, _
        Optional ByVal aDelPosSpec As E_STRING_SPEC = E_STRING_SPEC_POS_START, _
        Optional ByVal aDelIdx As Long = D_IDX_START) As String
    Dim wkArgInf As T_STRING_ARG_DEL_INF: wkArgInf = G_String_InitArgDelInf
    
    With wkArgInf
        .SrchInf.Target = aTarget
        .SrchInf.Search = aDelete
        .SrchInf.GetIdx = aDelIdx
        
        .DelPosSpec = aDelPosSpec
    End With
    
    F_String_ReturnDeleteStr = F_String_ReturnDeleteStr_Inf(wkArgInf)
End Function

'==============================================================================
' 内部処理
'==============================================================================
' パターン作成
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 検索パターン作成
'------------------------------------------------------------------------------
Private Function PF_String_ReturnSearchPattern( _
        ByVal aSearch As String, ByVal aSrchSpec As E_STRING_SPEC) As String
    Dim wkRtn As String
    
    'エスケープシーケンス変換（必要分のみ）
    wkRtn = PF_String_ReturnChangeEscSeq(aSearch, Array("\", "{", "}", "(", ")", "[", "]"))
    
    '開始側パターン設定
    If M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_POS_START) = True Then
        wkRtn = F_String_ReturnAdd(wkRtn, "^", aAddSpec:=E_STRING_SPEC_POS_START)
    End If
            
    '終了側パターン設定
    If M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_POS_END) = True Then
        wkRtn = F_String_ReturnAdd(wkRtn, "$", aExcluded:="\$")
    End If
    
    PF_String_ReturnSearchPattern = wkRtn
End Function

' エスケープシーケンス置換
Private Function PF_String_ReturnChangeEscSeq( _
        ByVal aSearch As String, _
        ByVal aEscSeqAry As Variant) As String
    Dim wkRtn As String
    Dim wkTmp As Variant
    
    '初期化
    wkRtn = aSearch
    
    'エスケープシーケンスに置き換え
    For Each wkTmp In aEscSeqAry
        wkRtn = Replace(wkRtn, wkTmp, "\" & wkTmp)
    Next wkTmp
    
    '初期化
    PF_String_ReturnChangeEscSeq = wkRtn
End Function

'------------------------------------------------------------------------------
' チェックパターン作成
'------------------------------------------------------------------------------
Private Function PF_String_ReturnCheckPattern( _
        ByVal aSrchPtn As String, ByVal aSrchSpec As E_STRING_SPEC, _
        ByVal aChkPtn As String, ByVal aChkSpec As E_STRING_SPEC) As String
    Dim wkRtn As String
    Dim wkPattern As String
    
    wkRtn = "(" & aSrchPtn & ")"
    
    '一致パターン追加ありの場合
    If aChkPtn <> "" And aChkSpec <> E_STRING_SPEC_NONE Then
        '開始側チェック指定あり、かつ開始側一致指定なしの場合、検索文字列開始側にチェックパターン追加
        If M_Common.F_CheckBitOn(aChkSpec, E_STRING_SPEC_MATCH_START) = True And _
                M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_MATCH_START) <> True Then
            wkRtn = aChkPtn & wkRtn
        End If
        '終端側チェック指定あり、かつ終端側一致指定なしの場合、検索文字列終端側にチェックパターン追加
        If M_Common.F_CheckBitOn(aChkSpec, E_STRING_SPEC_MATCH_END) = True And _
                M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_MATCH_END) <> True Then
            wkRtn = wkRtn & aChkPtn
        End If
    End If
    
    '開始側パターン設定
    If M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_POS_START) = True Then
        wkRtn = F_String_ReturnAdd(wkRtn, "^", aAddSpec:=E_STRING_SPEC_POS_START)
    End If
            
    '終了側パターン設定
    If M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_POS_END) = True Then
        wkRtn = F_String_ReturnAdd(wkRtn, "$", aExcluded:="\$")
    End If
    
    PF_String_ReturnCheckPattern = wkRtn
End Function

'------------------------------------------------------------------------------
' 単語検索パターン作成
'------------------------------------------------------------------------------
Private Function PF_String_ReturnCheckPatternWord( _
        ByVal aSrchPtn As String, ByVal aSrchSpec As E_STRING_SPEC, _
        ByVal aChkPtn As String, ByVal aChkSpec As E_STRING_SPEC) As String
    Dim wkRtn As String
    Dim wkTmpStr As String
    
    '中間検索指定ありの場合、中央検索指定
    If M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_MATCH_MID) = True Then
        wkRtn = PF_String_ReturnCheckPattern(aSrchPtn, aSrchSpec, aChkPtn, aChkSpec)
    End If
        
    '先頭検索指定または中間検索指定あり、かつパターン指定に終端ありの場合
    If M_Common.F_CheckBitOn(aSrchSpec, (E_STRING_SPEC_MATCH_START Or E_STRING_SPEC_MATCH_MID)) = True And _
            M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_MATCH_END) = True Then
        '開始位置一致確認を追加
        wkTmpStr = PF_String_ReturnCheckPattern(aSrchPtn, E_STRING_SPEC_MATCH_START, aChkPtn, E_STRING_SPEC_MATCH_END)
        wkRtn = F_String_ReturnAdd(wkRtn, wkTmpStr, aDlmt:="|")
    End If
        
    '終端検索指定または中間検索指定あり、かつパターン指定に先頭ありの場合
    If M_Common.F_CheckBitOn(aSrchSpec, (E_STRING_SPEC_MATCH_END Or E_STRING_SPEC_MATCH_MID)) = True And _
            M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_MATCH_START) = True Then
        '終端位置一致確認を追加
        wkTmpStr = PF_String_ReturnCheckPattern(aSrchPtn, E_STRING_SPEC_MATCH_END, aChkPtn, E_STRING_SPEC_MATCH_START)
        wkRtn = F_String_ReturnAdd(wkRtn, wkTmpStr, aDlmt:="|")
    End If
    
    If wkRtn = "" Then
        wkRtn = aSrchPtn
    End If
    
    PF_String_ReturnCheckPatternWord = wkRtn
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 文字列調整
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 終了位置調整
'------------------------------------------------------------------------------
Private Function PF_String_GetPosEndAdjust( _
        ByRef aRtn As Long, _
        ByVal aLenMax As Long, _
        ByVal aSttPos As Long, _
        ByVal aLength As Long) As Boolean
    Dim wkRtn As Long: wkRtn = aRtn
    
    '終了位置が開始位置より小さい場合
    If wkRtn < aSttPos Then
        '長さ指定ありなら反映
        If aLength >= D_POS_START Then
            wkRtn = aSttPos + aLength - 1
        '長さ指定なしなら最大長反映
        Else
            wkRtn = aLenMax
        End If
    End If
    If wkRtn > aLenMax Then
        wkRtn = aLenMax
    End If
    
    If wkRtn >= D_POS_START Then
        aRtn = wkRtn
        PF_String_GetPosEndAdjust = True
    End If
End Function

'------------------------------------------------------------------------------
' 文字列長調整
'------------------------------------------------------------------------------
Private Function PF_String_GetLengthAdjust( _
        ByRef aRtn As Long, _
        ByVal aLenMax As Long, _
        ByVal aSttPos As Long, _
        ByVal aEndPos As Long) As Boolean
    Dim wkRtn As Long: wkRtn = aRtn
    
    Dim wkLenMax As Long: wkLenMax = aLenMax - aSttPos + 1
    
    '文字列長さが範囲外の場合
    If wkRtn < D_POS_START Then
        '終了指定ありなら反映
        If aEndPos >= aSttPos Then
            wkRtn = aEndPos - aSttPos + 1
        '終了指定なしなら最大長反映
        Else
            wkRtn = wkLenMax
        End If
    End If
    If wkRtn > wkLenMax Then
        wkRtn = wkLenMax
    End If
    
    If wkRtn >= D_POS_START Then
        aRtn = wkRtn
        PF_String_GetLengthAdjust = True
    End If
End Function
