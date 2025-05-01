Attribute VB_Name = "M_Common"
Option Explicit
'##############################################################################
' 共通処理
'##############################################################################
' 参照設定          |   ―
'------------------------------------------------------------------------------

'==============================================================================
' 公開定義
'==============================================================================
' 定数定義
'------------------------------------------------------------------------------
Public Const D_POS_START As Long = 1
Public Const D_POS_NONE As Long = D_POS_START - 1
Public Const D_POS_NOW As Long = D_POS_NONE - 1
Public Const D_POS_END As Long = D_POS_NOW - 1

Public Const D_ROW_START As Long = D_POS_START
Public Const D_ROW_NONE As Long = D_POS_NONE
Public Const D_ROW_NOW As Long = D_POS_NOW
Public Const D_ROW_END As Long = D_POS_END

Public Const D_CLM_START As Integer = D_POS_START
Public Const D_CLM_NONE As Integer = D_POS_NONE
Public Const D_CLM_NOW As Long = D_POS_NOW
Public Const D_CLM_END As Integer = D_POS_END

Public Const D_IDX_START As Long = 1
Public Const D_IDX_NONE As Long = D_IDX_START - 1
Public Const D_IDX_NOW As Long = D_IDX_NONE - 1
Public Const D_IDX_END As Long = D_IDX_NOW - 1
Public Const D_IDX_ALL As Long = D_IDX_END - 1

Public Enum E_RET
    E_RET_NG = 0
    E_RET_OK
    E_RET_OK_1
    E_RET_OK_2
End Enum

Public Enum E_CHECK
    E_CHECK_NONE = 0
    E_CHECK_OR
    E_CHECK_MATCH
End Enum

Public Const D_LEN_FULLPATH_MAX As Long = 256

'==============================================================================
' 内部定義
'==============================================================================

'==============================================================================
' 公開処理
'==============================================================================
' 変数処理
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' アイテム設定
'------------------------------------------------------------------------------
Public Sub S_SetItem( _
        ByRef aRtn As Variant, _
        ByVal aItem As Variant)
    If IsObject(aItem) <> True Then
        aRtn = aItem
    Else
        Set aRtn = aItem
    End If
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 配列処理
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 配列追加
'------------------------------------------------------------------------------
Public Function F_GetArrayAdd( _
        ByRef aRtnAry As Variant, _
        ByVal aAdd As Variant, _
        Optional ByVal aIdx As Long = D_IDX_END) As Boolean
    Dim wkRtnAry As Variant: wkRtnAry = aRtnAry
    
    Dim wkIdx As Long: wkIdx = aIdx
    Dim wkIStt As Long, wkIEnd As Long, wkICnt As Long
    
    '配列がある場合
    If IsArray(wkRtnAry) = True Then
        '現在の配列から調整
        wkIStt = LBound(wkRtnAry)
        wkIEnd = UBound(wkRtnAry)
        If wkIdx < D_IDX_START Then
            wkIdx = wkIEnd + 1
        End If
        
        '終了位置チェック
        If wkIdx > wkIEnd Then
            ReDim Preserve wkRtnAry(wkIStt To wkIdx)
        '開始位置チェック
        ElseIf wkIdx < wkIStt Then
            ReDim wkRtnAry(wkIdx To wkIEnd)
            
            '配列再設定
            For wkICnt = wkIStt To wkIEnd
                S_SetItem wkRtnAry(wkICnt), aRtnAry(wkICnt)
            Next wkICnt
        End If
    '配列がない場合は新規作成
    Else
        '開始位置調整
        If wkIdx < D_IDX_START Then
            wkIStt = D_IDX_START
            wkIdx = D_IDX_START
        Else
            wkIStt = wkIdx
        End If
        
        '値ありなら開始位置再調整
        If IsEmpty(wkRtnAry) <> True Then
            wkIStt = wkIStt - 1
            If wkIStt < D_IDX_START Then
                wkIStt = D_IDX_START
            End If
        End If
        
        '配列を生成
        ReDim wkRtnAry(wkIStt To wkIdx)
        
        '元の値を先頭に設定
        S_SetItem wkRtnAry(wkIStt), aRtnAry
    End If
    
    S_SetItem wkRtnAry(wkIdx), aAdd
    
    aRtnAry = wkRtnAry
    F_GetArrayAdd = True
End Function
Public Function F_ReturnArrayAdd( _
        ByVal aArray As Variant, _
        ByVal aAdd As Variant, _
        Optional ByVal aIdx As Long = D_IDX_END) As Variant
    Dim wkRtnAry As Variant: wkRtnAry = aArray
    
    F_GetArrayAdd wkRtnAry, aAdd, aIdx:=aIdx
    
    F_ReturnArrayAdd = wkRtnAry
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 演算処理
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 数値チェック
'------------------------------------------------------------------------------
Public Function F_CheckNumeric( _
        ByVal aCheck As Variant) As E_RET
    Dim wkRtn As E_RET: wkRtn = E_RET_NG
    
    If IsNull(aCheck) <> True Then
        Select Case VarType(aCheck)
            Case vbDecimal, vbByte, vbInteger, vbLong, vbLongLong
                wkRtn = E_RET_OK
            
            Case vbSingle, vbDouble, vbDate
                wkRtn = E_RET_OK_1
            
            Case vbString
                If IsNumeric(aCheck) = True Then
                    wkRtn = E_RET_OK_2
                End If
        End Select
    End If
    
    F_CheckNumeric = wkRtn
End Function

'------------------------------------------------------------------------------
' ビットチェック
'------------------------------------------------------------------------------
Public Function F_CheckBitOn( _
        ByVal aNum As Long, ByVal aCheck As Long, _
        Optional ByVal aMask As Long = &H0, _
        Optional ByVal aChkSpec As E_CHECK = E_CHECK_OR) As Boolean
    Dim wkRet As Long
    Dim wkMask As Long
    
    If aChkSpec = E_CHECK_OR Then
        wkRet = ((aNum And aCheck) > 0)
    Else
        wkMask = aMask
        If wkMask = &H0 Then
            wkMask = aCheck
        End If
        
        wkRet = ((aNum And wkMask) = aCheck)
    End If
    
    F_CheckBitOn = wkRet
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' コレクション処理
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' コレクション追加
'------------------------------------------------------------------------------
Public Function F_GetCollectionAdd( _
        ByRef aRtn As Collection, _
        ByVal aItem As Variant, _
        Optional ByVal aKey As Variant = Empty) As Boolean
    If aRtn Is Nothing Then
        Set aRtn = New Collection
    End If
    
    On Error GoTo PROC_ERROR
    If IsNull(aKey) <> True Then
        aRtn.Add aItem, Key:=aKey
    Else
        aRtn.Add aItem
    End If
    On Error GoTo 0
    
    F_GetCollectionAdd = True
PROC_ERROR:
End Function

'------------------------------------------------------------------------------
' コレクション取得
'------------------------------------------------------------------------------
Public Function F_GetCollectionItem( _
        ByRef aRtn As Variant, _
        ByVal aList As Collection, _
        ByVal aKey As Variant) As Boolean
    If IsNull(aKey) = True Or _
            aList Is Nothing Then
        Exit Function
    ElseIf aList.Count <= 0 Then
        Exit Function
    End If
    
    On Error GoTo PROC_ERROR
    S_SetItem aRtn, aList.Item(aKey)
    On Error GoTo 0
    
    F_GetCollectionItem = True
PROC_ERROR:
End Function
