Attribute VB_Name = "M_List"
Option Explicit
'##############################################################################
' リスト処理
'##############################################################################
' 参照設定          |   Microsoft Scripting Runtime
'------------------------------------------------------------------------------
' 参照モジュール    |   ―
'------------------------------------------------------------------------------
' 共通バージョン    |   250427
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
' リスト追加
'------------------------------------------------------------------------------
Public Function F_List_GetAdd( _
        ByRef aRtn As Dictionary, _
        ByVal aKey As Variant, _
        ByVal aItem As Variant, _
        Optional ByVal aUpdtFlg As Boolean = True) As Boolean
    Dim wkRtn As Dictionary: Set wkRtn = aRtn
    
    '引数チェック
    Select Case VarType(aKey)
        Case vbEmpty, vbNull
            '上記以外にもあるがそもそも指定しないので問題なし
            Exit Function
        Case Else
            '問題なし
    End Select
    
    'リスト生成
    If wkRtn Is Nothing Then
        Set wkRtn = New Dictionary
    End If
    
    With wkRtn
        'キーが存在しない場合
        If .Exists(aKey) <> True Then
            .Add aKey, aItem
        'キーが存在する場合
        Else
            '更新許可ならば上書き
            If aUpdtFlg = True Then
                M_Common.S_SetItem .Item(aKey), aItem
            Else
                Exit Function
            End If
        End If
    End With
    
    Set aRtn = wkRtn
    F_List_GetAdd = True
End Function
