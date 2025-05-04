Attribute VB_Name = "M_Config"
Option Explicit
'##############################################################################
' コンフィグ設定
'##############################################################################
' 共通バージョン    |   250504
'------------------------------------------------------------------------------

'==============================================================================
' 公開定義
'==============================================================================
' 定数定義
'------------------------------------------------------------------------------
'クラス情報
Public Enum E_CONFIG_IDX_CLASS_INF
    E_CONFIG_IDX_CLASS_INF_NONE = D_IDX_START - 1
    E_CONFIG_IDX_CLASS_INF_FULLPATH
    E_CONFIG_IDX_CLASS_INF_BOOK
    E_CONFIG_IDX_CLASS_INF_BOOK_NAME
    E_CONFIG_IDX_CLASS_INF_SHEET
    E_CONFIG_IDX_CLASS_INF_SHEET_NAME
    E_CONFIG_IDX_CLASS_INF_RANGE
    E_CONFIG_IDX_CLASS_INF_ARG_TYPE
    E_CONFIG_IDX_CLASS_INF_MAX
    E_CONFIG_IDX_CLASS_INF_EEND = E_CONFIG_IDX_CLASS_INF_MAX - 1
End Enum

'クラス指定
Public Enum E_CONFIG_IDX_CLASS
    E_CONFIG_IDX_CLASS_NONE = D_IDX_START - 1
    E_CONFIG_IDX_CLASS_UNIQUE
    E_CONFIG_IDX_CLASS_MAX
    E_CONFIG_IDX_CLASS_EEND = E_CONFIG_IDX_CLASS_MAX - 1
End Enum

' 行指定
Public Enum E_CONFIG_IDX_CLASS_ARG_TYPE
    E_CONFIG_IDX_CLASS_ARG_TYPE_NONE = D_IDX_START - 1
    E_CONFIG_IDX_CLASS_ARG_TYPE_SHEET
    E_CONFIG_IDX_CLASS_ARG_TYPE_RANGE
    E_CONFIG_IDX_CLASS_ARG_TYPE_MAX
    E_CONFIG_IDX_CLASS_ARG_TYPE_EEND = E_CONFIG_IDX_CLASS_ARG_TYPE_MAX - 1
End Enum

' 行指定
Public Enum E_CONFIG_ROW
    E_CONFIG_ROW_NONE = D_EXCEL_ROW_START - 1
    E_CONFIG_ROW_MAX
    E_CONFIG_ROW_EEND = E_CONFIG_ROW_MAX - 1
End Enum

'==============================================================================
' 内部定義
'==============================================================================
' 構造体定義
'------------------------------------------------------------------------------
Private Type PT_CONFIG_CLASS_INF
    Cls As Object
End Type

Private Type PT_INF
    EventFlg As Boolean
    
    ClsInf(D_IDX_START To E_CONFIG_IDX_CLASS_EEND) As PT_CONFIG_CLASS_INF
End Type

'------------------------------------------------------------------------------
' 変数定義
'------------------------------------------------------------------------------
Private pgInf As PT_INF

'==============================================================================
' 公開処理
'==============================================================================
' 初期化処理
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Sub S_Config_Init()
    Dim wkCnt As Long
    
    With pgInf
        .EventFlg = False
        
        For wkCnt = D_IDX_START To E_CONFIG_IDX_CLASS_EEND
            With .ClsInf(wkCnt)
                Set .Cls = Nothing
            End With
        Next wkCnt
    End With
End Sub
 
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' クラス処理
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' クラス取得
'------------------------------------------------------------------------------
Public Function F_Config_GetClass( _
        ByRef aRtn As Object, _
        ByRef aClsInf As Variant) As Boolean
    Dim wkRet As Boolean
    
    Dim wkICnt As Long
    
    Dim wkCls As Object
    
    'イベント中は取得不要
    If pgInf.EventFlg = True Then
        Exit Function
    End If
    
    'クラス情報チェック
    If PF_Config_CheckClsInf(aClsInf) <> True Then
        Exit Function
    End If
    
    For wkICnt = D_IDX_START To E_CONFIG_IDX_CLASS_EEND
        Select Case wkICnt
            Case E_CONFIG_IDX_CLASS_UNIQUE
                Set wkCls = New C_Unique
        End Select
        
        '初期化成功時はクラスを保持
        If wkCls.F_Init(aClsInf) = True Then
            With pgInf.ClsInf(wkICnt)
                Set .Cls = wkCls
                
                Set aRtn = wkCls
                wkRet = True
                Exit For
            End With
        End If
    Next wkICnt
    F_Config_GetClass = wkRet
End Function

' クラス引数情報チェック処理
Private Function PF_Config_CheckClsInf( _
        ByRef aRtn As Variant) As Boolean
    Dim wkRtn As Variant: wkRtn = aRtn
    
    Dim wkTmpSh As Worksheet
    
    'セル範囲がある場合
    If Not wkRtn(E_CONFIG_IDX_CLASS_INF_RANGE) Is Nothing Then
        Set wkTmpSh = wkRtn(E_CONFIG_IDX_CLASS_INF_RANGE).Worksheet
        
        Set wkRtn(E_CONFIG_IDX_CLASS_INF_SHEET) = wkTmpSh
        wkRtn(E_CONFIG_IDX_CLASS_INF_ARG_TYPE) = E_CONFIG_IDX_CLASS_ARG_TYPE_RANGE
    Else
        'シートがある場合
        If Not wkRtn(E_CONFIG_IDX_CLASS_INF_SHEET) Is Nothing Then
            Set wkTmpSh = wkRtn(E_CONFIG_IDX_CLASS_INF_SHEET)
            '処理続行
        'シート名がある場合
        ElseIf M_Excel.F_Excel_GetSheetName2Sheet(wkTmpSh, _
                                                    wkRtn(E_CONFIG_IDX_CLASS_INF_SHEET_NAME), _
                                                    wkRtn(E_CONFIG_IDX_CLASS_INF_BOOK)) = True Then
            '処理続行
            Set wkRtn(E_CONFIG_IDX_CLASS_INF_SHEET) = wkTmpSh
        '双方ない場合は異常終了
        Else
            Exit Function
        End If
        
        wkRtn(E_CONFIG_IDX_CLASS_INF_RANGE) = wkTmpSh.Cells(D_EXCEL_ROW_START, D_EXCEL_CLM_START)
        wkRtn(E_CONFIG_IDX_CLASS_INF_ARG_TYPE) = E_CONFIG_IDX_CLASS_ARG_TYPE_SHEET
    End If
        
    If wkRtn(E_CONFIG_IDX_CLASS_INF_ARG_TYPE) <> E_CONFIG_IDX_CLASS_ARG_TYPE_NONE Then
        With wkTmpSh
            wkRtn(E_CONFIG_IDX_CLASS_INF_SHEET_NAME) = .Name
            
            Set wkRtn(E_CONFIG_IDX_CLASS_INF_BOOK) = .Parent
            With .Parent
                wkRtn(E_CONFIG_IDX_CLASS_INF_FULLPATH) = .FullName
                wkRtn(E_CONFIG_IDX_CLASS_INF_BOOK_NAME) = .Name
            End With
        End With
    
        aRtn = wkRtn
        PF_Config_CheckClsInf = True
    End If
End Function

'------------------------------------------------------------------------------
' イベントフラグ設定
'------------------------------------------------------------------------------
Public Property Let L_Config_EventFlg( _
            ByVal aFlg As Boolean)
    pgInf.EventFlg = aFlg
End Property

'==============================================================================
' 内部処理
'==============================================================================
