VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
'##############################################################################
' ワークブック
'##############################################################################
' 共通バージョン    |   250504
'------------------------------------------------------------------------------

'==============================================================================
' 内部処理
'==============================================================================
' ワークブック関連
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 保存前処理
'------------------------------------------------------------------------------
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    '初期化
    M_Config.S_Config_Init
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ワークシート関連
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ダブルクリック時処理
'------------------------------------------------------------------------------
Private Sub Workbook_SheetBeforeDoubleClick(ByVal Sh As Object, ByVal Target As Range, Cancel As Boolean)
    Dim wkClsInf(D_IDX_START To E_CONFIG_IDX_CLASS_INF_EEND) As Variant
    Dim wkCls As Object
    
    Set wkClsInf(E_CONFIG_IDX_CLASS_INF_RANGE) = Target
    
    If M_Config.F_Config_GetClass(wkCls, wkClsInf) <> True Then
        Exit Sub
    End If
    
    'イベントフラグON
    M_Config.L_Config_EventFlg = True
    'クラス毎処理実施
    Cancel = wkCls.F_GetSheetBeforeDoubleClickResult()
    'イベントフラグOFF
    M_Config.L_Config_EventFlg = False
End Sub
