VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Unique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'##############################################################################
' 個別クラス
'##############################################################################
' 共通バージョン    |   250504
'------------------------------------------------------------------------------

'******************************************************************************
' 共通定義
'******************************************************************************
' 公開定義
'==============================================================================

'==============================================================================
' 内部定義
'==============================================================================
' 定数定義
'------------------------------------------------------------------------------
Private Enum PE_IDX_CLM
    PE_IDX_CLM_NONE = D_IDX_START - 1
    PE_IDX_CLM_MAX
    PE_IDX_CLM_EEND = PE_IDX_CLM_MAX - 1
End Enum

Private Enum PE_IDX_ROW_TITLE
    PE_IDX_ROW_TITLE_NONE = D_IDX_START - 1
    PE_IDX_ROW_TITLE_TITLE
    PE_IDX_ROW_TITLE_MAX
    PE_IDX_ROW_TITLE_EEND = PE_IDX_ROW_TITLE_MAX - 1
End Enum

Private Enum PE_ROW
    PE_ROW_COMMON = E_CONFIG_ROW_EEND
    PE_ROW_TITLE
    PE_ROW_TITLE_END = PE_ROW_TITLE + PE_IDX_ROW_TITLE_EEND - 1
    PE_ROW_AUTOFIL
    PE_ROW_DATA
End Enum

Private Const PD_SHEET_CHECKNAME As String = "*"

'------------------------------------------------------------------------------
' 変数定義
'------------------------------------------------------------------------------
Private pgCmnCls As New C_Common

'******************************************************************************
' 個別定義
'******************************************************************************

'******************************************************************************
' 共通処理
'******************************************************************************
' 公開処理
'==============================================================================
' 初期化処理
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_Init( _
        ByRef aClsInf As Variant) As Boolean
    PS_AddTitleInf
    
    F_Init = pgCmnCls.F_Init(aClsInf, PD_SHEET_CHECKNAME, PE_ROW_TITLE, PE_ROW_TITLE_END, PE_ROW_AUTOFIL, PE_ROW_DATA)
End Function

'------------------------------------------------------------------------------
' セル選択時処理
'------------------------------------------------------------------------------
Public Function F_GetSheetSelectionChangeResult() As Boolean
    Dim wkRet As Boolean
    Dim wkCIdx As Long: wkCIdx = pgCmnCls.F_ReturnColumnIndex()
    Dim wkSubIdx As Long
    
    Select Case wkCIdx
        Case Else
            wkRet = False
    End Select
    
    F_GetSheetSelectionChangeResult = wkRet
End Function

'------------------------------------------------------------------------------
' ダブルクリック時処理
'------------------------------------------------------------------------------
Public Function F_GetSheetBeforeDoubleClickResult() As Boolean
    Dim wkRet As Boolean
    Dim wkCIdx As Long: wkCIdx = pgCmnCls.F_ReturnColumnIndex()
    Dim wkSubIdx As Long
    
    Select Case wkCIdx
        Case Else
            wkRet = False
    End Select
    
    F_GetSheetBeforeDoubleClickResult = wkRet
End Function

'------------------------------------------------------------------------------
' セル入力時処理
'------------------------------------------------------------------------------
Public Function F_GetSheetChangeResult() As Boolean
    Dim wkRet As Boolean
    Dim wkCIdx As Long: wkCIdx = pgCmnCls.F_ReturnColumnIndex()
    Dim wkSubIdx As Long
    
    Select Case wkCIdx
        Case Else
            wkRet = False
    End Select
    
    F_GetSheetChangeResult = wkRet
End Function

'==============================================================================
' 内部処理
'==============================================================================
Private Sub PS_AddTitleInf()
    
    pgCmnCls.S_AddTitleInf "", PE_IDX_CLM_NONE
End Sub

'******************************************************************************
' 個別処理
'******************************************************************************
' 公開処理
'==============================================================================

'==============================================================================
' 内部処理
'==============================================================================
