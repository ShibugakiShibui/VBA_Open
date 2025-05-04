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
' 共通バージョン    |   ―
'------------------------------------------------------------------------------
' 個別バージョン    |   ファイル一覧作成_250504
'------------------------------------------------------------------------------
' 取込履歴          |   ―
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
    PE_IDX_CLM_FOLDER
    PE_IDX_CLM_PATH
    PE_IDX_CLM_FILE
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

Public Function F_GetSheetBeforeDoubleClickResult() As Boolean
    Dim wkRet As Boolean
    Dim wkCIdx As Long: wkCIdx = pgCmnCls.F_ReturnColumnIndex()
    Dim wkSubIdx As Long
    
    '列インデックスがフォルダ関連の場合
    Select Case wkCIdx
        Case PE_IDX_CLM_FOLDER
            If pgCmnCls.F_ReturnRow() = PE_ROW_TITLE Then
                wkRet = PF_InputFileList()
            End If
        
        Case PE_IDX_CLM_PATH, PE_IDX_CLM_FILE
            If pgCmnCls.F_ReturnRow() >= PE_ROW_DATA Then
                wkRet = PF_OpenFile()
            End If
        
        Case Else
            wkRet = False
    End Select
    
    F_GetSheetBeforeDoubleClickResult = wkRet
End Function

Public Function F_GetSheetChangeResult() As Boolean
    Dim wkRet As Boolean
    
    F_GetSheetChangeResult = wkRet
End Function

Public Function F_GetSheetSelectionChangeResult() As Boolean
    Dim wkRet As Boolean
    
    F_GetSheetSelectionChangeResult = wkRet
End Function

'==============================================================================
' 内部処理
'==============================================================================
Private Sub PS_AddTitleInf()
    Dim wkTitle As String
    Dim wkCIdx As PE_IDX_CLM
    
    pgCmnCls.S_AddTitleInf "フォルダ", PE_IDX_CLM_FOLDER
    pgCmnCls.S_AddTitleInf "パス", PE_IDX_CLM_PATH
    pgCmnCls.S_AddTitleInf "ファイル", PE_IDX_CLM_FILE
End Sub

'******************************************************************************
' 個別処理
'******************************************************************************
' 公開処理
'==============================================================================

'==============================================================================
' 内部処理
'==============================================================================
' ファイルリスト入力
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Function PF_InputFileList() As Boolean
    Dim wkList As Dictionary
    Dim wkListInfAry As Variant, wkListInf As Variant
    
    Dim wkGetFld As String
    Dim wkRltFld As Variant, wkNowRltFldAry As Variant, wkOldRltFldAry As Variant
    Dim wkFIStt As Long, wkFIMax As Long, wkFICnt As Long
    Dim wkOFIMax As Long
    
    Dim wkRow As Long
    
    Dim wkCIMax As Integer
    Dim wkCOfs As Integer
    
    'フォルダパス取得
    If M_File.F_File_GetDialogSelect(wkGetFld, msoFileDialogFolderPicker, aCrntFld:=pgCmnCls.F_ReturnCellValue(aRow:=PE_ROW_DATA, aCIdx:=PE_IDX_CLM_FOLDER)) <> True Then
        Exit Function
    'フォルダファイルリスト取得
    ElseIf M_File.F_File_GetFolderFileInfList(wkList, wkGetFld) <> True Then
        Exit Function
    End If
    
    If MsgBox("現在のリストを全削除し、再入力します。よろしければ「はい(Y)」を押してください", vbYesNo) <> vbYes Then
        Exit Function
    End If
    
    'データセル範囲範囲
    pgCmnCls.S_DeleteDataUsedRange
    
    wkRow = PE_ROW_DATA
    
    For Each wkRltFld In wkList
        ' フォルダ入力
        '----------------------------------------------------------------------
        '相対フォルダ入力の場合
        If wkRltFld <> wkGetFld Then
            wkCOfs = 1
            
            'フォルダで分割
            M_String.F_String_GetSplit wkNowRltFldAry, wkRltFld, "\"
            wkFIMax = UBound(wkNowRltFldAry)
            
            '旧フォルダがある場合は列開始位置を検索
            If IsArray(wkOldRltFldAry) = True Then
                wkOFIMax = UBound(wkOldRltFldAry)
                If wkOFIMax > wkFIMax Then
                    wkOFIMax = wkFIMax
                End If
                
                For wkFIStt = LBound(wkNowRltFldAry) To wkOFIMax
                    If wkOldRltFldAry(wkFIStt) <> wkNowRltFldAry(wkFIStt) Then
                        Exit For
                    End If
                    
                    wkCOfs = wkCOfs + 1
                Next wkFIStt
            End If
            
            'フォルダを入力
            For wkFICnt = wkFIStt To wkFIMax
                pgCmnCls.S_SetInputValue wkNowRltFldAry(wkFICnt), aRow:=wkRow, aCIdx:=PE_IDX_CLM_FOLDER, aCOfs:=wkCOfs
                pgCmnCls.S_SetInputValue M_String.F_String_ReturnAdd(wkGetFld, wkRltFld, aDlmt:="\"), aRow:=wkRow, aCIdx:=PE_IDX_CLM_PATH
                pgCmnCls.S_SetInputValue M_Common.D_MARK_INVALID, aRow:=wkRow, aCIdx:=PE_IDX_CLM_FILE
                wkRow = wkRow + 1
                
                wkCOfs = wkCOfs + 1
            Next wkFICnt
            
            '旧情報退避
            wkOldRltFldAry = wkNowRltFldAry
        '取得フォルダ入力の場合
        Else
            pgCmnCls.S_SetInputValue wkRltFld, aRow:=wkRow, aCIdx:=PE_IDX_CLM_FOLDER
            pgCmnCls.S_SetInputValue wkRltFld, aRow:=wkRow, aCIdx:=PE_IDX_CLM_PATH
            pgCmnCls.S_SetInputValue M_Common.D_MARK_INVALID, aRow:=wkRow, aCIdx:=PE_IDX_CLM_FILE
            wkRow = wkRow + 1
        End If
        
        ' ファイル入力
        '----------------------------------------------------------------------
        wkListInfAry = wkList.Item(wkRltFld)
        If IsArray(wkListInfAry) = True Then
            For Each wkListInf In wkListInfAry
                pgCmnCls.S_SetInputValue wkListInf(E_FILE_IDX_LIST_INF_FULLPATH), aRow:=wkRow, aCIdx:=PE_IDX_CLM_PATH
                pgCmnCls.S_SetInputValue wkListInf(E_FILE_IDX_LIST_INF_NAME), aRow:=wkRow, aCIdx:=PE_IDX_CLM_FILE
                wkRow = wkRow + 1
            Next wkListInf
            wkRow = wkRow + 1
        End If
    Next wkRltFld
    
    '入力更新
    pgCmnCls.S_UpdateInput aInfUpdtFlg:=True
    pgCmnCls.S_SetAutoFilter
    PF_InputFileList = True
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ファイルオープン
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Function PF_OpenFile() As Boolean
    Dim wkExePath As String: wkExePath = "C:\Program Files (x86)\Hidemaru\Hidemaru.exe"
    Dim wkFilePath As String: wkFilePath = pgCmnCls.F_ReturnCellValue(aCIdx:=PE_IDX_CLM_PATH)
    
    PF_OpenFile = M_File.F_File_Open(wkFilePath, wkExePath)
End Function
