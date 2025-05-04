Attribute VB_Name = "M_Excel"
Option Explicit
'##############################################################################
' Excel処理
'##############################################################################
' 参照設定          |   ―
'------------------------------------------------------------------------------
' 共通バージョン    |   250504
'------------------------------------------------------------------------------
' 個別バージョン    |   ―
'------------------------------------------------------------------------------
' 取込履歴          |   ファイル一覧作成_250504
'------------------------------------------------------------------------------

'==============================================================================
' 公開定義
'==============================================================================
' 定数定義
'------------------------------------------------------------------------------
Public Const D_EXCEL_ROW_START As Long = D_POS_START
Public Const D_EXCEL_ROW_NONE As Long = D_POS_NONE
Public Const D_EXCEL_ROW_NOW As Long = D_POS_NOW
Public Const D_EXCEL_ROW_END As Long = D_POS_END

Public Const D_EXCEL_CLM_START As Integer = D_POS_START
Public Const D_EXCEL_CLM_NONE As Integer = D_POS_NONE
Public Const D_EXCEL_CLM_NOW As Long = D_POS_NOW
Public Const D_EXCEL_CLM_END As Integer = D_POS_END

'------------------------------------------------------------------------------
' 構造体定義
'------------------------------------------------------------------------------
Public Type T_EXCEL_POS_ROW_INF
    Stt As Long
    End As Long
    Cnt As Long
End Type

Public Type T_EXCEL_POS_CLM_INF
    Stt As Integer
    End As Integer
    Cnt As Long
End Type

Public Type T_EXCEL_POS_INF
    Row As T_EXCEL_POS_ROW_INF
    Clm As T_EXCEL_POS_CLM_INF
End Type

'==============================================================================
' 内部定義
'==============================================================================

'==============================================================================
' 公開処理
'==============================================================================
' ワークシート関連処理
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ワークシート名→ワークシート変換
'------------------------------------------------------------------------------
Public Function F_Excel_GetSheetName2Sheet( _
        ByRef aRtn As Worksheet, _
        ByVal aName As String, _
        Optional ByVal aBk As Workbook = Nothing) As Boolean
    Dim wkRet As Boolean
    Dim wkRtn As Worksheet
    
    Dim wkBk As Workbook: Set wkBk = aBk
    
    If aName = "" Then
        Exit Function
    End If
    
    If wkBk Is Nothing Then
        Set wkBk = ThisWorkbook
    End If
    
    For Each wkRtn In wkBk.Worksheets
        If wkRtn.Name Like aName Then
            wkRet = True
            Exit For
        End If
    Next wkRtn
    
    If wkRet = True Then
        Set aRtn = wkRtn
        F_Excel_GetSheetName2Sheet = True
    End If
End Function
Public Function F_Excel_ReturnSheetName2Sheet( _
        ByVal aName As String, _
        Optional ByVal aBk As Workbook = Nothing) As Worksheet
    F_Excel_GetSheetName2Sheet F_Excel_ReturnSheetName2Sheet, aName, aBk:=aBk
End Function

'------------------------------------------------------------------------------
' 全表示
'------------------------------------------------------------------------------
Public Sub S_Excel_ShowAll( _
        ByVal aSh As Worksheet)
    If aSh Is Nothing Then
        Exit Sub
    End If
    
    With aSh
        .Cells.Rows.Hidden = False
        .Cells.Columns.Hidden = False
        
        'オートフィルタが設定されている場合は全表示
        If .AutoFilterMode <> True Then
            'フィルタ設定無しの場合は無視
        ElseIf .FilterMode <> True Then
            '絞り込みされていない場合は虫
        Else
            .ShowAllData
        End If
    End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' セル関連処理
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 位置情報初期化
'------------------------------------------------------------------------------
' 行位置初期化
Public Property Get G_Excel_InitPosRowInf( _
        Optional ByVal aRg As Range = Nothing) As T_EXCEL_POS_ROW_INF
    Dim wkRtn As T_EXCEL_POS_ROW_INF
    
    If aRg Is Nothing Then
        wkRtn.Stt = D_EXCEL_ROW_NONE
        wkRtn.End = D_EXCEL_ROW_NONE
        wkRtn.Cnt = 0
    Else
        With aRg
            wkRtn.Stt = .Row
            wkRtn.Cnt = .Rows.Count
            wkRtn.End = .Row + wkRtn.Cnt - 1
        End With
    End If
    
    G_Excel_InitPosRowInf = wkRtn
End Property

' 列位置初期化
Public Property Get G_Excel_InitPosClmInf( _
        Optional ByVal aRg As Range = Nothing) As T_EXCEL_POS_CLM_INF
    Dim wkRtn As T_EXCEL_POS_CLM_INF
    
    If aRg Is Nothing Then
        wkRtn.Stt = D_EXCEL_CLM_NONE
        wkRtn.End = D_EXCEL_CLM_NONE
        wkRtn.Cnt = 0
    Else
        With aRg
            wkRtn.Stt = .Column
            wkRtn.Cnt = .Columns.Count
            wkRtn.End = .Column + wkRtn.Cnt - 1
        End With
    End If
    
    G_Excel_InitPosClmInf = wkRtn
End Property

Public Property Get G_Excel_InitPosInf( _
        Optional ByVal aRg As Range = Nothing) As T_EXCEL_POS_INF
    Dim wkRtn As T_EXCEL_POS_INF
    
    With wkRtn
        .Row = G_Excel_InitPosRowInf(aRg)
        .Clm = G_Excel_InitPosClmInf(aRg)
    End With
    
    G_Excel_InitPosInf = wkRtn
End Property

'------------------------------------------------------------------------------
' セル範囲値取得
'------------------------------------------------------------------------------
Public Function F_Excel_ReturnRangeValueArray( _
        ByVal aRg As Range) As Variant
    Dim wkRtnAry As Variant
    
    If aRg Is Nothing Then
        Exit Function
    End If
    
    wkRtnAry = aRg.Value
    '配列でない場合
    If IsArray(wkRtnAry) <> True Then
        ReDim wkRtnAry(D_EXCEL_ROW_START To D_EXCEL_CLM_START)
        wkRtnAry(D_EXCEL_ROW_START, D_EXCEL_CLM_START) = aRg.Value
    End If
    
    F_Excel_ReturnRangeValueArray = wkRtnAry
End Function

'------------------------------------------------------------------------------
' オートフィルタ設定
'------------------------------------------------------------------------------
Public Sub S_Excel_SetAutoFilter( _
        Optional ByVal aRg As Range = Nothing, _
        Optional ByVal aSh As Worksheet = Nothing)
    Dim wkSh As Worksheet: Set wkSh = aSh
    Dim wkRg As Range: Set wkRg = aRg
    
    If Not wkRg Is Nothing Then
        Set wkSh = wkRg.Worksheet
    Else
        If wkSh Is Nothing Then
            Set wkSh = ActiveSheet
        End If
        wkRg = wkSh.UsedRange
    End If
    
    'オートフィルタが設定されている場合は一旦解除
    If wkSh.AutoFilterMode = True Then
        wkSh.Cells.AutoFilter
    End If
    
    'オートフィルタ設定
    wkRg.AutoFilter
End Sub

'==============================================================================
' 内部処理
'==============================================================================
