VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "C_Common"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'##############################################################################
' 共通クラス
'##############################################################################
' 共通バージョン    |   250504
'------------------------------------------------------------------------------

'==============================================================================
' 公開定義
'==============================================================================

'==============================================================================
' 内部定義
'==============================================================================
' 定数定義
'------------------------------------------------------------------------------
Private Const PD_RGVALARY_CELL_MAX As Long = 327670
Private Const PD_RGVALARY_CELL_QUE_MAX As Long = 3

'------------------------------------------------------------------------------
' 構造体定義
'------------------------------------------------------------------------------
Private Type PT_TITLE_INF
    TitleAry As Variant
    ExistFlg As Boolean
    PosClmInf As T_EXCEL_POS_CLM_INF
End Type

Private Type PT_RGVALARY_INF
    RgValAry As Variant
    PosRowInf As T_EXCEL_POS_ROW_INF
    GetTime As Date
End Type

Private Type PT_INPUT_INF
    PosInf As T_EXCEL_POS_INF
    RgValAry As Variant
    InsertFlg As Boolean
    CmptFlg As Boolean
End Type

Private Type PT_INF
    TitInfAry() As PT_TITLE_INF
    TitInfCnt As Long
    
    ClsInf As Variant
    PosInf  As T_EXCEL_POS_INF
    CIdx As Long
    
    TitRowInf As T_EXCEL_POS_ROW_INF
    AFilRow As Long
    DatRStt As Long
    
    UsedRg As Range
    UsedPosInf As T_EXCEL_POS_INF
    
    RgVAInfAry(D_IDX_START To PD_RGVALARY_CELL_QUE_MAX) As PT_RGVALARY_INF
    RgVARMax As Long
    
    InputInfAry() As PT_INPUT_INF
    InputInfCnt As Long
    InputInsertFlg As Boolean
    
    InitFlg As Boolean
End Type

'------------------------------------------------------------------------------
' 変数定義
'------------------------------------------------------------------------------
Private pgInf As PT_INF

'==============================================================================
' 公開処理
'==============================================================================
' 初期化処理関連
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' タイトル情報追加
'------------------------------------------------------------------------------
Public Sub S_AddTitleInf( _
        ByVal aTitAry As Variant, _
        ByVal aCIdx As Long, _
        Optional ByVal aSubCIdx As Long = D_IDX_START, _
        Optional ByVal aExistFlg As Boolean = True)
    Dim wkTitInf As PT_TITLE_INF: wkTitInf = PG_InitTitleInf()
    Dim wkTmpAry As Variant: wkTmpAry = aTitAry
    Dim wkCIdx As Long, wkCICnt As Long
    
    'タイトル情報が配列でない場合は配列化
    If IsArray(wkTmpAry) = True Then
    ElseIf IsNull(wkTmpAry) = True Then
        Exit Sub
    Else
        wkTmpAry = Array(wkTmpAry)
    End If
    
    With wkTitInf
        .TitleAry = wkTmpAry
        .ExistFlg = aExistFlg
    End With
    
    With pgInf
        wkCIdx = aCIdx + aSubCIdx - 1
        '列インデックスが範囲外の場合は対象外
        If wkCIdx < D_IDX_START Then
            Exit Sub
        '列インデックス未作成の場合は作成
        ElseIf .TitInfCnt < wkCIdx Then
            ReDim Preserve .TitInfAry(D_IDX_START To wkCIdx)
            
            '作成分を初期化
            For wkCICnt = .TitInfCnt + 1 To wkCIdx
                .TitInfAry(wkCICnt) = PG_InitTitleInf()
            Next wkCICnt
            .TitInfCnt = wkCIdx
        End If
        
        .TitInfAry(wkCIdx) = wkTitInf
    End With
End Sub

'------------------------------------------------------------------------------
' 初期化処理
'------------------------------------------------------------------------------
Public Function F_Init( _
        ByRef aClsInf As Variant, _
        ByVal aChkShNm As String, _
        ByVal aTitRStt As Long, aTitREnd As Long, aAFilRow As Long, aDatRStt As Long) As Boolean
    Dim wkInf As PT_INF: wkInf = pgInf
    Dim wkTitInf As Variant
    
    Dim wkCnt As Long
    Dim wkTmpAry As Variant
    
    With wkInf
        '初期化チェック
        If .TitInfCnt < D_IDX_START Then
            Exit Function
        End If
    
        'クラス情報チェック
        .ClsInf = aClsInf
        If .ClsInf(E_CONFIG_IDX_CLASS_INF_ARG_TYPE) < D_IDX_START Or .ClsInf(E_CONFIG_IDX_CLASS_INF_ARG_TYPE) > E_CONFIG_IDX_CLASS_ARG_TYPE_EEND Then
            Exit Function
        End If
        
        '行関係チェック
        If aTitRStt < D_EXCEL_ROW_START Then
            Exit Function
        End If
        .TitRowInf.Stt = aTitRStt
        .TitRowInf.End = aTitREnd
        .AFilRow = aAFilRow
        .DatRStt = aDatRStt
        
        If .TitRowInf.End < .TitRowInf.Stt Then
            .TitRowInf.End = .TitRowInf.Stt
        End If
        If .DatRStt < .TitRowInf.End Then
            .DatRStt = .TitRowInf.End + 1
        End If
    End With
    
    '初期化
    PS_Init_Sub wkInf
    
    With wkInf
        'タイトルセル範囲値取得
        wkTmpAry = M_Excel.F_Excel_ReturnRangeValueArray(Range(.UsedRg.Rows(.TitRowInf.Stt), .UsedRg.Rows(.TitRowInf.End)))
        
        For wkCnt = D_IDX_START To .TitInfCnt
            'タイトル情報チェックが正常でないなら終了
            If PF_CheckTitleInf(.TitInfAry(wkCnt), wkTmpAry) <> True Then
                Exit Function
            End If
            
            '現在列インデックス保持（後勝ち）
            If .PosInf.Clm.Stt >= .TitInfAry(wkCnt).PosClmInf.Stt And .PosInf.Clm.Stt <= .TitInfAry(wkCnt).PosClmInf.End Then
                .CIdx = wkCnt
            End If
        Next wkCnt
        
        .InitFlg = True
    End With
    
    pgInf = wkInf
    F_Init = pgInf.InitFlg
End Function

' サブルーチン（初期化）
Private Function PS_Init_Sub( _
        ByRef aRtn As PT_INF)
    Dim wkCnt As Long
    
    With aRtn
        .PosInf = M_Excel.G_Excel_InitPosInf(.ClsInf(E_CONFIG_IDX_CLASS_INF_RANGE))
        
        '開始位置から使用範囲を設定
        Set .UsedRg = Range(.ClsInf(E_CONFIG_IDX_CLASS_INF_SHEET).Cells(D_EXCEL_ROW_START, D_EXCEL_CLM_START), .ClsInf(E_CONFIG_IDX_CLASS_INF_SHEET).UsedRange)
        .UsedPosInf = M_Excel.G_Excel_InitPosInf(.UsedRg)
        If .UsedPosInf.Row.End < .DatRStt Then
            .UsedPosInf.Row.End = .DatRStt
            Set .UsedRg = Range(.UsedRg.Rows(D_EXCEL_ROW_START), .UsedRg.Rows(.UsedPosInf.Row.End))
        End If
        
        'セル範囲値情報初期化
        For wkCnt = D_IDX_START To PD_RGVALARY_CELL_QUE_MAX
            .RgVAInfAry(wkCnt) = PG_InitRangeValueAryInf()
        Next wkCnt
        .RgVARMax = ((PD_RGVALARY_CELL_MAX / .UsedPosInf.Clm.End) + 0.5)
        
        '入力情報初期化
        Erase .InputInfAry
        .InputInfCnt = D_IDX_START - 1
        .InputInsertFlg = False
    End With
End Function

' サブルーチン（タイトル情報チェック）
Private Function PF_CheckTitleInf( _
        ByRef aRtn As PT_TITLE_INF, _
        ByVal aTitRgValAry As Variant) As Boolean
    Dim wkRet As Boolean
    
    Dim wkTitAry As Variant: wkTitAry = aRtn.TitleAry
    Dim wkAOfs As Long
    
    Dim wkRIStt As Long, wkRIEnd As Long, wkRICnt As Long
    Dim wkRINow As Long
    Dim wkCStt As Integer, wkCMax As Integer, wkCCnt As Integer
    Dim wkCEnd As Integer
    
    Dim wkTmpFlg As Boolean
    
    '初期化
    wkRIStt = LBound(aTitRgValAry, 1)
    wkRIEnd = UBound(aTitRgValAry, 1)
    wkRINow = wkRIStt - 1
    
    wkCStt = LBound(aTitRgValAry, 2)
    wkCMax = UBound(aTitRgValAry, 2)
    
    '引数チェック
    If IsArray(wkTitAry) <> True Then
        '必須フラグONの場合は異常終了
        If aRtn.ExistFlg = True Then
            Exit Function
        End If
        '設定していない場合は無視で正常扱い
        wkRet = True
    ElseIf wkTitAry(LBound(wkTitAry)) = "" Then
        '最上位タイトルが空白の場合は異常終了
        Exit Function
    Else
        wkAOfs = LBound(wkTitAry) - wkRIStt
        If UBound(wkTitAry) > (wkRIEnd + wkAOfs) Then
            Exit Function
        End If
        '調整
        wkRIEnd = UBound(wkTitAry) - wkAOfs
        
        '開始列検索
        For wkCCnt = wkCStt To wkCMax
            'チェック済行再チェック
            For wkRICnt = wkRIStt To wkRINow
                '値の入力がある場合は行を戻す
                If aTitRgValAry(wkRICnt, wkCCnt) <> "" Then
                    wkRINow = wkRICnt - 1
                    Exit For
                End If
            Next wkRICnt
            
            For wkRICnt = wkRINow + 1 To wkRIEnd
                '不一致の場合は次列チェックへ
                If Not aTitRgValAry(wkRICnt, wkCCnt) Like wkTitAry(wkRICnt + wkAOfs) Then
                    Exit For
                End If
                
                '現行を保持
                wkRINow = wkRICnt
                If wkRINow >= wkRIEnd Then
                    wkRet = True
                    wkCStt = wkCCnt
                    Exit For
                End If
            Next wkRICnt
            
            If wkRet = True Then
                Exit For
            End If
        Next wkCCnt
        
        If wkRet = True Then
            '終了列検索
            wkTmpFlg = False
            wkCEnd = wkCStt
            For wkCCnt = wkCStt + 1 To wkCMax
                'チェック済行再チェック
                For wkRICnt = wkRIStt To wkRIEnd
                    '値の入力がある場合は終了列を設定して終了
                    If aTitRgValAry(wkRICnt, wkCCnt) = "" Then
                        '空白は対象
                    ElseIf aTitRgValAry(wkRICnt, wkCCnt) = wkTitAry(wkRICnt + wkAOfs) Then
                        '一致は対象
                    Else
                        '対象外の場合は最終行を保持して終了
                        wkTmpFlg = True
                        Exit For
                    End If
                Next wkRICnt
                
                '終了列が見つかった場合は終了
                If wkTmpFlg = True Then
                    Exit For
                End If
                wkCEnd = wkCCnt
            Next wkCCnt
            
            '見つかった場合は列位置を保持
            aRtn.PosClmInf.Stt = wkCStt
            aRtn.PosClmInf.End = wkCEnd
        Else
            '見つからなかった場合は列位置を初期化
            aRtn.PosClmInf = M_Excel.G_Excel_InitPosClmInf
        End If
    End If
    
    '見つかった場合、または存在チェック不要の場合は正常
    If wkRet = True Or aRtn.ExistFlg <> True Then
        PF_CheckTitleInf = True
    End If
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 行取得
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_ReturnRow() As Long
    F_ReturnRow = pgInf.PosInf.Row.Stt
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 列取得
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 開始位置取得
'------------------------------------------------------------------------------
Public Function F_GetColumn( _
        ByRef aRtn As Integer, _
        Optional ByVal aCIdx As Long = D_IDX_NOW, Optional ByVal aSubCIdx As Long = D_IDX_START) As Boolean
    Dim wkTitInf As PT_TITLE_INF
    
    If PF_GetColumn_Sub(wkTitInf, aCIdx, aSubCIdx) <> True Then
        Exit Function
    End If
    
    aRtn = wkTitInf.PosClmInf.Stt
    F_GetColumn = True
End Function
Public Function F_ReturnColumn( _
        Optional ByVal aCIdx As Long = D_IDX_NOW, Optional ByVal aSubCIdx As Long = D_IDX_START) As Integer
    F_GetColumn F_ReturnColumn, aCIdx, aSubCIdx
End Function

'サブルーチン
Private Function PF_GetColumn_Sub( _
        ByRef aRtn As PT_TITLE_INF, _
        ByVal aCIdx As Long, ByVal aSubCIdx As Long) As Boolean
    Dim wkRtn As PT_TITLE_INF: wkRtn = PG_InitTitleInf
    Dim wkCIdx As Long
    
    With pgInf
        '指定が無い場合は現在値を設定
        If aCIdx < D_IDX_START Then
            wkCIdx = .CIdx
        '指定がある場合はインデックス算出
        Else
            wkCIdx = aCIdx + aSubCIdx - 1
        End If
        
        If wkCIdx >= D_IDX_START And wkCIdx <= .TitInfCnt Then
            wkRtn = .TitInfAry(wkCIdx)
        End If
    End With
    
    With wkRtn.PosClmInf
        '未設定の場合は異常終了
        If .Stt < D_EXCEL_CLM_START Or .End < .Stt Then
            Exit Function
        End If
    End With
    
    aRtn = wkRtn
    PF_GetColumn_Sub = True
End Function

'------------------------------------------------------------------------------
' 終了位置取得
'------------------------------------------------------------------------------
Public Function F_GetEndColumn( _
        ByRef aRtn As Integer, _
        ByVal aCIdx As Long, Optional ByVal aSubCIdx As Long = D_IDX_START) As Boolean
    Dim wkTitInf As PT_TITLE_INF
    
    If PF_GetColumn_Sub(wkTitInf, aCIdx, aSubCIdx) <> True Then
        Exit Function
    End If
    
    aRtn = wkTitInf.PosClmInf.End
    F_GetEndColumn = True
End Function
Public Function F_ReturnEndColumn( _
        ByVal aCIdx As Long, Optional ByVal aSubCIdx As Long = D_IDX_START) As Integer
    F_GetEndColumn F_ReturnEndColumn, aCIdx, aSubCIdx
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 列インデックス取得
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_GetColumnIndex( _
        ByRef aRtn As Long, _
        Optional ByVal aClm As Integer = D_EXCEL_CLM_NOW) As Boolean
    Dim wkRet As Boolean
    Dim wkGetCIdx As Long
    
    Dim wkClm As Integer: aClm = wkClm
    Dim wkCIdx As Long
    
    With pgInf
        If wkClm < D_EXCEL_CLM_START Then
            wkClm = .PosInf.Clm.Stt
        End If
    
        For wkCIdx = D_IDX_START To .TitInfCnt
            With .TitInfAry(wkCIdx).PosClmInf
                '範囲内の情報があれば保持（後勝ち）
                If wkClm >= .Stt And wkClm <= .End Then
                    wkGetCIdx = wkCIdx
                    wkRet = True
                End If
            End With
        Next wkCIdx
    End With
    
    If wkRet = True Then
        aRtn = wkGetCIdx
    End If
    F_GetColumnIndex = wkRet
End Function
Public Function F_ReturnColumnIndex( _
        Optional ByVal aClm As Integer = D_EXCEL_CLM_NOW) As Long
    F_GetColumnIndex F_ReturnColumnIndex, aClm
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' セル範囲取得
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' データセル範囲取得
'------------------------------------------------------------------------------
Public Function F_ReturnDataUsedRange() As Range
    With pgInf
        Set F_ReturnDataUsedRange = Range(.UsedRg.Rows(.DatRStt), .UsedRg.Rows(.UsedRg.Rows.Count))
    End With
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' セル範囲クリア
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' データセル範囲クリア
'------------------------------------------------------------------------------
Public Sub S_DeleteDataUsedRange()
    Dim wkRg As Range
    
    Set wkRg = Me.F_ReturnDataUsedRange()
    
    M_Excel.S_Excel_ShowAll wkRg.Worksheet
    wkRg.Value = ""
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' セル値取得
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 列指定
Public Function F_ReturnCellValue_Clm( _
        Optional ByVal aRow As Long = D_EXCEL_ROW_NOW, _
        Optional ByVal aClm As Integer = D_EXCEL_CLM_NOW) As Variant
    Dim wkRet As Boolean
    Dim wkRtn As Variant
    Dim wkGetIdx As Long
    
    Dim wkMinTime As Date
    
    Dim wkInputInf As PT_INPUT_INF
    
    Dim wkCnt As Long
    Dim wkRow As Long: wkRow = aRow
    Dim wkRStt As Long, wkREnd As Long
    Dim wkClm As Integer: wkClm = aClm
    Dim wkRIdx As Long
    Dim wkCIdx As Integer
    
    With pgInf
        '取得位置調整
        If wkRow < D_EXCEL_ROW_START Then
            wkRow = .PosInf.Row.Stt
        End If
        If wkClm < D_EXCEL_CLM_START Then
            wkClm = .PosInf.Clm.Stt
        End If
        
        '入力情報がある場合
        If .InputInfCnt >= D_IDX_START Then
            For wkCnt = D_IDX_START To .InputInfCnt
                With .InputInfAry(wkCnt)
                    '挿入フラグありの場合
                    If .InsertFlg = True Then
                        '無視
                    '行、列が入力情報範囲内の場合は値を退避（後勝ち）
                    ElseIf wkRow <= .PosInf.Row.Stt And wkRow >= .PosInf.Row.End And _
                            wkClm <= .PosInf.Clm.Stt And wkClm >= .PosInf.Clm.End Then
                        wkRIdx = wkRow - .PosInf.Row.Stt + 1
                        wkCIdx = wkClm - .PosInf.Clm.Stt + 1
                        wkRtn = .RgValAry(wkRIdx, wkCIdx)
                        wkRet = True
                    End If
                End With
            Next wkCnt
        End If
        
        '返却値が取得できた場合
        If wkRet = True Then
        '取得位置が使用セル範囲内の場合
        ElseIf wkRow <= .UsedPosInf.Row.End And wkClm <= .UsedPosInf.Clm.End Then
            For wkCnt = D_IDX_START To PD_RGVALARY_CELL_QUE_MAX
                With pgInf.RgVAInfAry(wkCnt)
                    '行範囲内ならループ終了
                    If wkRow >= .PosRowInf.Stt And wkRow <= .PosRowInf.End Then
                        wkGetIdx = wkCnt
                        Exit For
                    End If
                    
                    If wkCnt <= D_IDX_START Or wkMinTime > .GetTime Then
                        wkMinTime = .GetTime
                        wkGetIdx = wkCnt
                    End If
                End With
            Next wkCnt
            
            '見つからなかった場合は現在位置を取得
            If wkCnt > PD_RGVALARY_CELL_QUE_MAX Then
                '取得行を設定
                wkRStt = Int(wkRow / .RgVARMax) * .RgVARMax
                wkREnd = wkRStt + .RgVARMax - 1
                If wkRStt < D_EXCEL_ROW_START Then
                    wkRStt = D_EXCEL_ROW_START
                End If
                If wkREnd > .UsedPosInf.Row.End Then
                    wkREnd = .UsedPosInf.Row.End
                End If
                
                .RgVAInfAry(wkGetIdx).RgValAry = M_Excel.F_Excel_ReturnRangeValueArray(Range(.UsedRg.Rows(wkRStt), .UsedRg.Rows(wkREnd)))
                .RgVAInfAry(wkGetIdx).PosRowInf.Stt = wkRStt
                .RgVAInfAry(wkGetIdx).PosRowInf.End = wkREnd
                .RgVAInfAry(wkGetIdx).PosRowInf.Cnt = wkREnd - wkRStt + 1
                .RgVAInfAry(wkGetIdx).GetTime = Now
            End If
            
            '取得行インデックス算出
            wkRIdx = wkRow - .RgVAInfAry(wkGetIdx).PosRowInf.Stt + 1
            
            'セル値設定
            wkRtn = .RgVAInfAry(wkGetIdx).RgValAry(wkRIdx, wkClm)
        '取得位置が使用セル範囲外の場合、直接取得
        Else
            wkRtn = .ClsInf(E_CONFIG_IDX_CLASS_INF_SHEET).Cells(wkRow, wkClm)
        End If
    End With
    
    F_ReturnCellValue_Clm = wkRtn
End Function

' 列インデックス指定
Public Function F_ReturnCellValue( _
        Optional ByVal aRow As Long = D_EXCEL_ROW_NOW, _
        Optional ByVal aCIdx As Long = D_IDX_NOW, _
        Optional ByVal aSubCIdx As Long = D_IDX_START) As Variant
    Dim wkClm As Integer
    
    '列インデックス→列変換
    If Me.F_GetColumn(wkClm, aCIdx, aSubCIdx) <> True Then
        Exit Function
    End If
    
    F_ReturnCellValue = F_ReturnCellValue_Clm(aRow, wkClm)
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 入力処理関連
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 入力情報設定
'------------------------------------------------------------------------------
' 列指定
Public Sub S_SetInputValue_Clm( _
        ByVal aRgValAry As Variant, _
        Optional ByVal aRow As Long = D_EXCEL_ROW_NOW, _
        Optional ByVal aClm As Integer = D_EXCEL_CLM_NOW, _
        Optional ByVal aInsertFlg As Boolean = False)
    Dim wkInRgValAry As Variant
    
    Dim wkRStt As Long: wkRStt = aRow
    Dim wkREnd As Long
    Dim wkCStt As Integer: wkCStt = aClm
    Dim wkCEnd As Integer
    Dim wkRIStt As Long, wkRIEnd As Long, wkRIOfs As Long, wkRICnt As Long
    Dim wkCIStt As Integer, wkCIEnd As Integer, wkCIOfs As Integer, wkCICnt As Integer
    
    Dim wkROfs As Long, wkCOfs As Integer
    
    With pgInf
        '位置調整
        If wkRStt < D_EXCEL_ROW_START Then
            wkRStt = .PosInf.Row.Stt
        End If
        If wkCEnd < D_EXCEL_CLM_NOW Then
            wkCEnd = .PosInf.Clm.Stt
        End If
        
        .InputInfCnt = .InputInfCnt + 1
        ReDim Preserve .InputInfAry(D_IDX_START To .InputInfCnt)
        
        '入力情報に設定
        With .InputInfAry(.InputInfCnt)
            With .PosInf
                .Row.Stt = wkRStt
                .Clm.Stt = wkCStt
            End With
            .InsertFlg = aInsertFlg
        
            If IsArray(aRgValAry) <> True Then
                ReDim wkInRgValAry(D_IDX_START To D_IDX_START, D_IDX_START To D_IDX_START)
                wkInRgValAry(D_IDX_START, D_IDX_START) = aRgValAry
            Else
                'インデックス指定でない場合は調整
                If LBound(aRgValAry, 1) <> D_EXCEL_ROW_START Or LBound(aRgValAry, 2) <> D_EXCEL_CLM_START Then
                    wkRIStt = LBound(aRgValAry, 1)
                    wkRIEnd = UBound(aRgValAry, 1)
                    wkRIOfs = D_IDX_START - wkRIStt
                    
                    wkCIStt = UBound(aRgValAry, 2)
                    wkCIEnd = UBound(aRgValAry, 2)
                    wkCIOfs = D_IDX_START - wkCIStt
                    
                    ReDim wkInRgValAry(D_IDX_START To D_IDX_START, wkRIEnd + wkRIOfs To wkCIEnd + wkCIOfs)
                    
                    For wkRICnt = wkRIStt To wkRIEnd
                        For wkCICnt = wkCIStt To wkCIEnd
                            wkInRgValAry(wkRICnt + wkRIOfs, wkCICnt + wkCIOfs) = aRgValAry(wkRICnt, wkCICnt)
                        Next wkCICnt
                    Next wkRICnt
                End If
                
                wkInRgValAry = aRgValAry
            End If
            
            .RgValAry = wkInRgValAry
        
            With .PosInf
                .Row.Cnt = UBound(wkInRgValAry, 1)
                .Row.End = .Row.Stt + .Row.Cnt - 1
                .Clm.Cnt = UBound(wkInRgValAry, 2)
                .Clm.End = .Clm.Stt + .Clm.Cnt - 1
            End With
        End With
        
        If aInsertFlg = True Then
            .InputInsertFlg = True
        End If
    End With
End Sub

' 列インデックス指定
Public Sub S_SetInputValue( _
        ByVal aValue As Variant, _
        Optional ByVal aRow As Long = D_EXCEL_ROW_NOW, _
        Optional ByVal aCIdx As Long = D_IDX_NOW, _
        Optional ByVal aSubCIdx As Long = D_IDX_START, _
        Optional ByVal aCOfs As Integer = 0)
    Dim wkClm As Integer
    Dim wkCEnd As Integer, wkCOfs As Integer
    
    '引数チェック
    If aCOfs < 0 Then
        Exit Sub
    End If
    
    '列が取得できなかった場合は入力対象外
    If Me.F_GetColumn(wkClm, aCIdx, aSubCIdx) <> True Then
        Exit Sub
    'オフセット指定時、入力範囲が範囲外の場合は入力対象外
    ElseIf aCOfs > 0 Then
        '入力情報が配列出ない場合
        If IsArray(aValue) <> True Then
            '次チェック実施
        Else
            wkCOfs = UBound(aValue, 2) - LBound(aValue, 2)
        End If
        wkCOfs = wkCOfs + aCOfs
        
        wkCEnd = Me.F_ReturnEndColumn(aCIdx, aSubCIdx)
        If wkCEnd < (wkClm + aCOfs) Then
            Exit Sub
        End If
        
        wkClm = wkClm + aCOfs
    End If
    
    Me.S_SetInputValue_Clm aValue, aRow:=aRow, aClm:=wkClm
End Sub

'------------------------------------------------------------------------------
' 入力情報更新
'------------------------------------------------------------------------------
Public Sub S_UpdateInput( _
        Optional ByVal aInfUpdtFlg As Boolean = False)
    Dim wkSh As Worksheet
    Dim wkRg As Range
    
    Dim wkInInf As PT_INPUT_INF, wkInInf2 As PT_INPUT_INF
    Dim wkInRICnt As Long
    Dim wkInCICnt As Integer, wkInCIStt As Integer, wkInCIEnd As Integer
    
    Dim wkInPosInf As T_EXCEL_POS_INF: wkInPosInf = M_Excel.G_Excel_InitPosInf
    Dim wkInRgValAry As Variant
    Dim wkRIOfs As Long
    Dim wkCIOfs As Integer
    
    Dim wkCnt As Long, wkCnt2 As Long
    Dim wkRStt As Long, wkREnd As Long

    Dim wkTmpFlg As Boolean
    
    With pgInf
        '入力情報がなければ何もせず終了
        If .InputInfCnt < D_IDX_START Then
            Exit Sub
        End If
        
        Set wkSh = .ClsInf(E_CONFIG_IDX_CLASS_INF_SHEET)
        
        '挿入セルを処理
        For wkCnt = D_IDX_START To .InputInfCnt
            'ループ毎初期化
            wkInInf = .InputInfAry(wkCnt)
            wkTmpFlg = True
            
            '挿入の場合
            If wkInInf.InsertFlg = True Then
                '全入力情報ループし、行を調整
                For wkCnt2 = D_IDX_START To .InputInfCnt
                    wkInInf2 = .InputInfAry(wkCnt2)
                        
                    '一致の場合は無視
                    If wkCnt2 = wkCnt Then
                    '入力行が挿入行以上の場合は調整して反映
                    ElseIf wkInInf2.PosInf.Row.Stt >= wkInInf.PosInf.Row.Stt Then
                        wkInInf2.PosInf.Row.Stt = wkInInf2.PosInf.Row.Stt + wkInInf.PosInf.Row.Cnt
                        wkInInf2.PosInf.Row.End = wkInInf2.PosInf.Row.End + wkInInf.PosInf.Row.Cnt
                        .InputInfAry(wkCnt2) = wkInInf2
                    End If
                Next wkCnt2
                        
                '前回までに挿入行が未退避の場合
                If wkRStt < D_EXCEL_ROW_START Then
                    '位置を退避し、次で挿入を実施
                    wkRStt = wkInInf.PosInf.Row.Stt
                    wkREnd = wkInInf.PosInf.Row.End
                    wkTmpFlg = False
                '前回までに挿入行が退避済、かつ今回が連続している場合
                ElseIf (wkREnd + 1) = wkInInf.PosInf.Row.Stt Then
                    '終了位置を更新し次で挿入を実施
                    wkREnd = wkInInf.PosInf.Row.End
                    wkTmpFlg = False
                End If
                        
                '挿入実施要、または最終の場合は行挿入実施
                If wkTmpFlg = True Or wkCnt >= .InputInfCnt Then
                    Range(wkSh.Cells.Rows(wkRStt), wkSh.Cells.Rows(wkREnd)).Insert shift:=xlShiftDown
                            
                    '最終でない場合は今回の位置を退避
                    If wkCnt < .InputInfCnt Then
                        wkRStt = wkInInf.PosInf.Row.Stt
                        wkREnd = wkInInf.PosInf.Row.End
                    End If
                End If
            End If
            
            '全体の位置を更新
            If wkInPosInf.Row.Stt > wkInInf.PosInf.Row.Stt Or wkInPosInf.Row.Stt < D_EXCEL_ROW_START Then
                wkInPosInf.Row.Stt = wkInInf.PosInf.Row.Stt
            End If
            If wkInPosInf.Row.End < wkInInf.PosInf.Row.End Then
                wkInPosInf.Row.End = wkInInf.PosInf.Row.End
            End If
                
            If wkInPosInf.Clm.Stt > wkInInf.PosInf.Clm.Stt Or wkInPosInf.Clm.Stt < D_EXCEL_CLM_START Then
                wkInPosInf.Clm.Stt = wkInInf.PosInf.Clm.Stt
            End If
            If wkInPosInf.Clm.End < wkInInf.PosInf.Clm.End Then
                wkInPosInf.Clm.End = wkInInf.PosInf.Clm.End
            End If
        Next wkCnt
        
        '入力に反映 ※入力セル数が多い場合はエラーになるが、現状問題ない範囲でしか使用しないため一括で反映
        With wkSh
            Set wkRg = Range(.Cells(wkInPosInf.Row.Stt, wkInPosInf.Clm.Stt), .Cells(wkInPosInf.Row.End, wkInPosInf.Clm.End))
        End With
        wkInRgValAry = wkRg.Value
        
        '入力情報全ループ
        wkTmpFlg = False
        For wkCnt = D_IDX_START To .InputInfCnt
            With .InputInfAry(wkCnt)
                If .CmptFlg <> True Then
                    'オフセット算出
                    wkRIOfs = .PosInf.Row.Stt - wkInPosInf.Row.Stt
                    wkCIOfs = .PosInf.Clm.Stt - wkInPosInf.Clm.Stt
                    wkInCIStt = LBound(.RgValAry, 2)
                    wkInCIEnd = UBound(.RgValAry, 2)
                    
                    For wkInRICnt = LBound(.RgValAry, 1) To UBound(.RgValAry, 1)
                        For wkInCICnt = wkInCIStt To wkInCIEnd
                            If wkInRgValAry((wkInRICnt + wkRIOfs), (wkInCICnt + wkCIOfs)) <> .RgValAry(wkInRICnt, wkInCICnt) Then
                                wkInRgValAry((wkInRICnt + wkRIOfs), (wkInCICnt + wkCIOfs)) = .RgValAry(wkInRICnt, wkInCICnt)
                                wkTmpFlg = True
                            End If
                        Next wkInCICnt
                    Next wkInRICnt
                    
                    .CmptFlg = True
                End If
            End With
        Next wkCnt
        If wkTmpFlg = True Then
            '全表示してセルに反映
            M_Excel.S_Excel_ShowAll wkSh
            wkRg.Value = wkInRgValAry
        End If
        
        '入力情報初期化
        Erase .InputInfAry
        .InputInfCnt = D_IDX_START - 1
    End With
    
    '入力範囲更新指定あり
    If aInfUpdtFlg = True Then
        Set pgInf.ClsInf(E_CONFIG_IDX_CLASS_INF_RANGE) = wkRg
        PS_Init_Sub pgInf
    End If
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' その他設定
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' オートフィルタ設定
'------------------------------------------------------------------------------
Public Sub S_SetAutoFilter()
    Dim wkRg As Range
    
    With pgInf
        If .AFilRow < D_EXCEL_ROW_START Then
            Exit Sub
        End If
        
        'オートフィルタ設定
        Set wkRg = Range(.UsedRg.Rows(.AFilRow), .UsedRg.Rows(.UsedPosInf.Row.Cnt))
        M_Excel.S_Excel_SetAutoFilter aRg:=wkRg
    End With
End Sub

'==============================================================================
' 内部処理
'==============================================================================
' 構造体初期化
'------------------------------------------------------------------------------
Private Property Get PG_InitTitleInf() As PT_TITLE_INF
    Dim wkInf As PT_TITLE_INF
    
    With wkInf
        .TitleAry = Empty
        .PosClmInf = M_Excel.G_Excel_InitPosClmInf
        .ExistFlg = False
    End With
    
    PG_InitTitleInf = wkInf
End Property

Private Property Get PG_InitRangeValueAryInf() As PT_RGVALARY_INF
    Dim wkInf As PT_RGVALARY_INF
    
    With wkInf
        .RgValAry = Empty
        .PosRowInf = M_Excel.G_Excel_InitPosRowInf()
        .GetTime = 0#
    End With
    
    PG_InitRangeValueAryInf = wkInf
End Property

Private Property Get PG_InitInputInf() As PT_INPUT_INF
    Dim wkInf As PT_INPUT_INF
    
    With wkInf
        .PosInf = M_Excel.G_Excel_InitPosInf
        .RgValAry = Empty
        .InsertFlg = False
        .CmptFlg = False
    End With
    
    PG_InitInputInf = wkInf
End Property
