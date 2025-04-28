Attribute VB_Name = "M_File"
Option Explicit
'##############################################################################
' ファイル処理
'##############################################################################
' 参照設定          |   Microsoft Scripting Runtime
'------------------------------------------------------------------------------
' 参照モジュール    |   M_String
'------------------------------------------------------------------------------

'==============================================================================
' 公開定義
'==============================================================================
' 定数定義
'------------------------------------------------------------------------------
Public Enum E_FILE_IDX_LIST_INF
    E_FILE_IDX_LIST_INF_NONE = D_IDX_START - 1
    E_FILE_IDX_LIST_INF_FULLPATH                                                'フルパス
    E_FILE_IDX_LIST_INF_RLTPATH                                                 '相対パス
    E_FILE_IDX_LIST_INF_NAME                                                    'ファイル名
    E_FILE_IDX_LIST_INF_MAX
    E_FILE_IDX_LIST_INF_EEND = E_FILE_IDX_LIST_INF_MAX - 1
End Enum

Public Enum E_FILE_SPEC_TEXT
    E_FILE_SPEC_TEXT_NONE = &H0
    E_FILE_SPEC_TEXT_LINE = &H1
    E_FILE_SPEC_TEXT_ALL = &H2
    E_FILE_SPEC_TEXT_LINE_ALL = E_FILE_SPEC_TEXT_LINE Or E_FILE_SPEC_TEXT_ALL
End Enum

'==============================================================================
' 内部定義
'==============================================================================
' 構造体定義
'------------------------------------------------------------------------------
Private Type PT_FILE_LIST_INF
    List As Dictionary
    Path As String
    ExtSpec As String
    ExtSpecAry As Variant
End Type

'------------------------------------------------------------------------------
' 変数定義
'------------------------------------------------------------------------------
Private pgInf As PT_FILE_LIST_INF

'==============================================================================
' 公開処理
'==============================================================================
' フォルダ内ファイル情報一覧取得
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' 初期化
'------------------------------------------------------------------------------
Public Sub S_File_InitInf()
    With pgInf
        .Path = ""
        .ExtSpec = ""
        .ExtSpecAry = Empty
    End With
End Sub

'------------------------------------------------------------------------------
' フォルダ内ファイル情報一覧取得
'------------------------------------------------------------------------------
Public Function F_File_GetFolderFileInfList( _
        ByRef aRtn As Dictionary, _
        ByVal aPath As String, _
        Optional ByVal aExtSpec As String = "*.*") As Boolean
    Dim wkPath As String: wkPath = M_String.F_String_ReturnDelete(aPath, "\", E_STRING_SPEC_POS_END)
    
    Dim wkChkRet As E_RET
    Dim wkExtSpecAry As Variant
    
    '再取得チェック
    wkChkRet = PF_File_CheckFileInf(wkPath, aExtSpec)
    With pgInf
        If wkChkRet = E_RET_NG Then
            '再取得不要で終了
            Exit Function
        Else
            '再取得必要な場合は取得を実施
            If wkChkRet = E_RET_OK_1 Then
                PS_File_GetFolderFileInfList_Sub .List, .Path, "", .ExtSpecAry
            End If
        End If
        
        If .List.Count > 0 Then
            Set aRtn = .List
            F_File_GetFolderFileInfList = True
        End If
    End With
End Function

' ファイル情報チェック
Private Function PF_File_CheckFileInf( _
        ByVal aPath As String, _
        ByVal aExtSpec As String) As E_RET
    Dim wkRet As E_RET: wkRet = E_RET_NG

    Dim wkFso As New FileSystemObject
    Dim wkTmpAry As Variant, wkTmp As Variant
    
    If aPath = "" Or aExtSpec = "" Or Dir(aPath, vbDirectory) = "" Then
        'パス、拡張子指定無し、フォルダが存在しない場合は対象外
        PF_File_CheckFileInf = wkRet
        Exit Function
    End If
    
    With pgInf
        'パス、拡張子指定が一致の場合は再取得不要
        If .Path = aPath And .ExtSpec = aExtSpec Then
            wkRet = E_RET_OK
        End If
            
        '再取得時は初期化
        If wkRet <> E_RET_OK Then
            .Path = aPath
            .ExtSpec = aExtSpec
            
            If M_String.F_String_GetSplitExtension(.ExtSpecAry, .ExtSpec) <> True Then
                .ExtSpecAry = Empty
            End If
                
            If Not .List Is Nothing Then
                .List.RemoveAll
            Else
                Set .List = New Dictionary
            End If
                
            wkRet = E_RET_OK_1
        End If
    End With
    
    PF_File_CheckFileInf = wkRet
End Function

'サブルーチン
Private Sub PS_File_GetFolderFileInfList_Sub( _
        ByRef aRtn As Dictionary, _
        ByVal aFullFld As String, _
        ByVal aCrtFld As String, _
        ByVal aExtSpecAry As Variant)
    Dim wkFileInfAryAry As Variant, wkFileInfAry(D_IDX_START To E_FILE_IDX_LIST_INF_EEND) As Variant
    Dim wkKeyFld As String
    
    Dim wkFso As New FileSystemObject
    Dim wkFile As File
    Dim wkFolder As Folder
    Dim wkFileNm As String
    
    Dim wkCrtFld As String: wkCrtFld = aCrtFld
    Dim wkFullFld As String: wkFullFld = aFullFld
    Dim wkRltFld As String
    
    Dim wkExtSpec As Variant
    Dim wkAddFlg As Boolean
    
    'カレントフォルダ設定
    If wkCrtFld = "" Then
        wkCrtFld = wkFullFld
        wkRltFld = ""
    Else
        '相対フォルダパス作成（フルフォルダパスからカレントフォルダパス削除）
        wkRltFld = M_String.F_String_ReturnDelete(wkFullFld, wkCrtFld, E_STRING_SPEC_POS_START)
        wkRltFld = M_String.F_String_ReturnDelete(wkRltFld, "\", (E_STRING_SPEC_POS_START Or E_STRING_SPEC_POS_END))
    End If
    
    '全ファイル確認
    For Each wkFile In wkFso.GetFolder(wkFullFld).Files
        wkFileNm = wkFile.Name
        
        '拡張子指定がある場合
        If IsArray(aExtSpecAry) = True Then
            wkAddFlg = True
            
            '拡張子指定と一致した場合は追加でループ終了
            For Each wkExtSpec In aExtSpecAry
                If wkFileNm Like wkExtSpec Then
                    wkAddFlg = True
                    Exit For
                End If
            Next wkExtSpec
        '拡張子指定がない場合
        Else
            wkAddFlg = True
        End If
        
        '追加可能な場合
        If wkAddFlg = True Then
            wkFileInfAry(E_FILE_IDX_LIST_INF_NAME) = wkFileNm
            'フルパス設定
            wkFileInfAry(E_FILE_IDX_LIST_INF_FULLPATH) = wkFile.Path
            '相対パス設定
            wkFileInfAry(E_FILE_IDX_LIST_INF_RLTPATH) = M_String.F_String_ReturnAdd(wkRltFld, wkFileNm, aDlmt:="\")
            
            'ファイル情報追加
            wkFileInfAryAry = M_Common.F_ReturnArrayAdd(wkFileInfAryAry, wkFileInfAry)
        End If
    Next wkFile
    
    'フォルダ内ファイル登録
    If wkRltFld <> "" Then
        wkKeyFld = wkRltFld
    Else
        wkKeyFld = wkCrtFld
    End If
    aRtn.Add wkKeyFld, wkFileInfAryAry
    
    'サブフォルダ検索
    For Each wkFolder In wkFso.GetFolder(wkFullFld).SubFolders
        PS_File_GetFolderFileInfList_Sub aRtn, wkFolder.Path, wkCrtFld, aExtSpecAry
    Next wkFolder
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' テキストファイルオープン
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_File_GetOpenTextFile( _
        ByRef aRtn As TextStream, _
        ByVal aPath As String, _
        Optional ByVal aIOMode As IOMode = ForReading, _
        Optional ByVal aCreate As Boolean = False) As Boolean
    Dim wkRet As Boolean: wkRet = False
    Dim wkRtn As TextStream
    Dim wkFso As New FileSystemObject
    
    '引数チェック
    On Error GoTo PROC_ERROR
    If aPath = "" Or Dir(aPath) = "" Then
        'ファイルパス指定なし、またはファイルが存在しない場合は異常終了
        Exit Function
    End If
    
    'ファイルオープン
    Set wkRtn = wkFso.OpenTextFile(aPath, aIOMode, aCreate)
    On Error GoTo 0
    'エラー無しの場合
    If Not wkRtn Is Nothing Then
        Set aRtn = wkRtn
        wkRet = True
    End If
    
PROC_ERROR:
    F_File_GetOpenTextFile = wkRet
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' テキストファイルリード
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_File_GetReadTextFile( _
        ByRef aRtn As Variant, _
        ByRef aTs As TextStream, _
        ByVal aTextSpec As E_FILE_SPEC_TEXT, _
        Optional ByVal aPath As String = "") As Boolean
    Dim wkRet As Boolean: wkRet = False
    Dim wkRtn As Variant
    
    '引数チェック
    If aTs Is Nothing Then
        'パスでファイルオープンできなかった場合は終了
        If F_File_GetOpenTextFile(aTs, aPath, aIOMode:=ForReading, aCreate:=False) <> True Then
            Exit Function
        End If
    End If
    
    '行ごとにリードする場合
    If M_Common.F_CheckBitOn(aTextSpec, E_FILE_SPEC_TEXT_LINE) = True Then
        '全行リードする場合
        If M_Common.F_CheckBitOn(aTextSpec, E_FILE_SPEC_TEXT_ALL) = True Then
            '最終行でない間はループ
            Do While aTs.AtEndOfStream <> True
                '行を配列で設定
                wkRtn = M_Common.F_ReturnArrayAdd(wkRtn, aTs.ReadLine)
                wkRet = True
            Loop
        '1行読み込む場合は最終行でなければリード
        ElseIf aTs.AtEndOfStream <> True Then
            wkRtn = aTs.ReadLine
            wkRet = True
        End If
    'ファイル全てを読み込む場合
    ElseIf M_Common.F_CheckBitOn(aTextSpec, E_FILE_SPEC_TEXT_ALL) = True Then
        wkRtn = aTs.ReadAll
        wkRet = True
    End If
    
    If wkRet = True Then
        aRtn = wkRtn
        F_File_GetReadTextFile = True
    End If
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' テキストファイルクローズ
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Sub S_File_CloseTextFile( _
        ByRef aTs As TextStream)
    If Not aTs Is Nothing Then
        aTs.Close
        Set aTs = Nothing
    End If
End Sub
