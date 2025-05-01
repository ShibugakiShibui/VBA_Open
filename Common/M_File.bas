Attribute VB_Name = "M_File"
Option Explicit
'##############################################################################
' �t�@�C������
'##############################################################################
' �Q�Ɛݒ�          |   Microsoft Scripting Runtime
'                   |   Microsoft Office xx.x Object Library
'------------------------------------------------------------------------------
' �Q�ƃ��W���[��    |   M_String
'------------------------------------------------------------------------------

'==============================================================================
' ���J��`
'==============================================================================
' �萔��`
'------------------------------------------------------------------------------
Public Enum E_FILE_IDX_LIST_INF
    E_FILE_IDX_LIST_INF_NONE = D_IDX_START - 1
    E_FILE_IDX_LIST_INF_FULLPATH                                                '�t���p�X
    E_FILE_IDX_LIST_INF_RLTPATH                                                 '���΃p�X
    E_FILE_IDX_LIST_INF_NAME                                                    '�t�@�C����
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
' ������`
'==============================================================================
' �\���̒�`
'------------------------------------------------------------------------------
Private Type PT_FILE_LIST_INF
    List As Dictionary
    Path As String
    ExtSpec As String
    ExtSpecAry As Variant
End Type

'------------------------------------------------------------------------------
' �萔��`
'------------------------------------------------------------------------------
Private Enum PE_FILE_IDX_FILTER_INF
    PE_FILE_IDX_FILTER_INF_NONE = D_IDX_START - 1
    PE_FILE_IDX_FILTER_INF_NAME
    PE_FILE_IDX_FILTER_INF_FILTER
    PE_FILE_IDX_FILTER_INF_MAX
    PE_FILE_IDX_FILTER_INF_EEND = PE_FILE_IDX_FILTER_INF_MAX - 1
End Enum

'------------------------------------------------------------------------------
' �ϐ���`
'------------------------------------------------------------------------------
Private pgInf As PT_FILE_LIST_INF

'==============================================================================
' ���J����
'==============================================================================
' �t�H���_���t�@�C�����ꗗ�擾
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ������
'------------------------------------------------------------------------------
Public Sub S_File_InitInf()
    With pgInf
        .Path = ""
        .ExtSpec = ""
        .ExtSpecAry = Empty
    End With
End Sub

'------------------------------------------------------------------------------
' �t�H���_���t�@�C�����ꗗ�擾
'------------------------------------------------------------------------------
Public Function F_File_GetFolderFileInfList( _
        ByRef aRtn As Dictionary, _
        ByVal aPath As String, _
        Optional ByVal aExtSpec As String = "*.*") As Boolean
    Dim wkPath As String: wkPath = M_String.F_String_ReturnDelete(aPath, "\", E_STRING_SPEC_POS_END)
    
    Dim wkChkRet As E_RET
    Dim wkExtSpecAry As Variant
    
    '�Ď擾�`�F�b�N
    wkChkRet = PF_File_CheckFileInf(wkPath, aExtSpec)
    With pgInf
        If wkChkRet = E_RET_NG Then
            '�Ď擾�s�v�ŏI��
            Exit Function
        Else
            '�Ď擾�K�v�ȏꍇ�͎擾�����{
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

' �t�@�C�����`�F�b�N
Private Function PF_File_CheckFileInf( _
        ByVal aPath As String, _
        ByVal aExtSpec As String) As E_RET
    Dim wkRet As E_RET: wkRet = E_RET_NG

    Dim wkFso As New FileSystemObject
    Dim wkTmpAry As Variant, wkTmp As Variant
    
    If aPath = "" Or aExtSpec = "" Or Dir(aPath, vbDirectory) = "" Then
        '�p�X�A�g���q�w�薳���A�t�H���_�����݂��Ȃ��ꍇ�͑ΏۊO
        PF_File_CheckFileInf = wkRet
        Exit Function
    End If
    
    With pgInf
        '�p�X�A�g���q�w�肪��v�̏ꍇ�͍Ď擾�s�v
        If .Path = aPath And .ExtSpec = aExtSpec Then
            wkRet = E_RET_OK
        End If
            
        '�Ď擾���͏�����
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

'�T�u���[�`��
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
    
    '�J�����g�t�H���_�ݒ�
    If wkCrtFld = "" Then
        wkCrtFld = wkFullFld
        wkRltFld = ""
    Else
        '���΃t�H���_�p�X�쐬�i�t���t�H���_�p�X����J�����g�t�H���_�p�X�폜�j
        wkRltFld = M_String.F_String_ReturnDelete(wkFullFld, wkCrtFld, E_STRING_SPEC_POS_START)
        wkRltFld = M_String.F_String_ReturnDelete(wkRltFld, "\", (E_STRING_SPEC_POS_START Or E_STRING_SPEC_POS_END))
    End If
    
    '�S�t�@�C���m�F
    For Each wkFile In wkFso.GetFolder(wkFullFld).Files
        wkFileNm = wkFile.Name
        
        '�g���q�w�肪����ꍇ
        If IsArray(aExtSpecAry) = True Then
            wkAddFlg = True
            
            '�g���q�w��ƈ�v�����ꍇ�͒ǉ��Ń��[�v�I��
            For Each wkExtSpec In aExtSpecAry
                If wkFileNm Like wkExtSpec Then
                    wkAddFlg = True
                    Exit For
                End If
            Next wkExtSpec
        '�g���q�w�肪�Ȃ��ꍇ
        Else
            wkAddFlg = True
        End If
        
        '�ǉ��\�ȏꍇ
        If wkAddFlg = True Then
            wkFileInfAry(E_FILE_IDX_LIST_INF_NAME) = wkFileNm
            '�t���p�X�ݒ�
            wkFileInfAry(E_FILE_IDX_LIST_INF_FULLPATH) = wkFile.Path
            '���΃p�X�ݒ�
            wkFileInfAry(E_FILE_IDX_LIST_INF_RLTPATH) = M_String.F_String_ReturnAdd(wkRltFld, wkFileNm, aDlmt:="\")
            
            '�t�@�C�����ǉ�
            wkFileInfAryAry = M_Common.F_ReturnArrayAdd(wkFileInfAryAry, wkFileInfAry)
        End If
    Next wkFile
    
    '�t�H���_���t�@�C���o�^
    If wkRltFld <> "" Then
        wkKeyFld = wkRltFld
    Else
        wkKeyFld = wkCrtFld
    End If
    aRtn.Add wkKeyFld, wkFileInfAryAry
    
    '�T�u�t�H���_����
    For Each wkFolder In wkFso.GetFolder(wkFullFld).SubFolders
        PS_File_GetFolderFileInfList_Sub aRtn, wkFolder.Path, wkCrtFld, aExtSpecAry
    Next wkFolder
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �e�L�X�g�t�@�C���I�[�v��
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_File_GetOpenTextFile( _
        ByRef aRtn As TextStream, _
        ByVal aPath As String, _
        Optional ByVal aIOMode As IOMode = ForReading, _
        Optional ByVal aCreate As Boolean = False) As Boolean
    Dim wkRet As Boolean: wkRet = False
    Dim wkRtn As TextStream
    Dim wkFso As New FileSystemObject
    
    '�����`�F�b�N
    On Error GoTo PROC_ERROR
    If aPath = "" Or Dir(aPath) = "" Then
        '�t�@�C���p�X�w��Ȃ��A�܂��̓t�@�C�������݂��Ȃ��ꍇ�ُ͈�I��
        Exit Function
    End If
    
    '�t�@�C���I�[�v��
    Set wkRtn = wkFso.OpenTextFile(aPath, aIOMode, aCreate)
    On Error GoTo 0
    '�G���[�����̏ꍇ
    If Not wkRtn Is Nothing Then
        Set aRtn = wkRtn
        wkRet = True
    End If
    
PROC_ERROR:
    F_File_GetOpenTextFile = wkRet
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �e�L�X�g�t�@�C�����[�h
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_File_GetReadTextFile( _
        ByRef aRtn As Variant, _
        ByRef aTs As TextStream, _
        ByVal aTextSpec As E_FILE_SPEC_TEXT, _
        Optional ByVal aPath As String = "") As Boolean
    Dim wkRet As Boolean: wkRet = False
    Dim wkRtn As Variant
    
    '�����`�F�b�N
    If aTs Is Nothing Then
        '�p�X�Ńt�@�C���I�[�v���ł��Ȃ������ꍇ�͏I��
        If F_File_GetOpenTextFile(aTs, aPath, aIOMode:=ForReading, aCreate:=False) <> True Then
            Exit Function
        End If
    End If
    
    '�s���ƂɃ��[�h����ꍇ
    If M_Common.F_CheckBitOn(aTextSpec, E_FILE_SPEC_TEXT_LINE) = True Then
        '�S�s���[�h����ꍇ
        If M_Common.F_CheckBitOn(aTextSpec, E_FILE_SPEC_TEXT_ALL) = True Then
            '�ŏI�s�łȂ��Ԃ̓��[�v
            Do While aTs.AtEndOfStream <> True
                '�s��z��Őݒ�
                wkRtn = M_Common.F_ReturnArrayAdd(wkRtn, aTs.ReadLine)
                wkRet = True
            Loop
        '1�s�ǂݍ��ޏꍇ�͍ŏI�s�łȂ���΃��[�h
        ElseIf aTs.AtEndOfStream <> True Then
            wkRtn = aTs.ReadLine
            wkRet = True
        End If
    '�t�@�C���S�Ă�ǂݍ��ޏꍇ
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
' �e�L�X�g�t�@�C���N���[�Y
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Sub S_File_CloseTextFile( _
        ByRef aTs As TextStream)
    If Not aTs Is Nothing Then
        aTs.Close
        Set aTs = Nothing
    End If
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �J�����g�t�H���_�ړ�
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_File_MoveCurrentFolder( _
        ByVal aFolder As String) As Boolean
    Dim wkRet As Boolean
    
    Dim wkTmpAry As Variant
    
    '�����`�F�b�N
    If Dir(aFolder, vbDirectory) = "" Then
        Exit Function
    ElseIf M_String.F_String_GetSplit(wkTmpAry, aFolder, ":\", aIncChkFlg:=True) <> True Then
        Exit Function
    End If
    
    ChDrive wkTmpAry(LBound(wkTmpAry))
    ChDir aFolder
    
    F_File_MoveCurrentFolder = True
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �_�C�A���O�I�����ʎ擾
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �t�B���^�ݒ�
'------------------------------------------------------------------------------
Public Function F_File_ReturnFilterInfAdd( _
        ByRef aInfAryAry As Variant, _
        ByVal aName As String, _
        ByVal aFilter As String) As Variant
    Dim wkInf(D_IDX_START To PE_FILE_IDX_FILTER_INF_EEND) As Variant
    
    If aFilter <> "" Then
        wkInf(PE_FILE_IDX_FILTER_INF_NAME) = aName
        wkInf(PE_FILE_IDX_FILTER_INF_FILTER) = aFilter
        
        F_File_ReturnFilterInfAdd = M_Common.F_ReturnArrayAdd(aInfAryAry, wkInf)
    End If
End Function

'------------------------------------------------------------------------------
' �����I��
'------------------------------------------------------------------------------
Public Function F_File_GetDialogSelectArray( _
        ByRef aRtnAry As Variant, _
        Optional ByVal aFilDlgType As MsoFileDialogType = msoFileDialogFilePicker, _
        Optional ByVal aCrntFld As String = "", _
        Optional ByVal aFilterInfAry As Variant = Empty) As Boolean
    F_File_GetDialogSelectArray = PF_File_GetDialogSelectArray_Sub(aRtnAry, aFilDlgType, aCrntFld, aFilterInfAry, True)
End Function

' �T�u���[�`��
Private Function PF_File_GetDialogSelectArray_Sub( _
        ByRef aRtnAry As Variant, _
        ByVal aFilDlgType As MsoFileDialogType, _
        ByVal aCrntFld As String, _
        ByVal aFilterInfAry As Variant, _
        ByVal aMultiSlctFlg As Boolean) As Boolean
    Dim wkRtnAry As Variant
    Dim wkTmpAry As Variant, wkTmp As Variant
    
    '�J�����g�t�H���_�ݒ�
    F_File_MoveCurrentDirectory aCrntFld
    
    '�_�C�A���O�\��
    With Application.FileDialog(aFilDlgType)
        '�t�B���^�ݒ�
        .Filters.Clear
        If aFilDlgType = msoFileDialogFolderPicker Then
            '�t�H���_�͖���
        Else
            With .Filters
                If IsArray(aFilterInfAry) = True Then
                    For Each wkTmpAry In aFilterInfAry
                        .Add wkTmpAry(PE_FILE_IDX_FILTER_INF_NAME), wkTmpAry(PE_FILE_IDX_FILTER_INF_FILTER)
                    Next wkTmpAry
                Else
                    .Add "���ׂẴt�@�C��", "*.*"
                End If
            End With
            .FilterIndex = 1
        End If
        
        '�����t�@���I������
        .AllowMultiSelect = aMultiSlctFlg
        
        '�_�C�A���O�\��
        If .Show <> 0 Then
            '�L�����Z���ȊO�̓p�X��ԋp
            For Each wkTmp In .SelectedItems
                wkRtnAry = M_Common.F_ReturnArrayAdd(wkRtnAry, wkTmp)
            Next
        End If
    End With
    
    If IsArray(wkRtnAry) = True Then
        aRtnAry = wkRtnAry
        PF_File_GetDialogSelectArray_Sub = True
    End If
End Function

'------------------------------------------------------------------------------
' �P���I��
'------------------------------------------------------------------------------
Public Function F_File_GetDialogSelect( _
        ByRef aRtn As String, _
        Optional ByVal aFilDlgType As MsoFileDialogType = msoFileDialogFilePicker, _
        Optional ByVal aCrntFld As String = "", _
        Optional ByVal aFilterInfAry As Variant = Empty) As Boolean
    Dim wkRtnAry As Variant
    
    If PF_File_GetDialogSelectArray_Sub(wkRtnAry, aFilDlgType, aCrntFld, aFilterInfAry, False) <> True Then
        Exit Function
    End If
    
    aRtn = wkRtnAry(LBound(wkRtnAry))
    F_File_GetDialogSelect = True
End Function
