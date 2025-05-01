Attribute VB_Name = "M_Cmd"
Option Explicit
'##############################################################################
' �R�}���h����
'##############################################################################
' �Q�Ɛݒ�          |   Windows Script Host Object Model
'------------------------------------------------------------------------------
' �Q�ƃ��W���[��    |   M_String
'------------------------------------------------------------------------------

'==============================================================================
' ���J��`
'==============================================================================
' �萔��`
'------------------------------------------------------------------------------
Public Enum E_CMD_IDX_GREP_INF
    E_CMD_IDX_GREP_INF_NONE = D_IDX_START - 1
    E_CMD_IDX_GREP_INF_RESULT
    E_CMD_IDX_GREP_INF_FULLPATH
    E_CMD_IDX_GREP_INF_RLTPATH
    E_CMD_IDX_GREP_INF_FILE
    E_CMD_IDX_GREP_INF_LINE
    E_CMD_IDX_GREP_INF_OFFSET
    E_CMD_IDX_GREP_INF_SOURCE
    E_CMD_IDX_GREP_INF_MAX
    E_CMD_IDX_GREP_INF_EEND = E_CMD_IDX_GREP_INF_MAX - 1
End Enum

'------------------------------------------------------------------------------
' �\���̒�`
'------------------------------------------------------------------------------
Public Type T_CMD_ARG_GREP_INF
    '����������
    SrchPtn As String
    
    '�P��`�F�b�N�w��
    ChkWordFlg As Boolean
    ChkWordPtn As String
    ChkWordSpec As E_STRING_SPEC
    
    FullPath As String
    ExtSpec As String
End Type

'==============================================================================
' ������`
'==============================================================================
' �萔��`
'------------------------------------------------------------------------------
Private Enum PE_CMD_POS_GREPRET
    PE_CMD_POS_GREPRET_RLTPATH = 0
    PE_CMD_POS_GREPRET_LINE
    PE_CMD_POS_GREPRET_OFFSET
    PE_CMD_POS_GREPRET_SOURCE
    PE_CMD_POS_GREPRET_MAX
    PE_CMD_POS_GREPRET_EEND = PE_CMD_POS_GREPRET_MAX - 1
End Enum

'==============================================================================
' ���J����
'==============================================================================
' �R�}���h���s���ʎ擾
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �R�}���h���s���ʎ擾
'------------------------------------------------------------------------------
Public Function F_Cmd_GetCommandResult( _
        ByRef aRtn As String, _
        ByVal aCmd As String) As Boolean
    Dim wkRet As Boolean: wkRet = False
    Dim wkRtn As String
    
    Dim wkShell As New WshShell
    Dim wkExec As Object
    
    '�����`�F�b�N
    If aCmd = "" Then
        Exit Function
    End If
    
    On Error GoTo PROC_ERROR
    '�R�}���h���s
    Set wkExec = wkShell.Exec("%ComSpec% /c " & aCmd)
    
    wkRtn = wkExec.StdOut.ReadAll
    On Error GoTo 0

    If wkRtn <> "" Then
        aRtn = wkRtn
        wkRet = True
    End If

PROC_ERROR:
    F_Cmd_GetCommandResult = wkRet
End Function

'------------------------------------------------------------------------------
' �R�}���h���s���ʔz��擾
'------------------------------------------------------------------------------
Public Function F_Cmd_GetCommandResultArray( _
        ByRef aRtnAry As Variant, _
        ByVal aCmd As String) As Boolean
    Dim wkRtnAry As Variant
    Dim wkRtnStr As String
    
    Dim wkTmpAry As Variant, wkTmp As Variant
    Dim wkTmpAry2 As Variant, wkTmp2 As Variant
    
    '�R�}���h�̌��ʂ�NG�ł���ΏI��
    If F_Cmd_GetCommandResult(wkRtnStr, aCmd) <> True Then
        Exit Function
    End If
    
    'CRLF��؂�ŕ���
    If M_String.F_String_GetSplit(wkTmpAry, wkRtnStr, vbCrLf) = True Then
        For Each wkTmp In wkTmpAry
            '��s�łȂ���Βǉ�
            If wkTmp <> "" Then
                wkRtnAry = M_Common.F_ReturnArrayAdd(wkRtnAry, wkTmp)
            End If
        Next wkTmp
    End If
    
    If IsArray(wkRtnAry) = True Then
        aRtnAry = wkRtnAry
        F_Cmd_GetCommandResultArray = True
    End If
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �R�}���h���s
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_Cmd_RunCommand( _
        ByVal aCmd As String, _
        Optional ByVal aWindowStype As Long = 0, _
        Optional ByVal aWaitOnReturn As Boolean = False) As Boolean
    Dim wkRet As Boolean: wkRet = False
    Dim wkRtn As String
    
    Dim wkShell As New WshShell
    Dim wkCmdRet As Integer
    
    '�����`�F�b�N
    If aCmd = "" Then
        Exit Function
    End If
    
    On Error GoTo PROC_ERROR
    '�R�}���h���s
    wkCmdRet = wkShell.Run(aCmd, WindowStyle:=aWindowStype, WaitOnReturn:=aWaitOnReturn)
    If aWaitOnReturn = True Then
        '����I�����Ă���ꍇ�͐����ԋp
        If wkCmdRet = 0 Then
            wkRet = True
        End If
    Else
        '��U����Ƃ���
        wkRet = True
    End If
    On Error GoTo 0
    
PROC_ERROR:
    F_Cmd_RunCommand = wkRet
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' Grep���ʎ擾
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ������񏉊���
'------------------------------------------------------------------------------
Public Property Get G_Cmd_InitArgGrepInf() As T_CMD_ARG_GREP_INF
    With G_Cmd_InitArgGrepInf
        .SrchPtn = ""
        
        .ChkWordFlg = False
        .ChkWordPtn = D_STRING_MATCH_CHECKWORD
        .ChkWordSpec = E_STRING_SPEC_MATCH_ALL
        
        .FullPath = ""
        .ExtSpec = ""
    End With
End Property

'------------------------------------------------------------------------------
' Grep���ʎ擾�i�������w��j
'------------------------------------------------------------------------------
Public Function F_Cmd_GetTextGrepResult_Inf( _
        ByRef aRtnAryAry As Variant, _
        ByRef aArgInf As T_CMD_ARG_GREP_INF) As Boolean
    Dim wkRet As Boolean
    Dim wkRtnAryAry As Variant
    
    Dim wkGrepWd As String
    Dim wkGrepPath As String
    Dim wkAddInf As T_STRING_ARG_ADD_INF: wkAddInf = M_String.G_String_InitArgAddInf
    Dim wkSrchPtn As String
    Dim wkChkPtn As String
    
    Dim wkGrepRetAry As Variant, wkGrepRet As Variant
    Dim wkGrepRetInf(D_IDX_START To E_CMD_IDX_GREP_INF_EEND) As Variant
    Dim wkRltGrepRet As String
    Dim wkFullPath As String
    Dim wkSrc As String
    Dim wkSrchInf As T_STRING_ARG_SEARCH_INF: wkSrchInf = M_String.G_String_InitArgSrchInf
    
    Dim wkCmd As String
    Dim wkExtSpecAry As Variant, wkExtSpec As Variant
    
    Dim wkTmpStr As String
    Dim wkTmpAry As Variant, wkTmp As Variant
    Dim wkTmpFlg As Boolean
    Dim wkTmpCnt As Long
    
    With aArgInf
        '�����`�F�b�N
        If .SrchPtn = "" Or .FullPath = "" Or Dir(.FullPath, vbDirectory) = "" Or .ExtSpec = "" Then
            Exit Function
        End If
        
        'Grep�p�X�쐬�i��U�S�t�@�C���ΏۂŌ����j
        With wkAddInf
            .Target = aArgInf.FullPath
            .Add = "*"
            .Dlmt = "\"
            .DlmtChkFlg = True
        End With
        wkGrepPath = M_String.F_String_ReturnAdd_Inf(wkAddInf)
        
        'Grep���{
        wkSrchPtn = PF_Cmd_ReturnWildCardRegExp2FindStr(.SrchPtn)
        wkCmd = "findstr /r /s /n /p /o """ & wkSrchPtn & """ """ & wkGrepPath
        If F_Cmd_GetCommandResultArray(wkGrepRetAry, wkCmd) <> True Then
            Exit Function
        End If
        
        '�g���q�w��𕪊�
        If M_String.F_String_GetSplitExtension(wkExtSpecAry, .ExtSpec) <> True Then
            wkExtSpecAry = M_Common.F_ReturnArrayAdd(wkExtSpecAry, "*")
        End If
        
        '������
        With wkSrchInf
            .SrchPtn = aArgInf.SrchPtn
            '�P�ꌟ������̏ꍇ�͒P��`�F�b�N�ǉ�
            If aArgInf.ChkWordFlg = True Then
                .ChkPtn = D_STRING_MATCH_CHECKWORD
                .ChkSpec = aArgInf.ChkWordSpec Or E_STRING_SPEC_MATCH_WORD
            End If
        End With
        
        'Grep���ʕ����[�v
        For Each wkGrepRet In wkGrepRetAry
            wkRltGrepRet = M_String.F_String_ReturnDelete(wkGrepRet, .FullPath, E_STRING_SPEC_POS_START)
            wkRltGrepRet = M_String.F_String_ReturnDelete(wkRltGrepRet, "\", E_STRING_SPEC_POS_START)
            
            If M_String.F_String_GetSplit(wkTmpAry, wkRltGrepRet, ":", aIncChkFlg:=True) <> True Then
                '�����ł��Ȃ������ꍇ�͖���
            ElseIf UBound(wkTmpAry) < PE_CMD_POS_GREPRET_EEND Then
                '�z�񐔂����Ȃ������ꍇ�͖���
            Else
                '�\�[�X�m�F
                wkSrc = ""
                For wkTmpCnt = PE_CMD_POS_GREPRET_SOURCE To UBound(wkTmpAry)
                    wkSrc = M_String.F_String_ReturnAdd(wkSrc, wkTmpAry(wkTmpCnt), aDlmt:=":")
                Next wkTmpCnt
                
                '�P��`�F�b�N���莞�͒P��`�F�b�N���{
                If .ChkWordFlg = True Then
                    wkSrchInf.Target = wkSrc
                    wkTmpFlg = M_String.F_String_Check_Inf(wkSrchInf)
                Else
                    wkTmpFlg = True
                End If
                If wkTmpFlg = True Then
                    wkFullPath = .FullPath & "\" & wkTmpAry(PE_CMD_POS_GREPRET_RLTPATH)
                    
                    '�g���q����
                    wkTmpFlg = False
                    For Each wkExtSpec In wkExtSpecAry
                        If wkTmpAry(PE_CMD_POS_GREPRET_RLTPATH) Like wkExtSpec Then
                            wkTmpFlg = True
                            Exit For
                        End If
                    Next wkExtSpec
                    
                    '�g���q����v�����ꍇ��Grep���ʂ𐶐����Ė߂�l�ɓo�^
                    If wkTmpFlg = True Then
                        '���N���A
                        Erase wkGrepRetInf
                        
                        wkGrepRetInf(E_CMD_IDX_GREP_INF_RESULT) = wkGrepRet
                        wkGrepRetInf(E_CMD_IDX_GREP_INF_RLTPATH) = wkTmpAry(PE_CMD_POS_GREPRET_RLTPATH)
                        wkGrepRetInf(E_CMD_IDX_GREP_INF_FULLPATH) = wkFullPath
                        wkGrepRetInf(E_CMD_IDX_GREP_INF_FILE) = M_String.F_String_ReturnDeleteStr(wkGrepRetInf(E_CMD_IDX_GREP_INF_RLTPATH), "\", aDelPosSpec:=E_STRING_SPEC_POS_START)
                        wkGrepRetInf(E_CMD_IDX_GREP_INF_LINE) = Val(wkTmpAry(PE_CMD_POS_GREPRET_LINE))
                        wkGrepRetInf(E_CMD_IDX_GREP_INF_OFFSET) = Val(wkTmpAry(PE_CMD_POS_GREPRET_OFFSET))
                        wkGrepRetInf(E_CMD_IDX_GREP_INF_SOURCE) = wkSrc
                        
                        wkRtnAryAry = M_Common.F_ReturnArrayAdd(wkRtnAryAry, wkGrepRetInf)
                    End If
                End If
            End If
        Next wkGrepRet
    End With
    
    If IsArray(wkRtnAryAry) = True Then
        aRtnAryAry = wkRtnAryAry
        F_Cmd_GetTextGrepResult_Inf = True
    End If
End Function

'------------------------------------------------------------------------------
' Grep���ʎ擾�i�����w��j
'------------------------------------------------------------------------------
Public Function F_Cmd_GetTextGrepResult( _
        ByRef aRtnAryAry As Variant, _
        ByVal aSrchPtn As String, _
        ByVal aFullPath As String, _
        Optional ByVal aExtSpec As String = "*", _
        Optional ByVal aChkWordFlg As Boolean = True) As Boolean
    Dim wkArgInf As T_CMD_ARG_GREP_INF: wkArgInf = G_Cmd_InitArgGrepInf()
    
    With wkArgInf
        .SrchPtn = aSrchPtn
        .FullPath = aFullPath
        .ExtSpec = aExtSpec
        .ChkWordFlg = aChkWordFlg
    End With
    
    F_Cmd_GetTextGrepResult = F_Cmd_GetTextGrepResult_Inf(aRtnAryAry, wkArgInf)
End Function

'==============================================================================
' ��������
'==============================================================================
' ���C���h�J�[�h�ϊ�
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' RepExp��findstr
'------------------------------------------------------------------------------
Private Function PF_Cmd_ReturnWildCardRegExp2FindStr( _
        ByVal aSrchPtn As String) As String
    Dim wkRtn As String: wkRtn = aSrchPtn
    Dim wkTmpAry As Variant, wkTmp As Variant
    
    If M_String.F_String_GetSplit(wkTmpAry, wkRtn, "|", aIncChkFlg:=True) = True Then
        wkRtn = ""
        
        For Each wkTmp In wkTmpAry
            If wkTmp <> "" Then
                wkRtn = wkRtn & " " & wkTmp
            End If
        Next wkTmp
    End If
    
    PF_Cmd_ReturnWildCardRegExp2FindStr = wkRtn
End Function
