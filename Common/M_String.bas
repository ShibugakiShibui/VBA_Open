Attribute VB_Name = "M_String"
Option Explicit
'##############################################################################
' �����񏈗�
'##############################################################################
' �Q�Ɛݒ�          |   Microsoft VBScript Regular Expressions 5.5
'------------------------------------------------------------------------------
' �Q�ƃ��W���[��    |   �\
'------------------------------------------------------------------------------
' ���ʃo�[�W����    |   250501
'------------------------------------------------------------------------------

'==============================================================================
' ���J��`
'==============================================================================
' �萔��`
'------------------------------------------------------------------------------
Public Enum E_STRING_SPEC
    E_STRING_SPEC_NONE = &H0
    E_STRING_SPEC_POS_START = &H1
    E_STRING_SPEC_POS_END = &H2
    E_STRING_SPEC_POS_MID = &H4
    E_STRING_SPEC_POS_BOTH = E_STRING_SPEC_POS_START Or E_STRING_SPEC_POS_END
    E_STRING_SPEC_MATCH_START = &H10
    E_STRING_SPEC_MATCH_END = &H20
    E_STRING_SPEC_MATCH_MID = &H40
    E_STRING_SPEC_MATCH_WORD = &H80
    E_STRING_SPEC_MATCH_ALL = E_STRING_SPEC_MATCH_START Or E_STRING_SPEC_MATCH_END
    E_STRING_SPEC_MATCH_MASK = E_STRING_SPEC_MATCH_START Or E_STRING_SPEC_MATCH_END Or E_STRING_SPEC_MATCH_MID
End Enum

Public Enum E_STRING_IDX_SRCH_INF
    E_STRING_IDX_SRCH_INF_NONE = D_IDX_START - 1
    E_STRING_IDX_SRCH_INF_POS_START
    E_STRING_IDX_SRCH_INF_LENGTH
    E_STRING_IDX_SRCH_INF_MAX
    E_STRING_IDX_SRCH_INF_EEND = E_STRING_IDX_SRCH_INF_MAX - 1
End Enum

Public Const D_STRING_MATCH_CHECKWORD As String = "[^A-Za-z0-9_]"

Public Const D_STRING_DLMT_EXTENSION As String = ";"

'------------------------------------------------------------------------------
' �\���̒�`
'------------------------------------------------------------------------------
Public Type T_STRING_ARG_ADD_INF
    '������
    Target As String
    
    '��؂�
    Dlmt As String
    DlmtChkFlg As Boolean
    
    '�ǉ�
    Add As String
    AddChkFlg As Boolean
    AddSpec As E_STRING_SPEC
    
    '�ΏۊO
    Excluded As String
End Type

Public Type T_STRING_ARG_SEARCH_INF
    '������
    Target As String
    
    '�����w��
    Search As String
    SrchSpec As E_STRING_SPEC
    SrchPtn As String
    
    '��v�w��
    ChkPtn As String
    ChkSpec As E_STRING_SPEC
    ChkPtnOfs As Long
    
    '�����ʒu�w��
    SttPos As Long
    EndPos As Long
    Length As Long
    
    '�擾�ʒu�w��
    GetIdx As Long
End Type

Public Type T_STRING_ARG_GET_INF
    SrchInf As T_STRING_ARG_SEARCH_INF
    
    SttStr As String
    EndStr As String
    
    AddBefFlg As Boolean
    AddSrchFlg As Boolean
End Type

Public Type T_STRING_ARG_DEL_INF
    SrchInf As T_STRING_ARG_SEARCH_INF
    
    DelPosSpec As E_STRING_SPEC
    AddDelFlg As Boolean
End Type

'==============================================================================
' ������`
'==============================================================================

'==============================================================================
' ���J����
'==============================================================================
' ������ǉ�
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ������񏉊���
'------------------------------------------------------------------------------
Public Property Get G_String_InitArgAddInf() As T_STRING_ARG_ADD_INF
    Dim wkInf As T_STRING_ARG_ADD_INF
    
    With wkInf
        .Target = ""
        
        .Dlmt = ""
        .DlmtChkFlg = True
        
        .Add = ""
        .AddChkFlg = False
        .AddSpec = E_STRING_SPEC_POS_END
        
        .Excluded = ""
    End With
    
    G_String_InitArgAddInf = wkInf
End Property

'------------------------------------------------------------------------------
' ������ǉ��i�������w��j
'------------------------------------------------------------------------------
Public Function F_String_ReturnAdd_Inf( _
        ByRef aArgInf As T_STRING_ARG_ADD_INF) As String
    Dim wkRtn As String
    
    Dim wkAddChkFlg As Boolean
    Dim wkTmpStr As String
    
    With aArgInf
        '������
        wkRtn = .Target
        wkAddChkFlg = .AddChkFlg
        
        '�ǉ�����̏ꍇ
        If .Add <> "" Then
            '��؂蕶���ڑ�
            If wkRtn = "" Then
                '�ǉ����������ꍇ�͖���
            ElseIf PF_String_GetAdd_Sub(wkRtn, .Dlmt, .AddSpec, .DlmtChkFlg, .Excluded) = True Then
                '��؂�ǉ������ꍇ�͒ǉ��`�F�b�N��L����
                wkAddChkFlg = True
            End If
            
            '�ǉ������ڑ�
            PF_String_GetAdd_Sub wkRtn, .Add, .AddSpec, wkAddChkFlg, .Excluded
        End If
    End With
    
    F_String_ReturnAdd_Inf = wkRtn
End Function

' �T�u���[�`��
Private Function PF_String_GetAdd_Sub( _
        ByRef aRtn As String, _
        ByVal aAdd As String, ByVal aAddSpec As E_STRING_SPEC, _
        ByVal aAddChkFlg As Boolean, ByVal aExcluded As String) As Boolean
    Dim wkRet As Boolean: wkRet = True
    
    Dim wkChkStr As String
    Dim wkExcStr As String
    
    If aAdd <> "" Then
        If aAddChkFlg = True Then
            '�ǉ��ʒu������w��̏ꍇ
            If aAddSpec = E_STRING_SPEC_POS_END Then
                wkChkStr = Right(aRtn, Len(aAdd))
                If aExcluded <> "" Then
                    wkExcStr = Right(aRtn, Len(aExcluded))
                End If
            '�ǉ��ʒu���O���w��̏ꍇ
            Else
                wkChkStr = Left(aRtn, Len(aAdd))
                If aExcluded <> "" Then
                    wkExcStr = Left(aRtn, Len(aExcluded))
                End If
            End If
            
            '�ڑ��ʒu�ɕ���������ꍇ
            If StrComp(wkChkStr, aAdd, vbBinaryCompare) = 0 Then
                '�ΏۊO�����������ꍇ�͒ǉ��ΏۊO
                If aExcluded = "" Then
                    wkRet = False
                '�ΏۊO�����ƕs��v�̏ꍇ�͒ǉ��ΏۊO
                ElseIf StrComp(wkExcStr, aExcluded, vbBinaryCompare) <> 0 Then
                    wkRet = False
                End If
            End If
        End If
        
        '�ǉ�����̏ꍇ
        If wkRet = True Then
            '�ǉ��ʒu�w��ɏ]���ĕ�����ǉ�
            If aAddSpec = E_STRING_SPEC_POS_END Then
                aRtn = aRtn & aAdd
            Else
                aRtn = aAdd & aRtn
            End If
        End If
    End If
    
    PF_String_GetAdd_Sub = wkRet
End Function

'------------------------------------------------------------------------------
' ������ǉ��i�����w��j
'------------------------------------------------------------------------------
Public Function F_String_ReturnAdd( _
        ByVal aTarget As String, ByVal aAdd As String, _
        Optional ByVal aDlmt As String = "", _
        Optional ByVal aAddSpec As E_STRING_SPEC = E_STRING_SPEC_POS_END, _
        Optional ByVal aExcluded As String = "") As String
    Dim wkArgInf As T_STRING_ARG_ADD_INF: wkArgInf = G_String_InitArgAddInf()
    
    With wkArgInf
        .Target = aTarget
        .Dlmt = aDlmt
        .Add = aAdd
        .AddSpec = aAddSpec
        .Excluded = aExcluded
    End With
    
    F_String_ReturnAdd = F_String_ReturnAdd_Inf(wkArgInf)
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �������z��擾
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ������񏉊���
'------------------------------------------------------------------------------
Public Property Get G_String_InitArgSrchInf() As T_STRING_ARG_SEARCH_INF
    Dim wkInf As T_STRING_ARG_SEARCH_INF
    
    With wkInf
        .Target = ""
        
        .Search = ""
        .SrchSpec = E_STRING_SPEC_MATCH_MID
        .SrchPtn = ""
        
        .ChkPtn = ""
        .ChkSpec = E_STRING_SPEC_MATCH_ALL
        .ChkPtnOfs = 1
        
        .SttPos = D_POS_START
        .EndPos = D_POS_END
        .Length = D_POS_END
        
        .GetIdx = D_IDX_ALL
    End With
    
    G_String_InitArgSrchInf = wkInf
End Property

'------------------------------------------------------------------------------
' �������z��擾�i�������w��j
'------------------------------------------------------------------------------
Public Function F_String_GetSearchInfArray_Inf( _
        ByRef aRtnAryAry As Variant, _
        ByRef aArgInf As T_STRING_ARG_SEARCH_INF) As Boolean
    Dim wkRtnAryAry As Variant
    Dim wkRtnAry(D_IDX_START To E_STRING_IDX_SRCH_INF_EEND) As Variant
    Dim wkAddFlg As Boolean
    
    Dim wkArgInf As T_STRING_ARG_SEARCH_INF: wkArgInf = aArgInf
    Dim wkSttPos As Long, wkEndPos As Long
    Dim wkTgtLen As Long
    
    Dim wkSrchPtn As String
    Dim wkSrchRegExp As RegExp
    Dim wkMatch As Variant
    Dim wkMatchCnt As Long
    
    Dim wkChkPtn As String
    Dim wkChkRegExp As RegExp
    Dim wkChkStt As Long, wkChkEnd As Long
    
    Dim wkGetIdx As Long
    Dim wkRtnIdx As Long
    
    On Error GoTo PROC_ERROR
    With aArgInf
        '�����`�F�b�N
        If .Target = "" Or (.Search = "" And .SrchPtn = "") Or .SttPos < D_POS_START Then
            Exit Function
        End If
        
        '�I�[�ʒu����
        wkSttPos = .SttPos
        wkEndPos = .EndPos
        wkTgtLen = Len(.Target)
        If PF_String_GetPosEndAdjust(wkEndPos, Len(.Target), wkSttPos, .Length) <> True Then
            Exit Function
        End If
        
        '�����p�^�[���ێ�
        wkSrchPtn = .SrchPtn
        If wkSrchPtn = "" Then
            '�����p�^�[�����쐬
            wkSrchPtn = PF_String_ReturnSearchPattern(.Search, .SrchSpec)
        End If
        
        '�����p�^�[���ݒ�A�������{
        Set wkSrchRegExp = New RegExp
        With wkSrchRegExp
            .IgnoreCase = False
            .Global = True
            .Pattern = wkSrchPtn
        End With
        Set wkMatch = wkSrchRegExp.Execute(Mid(.Target, wkSttPos, (wkEndPos - wkSttPos + 1)))
        '�������ʂ�������ΏI��
        If wkMatch.Count <= 0 Then
            Exit Function
        End If
        
        '��v�p�^�[��������ꍇ
        If .ChkPtn <> "" And .ChkPtnOfs > 0 Then
            '�`�F�b�N�w�肪����ꍇ
            If M_Common.F_CheckBitOn(.ChkSpec, E_STRING_SPEC_MATCH_MASK) = True Then
                '��v�p�^�[���쐬
                If M_Common.F_CheckBitOn(.ChkSpec, E_STRING_SPEC_MATCH_WORD) = True Then
                    wkChkPtn = PF_String_ReturnCheckPatternWord(wkSrchPtn, .SrchSpec, .ChkPtn, .ChkSpec)
                Else
                    wkChkPtn = PF_String_ReturnCheckPattern(wkSrchPtn, .SrchSpec, .ChkPtn, .ChkSpec)
                End If
            End If
            
            '�`�F�b�N�p�^�[��������ꍇ�A�`�F�b�N�p�^�[���ݒ�
            If wkChkPtn <> "" Then
                '�`�F�b�N�����ݒ�
                Set wkChkRegExp = New RegExp
                With wkChkRegExp
                    .IgnoreCase = False
                    .Global = False
                    .Pattern = wkChkPtn
                End With
            End If
        End If
        
        '�����q�b�g�����A�ʒu���𒊏o
        wkGetIdx = D_IDX_START
        wkRtnIdx = D_IDX_START
        For wkMatchCnt = 0 To wkMatch.Count - 1
            '������
            wkAddFlg = True
            wkRtnAry(E_STRING_IDX_SRCH_INF_POS_START) = wkSttPos + wkMatch.Item(wkMatchCnt).FirstIndex
            wkRtnAry(E_STRING_IDX_SRCH_INF_LENGTH) = wkMatch.Item(wkMatchCnt).Length
                
            If Not wkChkRegExp Is Nothing Then
                '�`�F�b�N�ʒu����
                wkChkStt = wkRtnAry(E_STRING_IDX_SRCH_INF_POS_START)
                wkChkEnd = wkChkStt + wkRtnAry(E_STRING_IDX_SRCH_INF_LENGTH) - 1
    
                '�J�n�ʒu����
                wkChkStt = wkChkStt - .ChkPtnOfs
                If wkChkStt < D_POS_START Then
                    wkChkStt = D_POS_START
                End If
                '�I���ʒu����
                wkChkEnd = wkChkEnd + .ChkPtnOfs
                If wkChkEnd > wkTgtLen Then
                    wkChkEnd = wkTgtLen
                End If
                    
                wkAddFlg = wkChkRegExp.Test(Mid(.Target, wkChkStt, (wkChkEnd - wkChkStt + 1)))
            End If
                
            If wkAddFlg = True Then
                '�擾�C���f�b�N�X���ݒ�Ȃ��܂��͎擾�C���f�b�N�X����v�̏ꍇ�A�߂�����ɐݒ�
                If .GetIdx < D_IDX_START Or wkGetIdx = .GetIdx Then
                    '�擾�J�E���^���S�擾�̏ꍇ�͐ݒ�ʒu�𒲐�
                    If .GetIdx = D_IDX_ALL Then
                        wkRtnIdx = wkGetIdx
                    End If
                    '�擾���ʂ�z��ɓo�^
                    wkRtnAryAry = M_Common.F_ReturnArrayAdd(wkRtnAryAry, wkRtnAry, aIdx:=wkRtnIdx)
                    
                    '�C���f�b�N�X����v�̏ꍇ�͎擾�����̂��߃��[�v�I��
                    If wkGetIdx = .GetIdx Then
                        Exit For
                    End If
                End If
                
                '�擾�J�E���^�X�V
                wkGetIdx = wkGetIdx + 1
            End If
        Next wkMatchCnt
    End With
    On Error GoTo 0
    
    '�w��p�^�[�������������ꍇ
    If IsArray(wkRtnAryAry) = True Then
        aRtnAryAry = wkRtnAryAry
        F_String_GetSearchInfArray_Inf = True
    End If
    
PROC_ERROR:
    '�������Ȃ�
End Function

'------------------------------------------------------------------------------
' �������z��擾�i�����w��j
'------------------------------------------------------------------------------
Public Function F_String_GetSearchInfArray( _
        ByRef aRtnAryAry As Variant, _
        ByVal aTarget As String, _
        Optional ByVal aSearch As String = "", Optional ByVal aSrchSpec As E_STRING_SPEC = E_STRING_SPEC_MATCH_MID, Optional ByVal aSrchPtn As String = "") As Boolean
    Dim wkArgInf As T_STRING_ARG_SEARCH_INF: wkArgInf = G_String_InitArgSrchInf()
    
    With wkArgInf
        .Target = aTarget
        .Search = aSearch
        .SrchSpec = aSrchSpec
        .SrchPtn = aSrchPtn
    End With
    
    F_String_GetSearchInfArray = F_String_GetSearchInfArray_Inf(aRtnAryAry, wkArgInf)
End Function

'------------------------------------------------------------------------------
' �P�ꌟ�����z��擾�i�����w��j
'------------------------------------------------------------------------------
Public Function F_String_GetSearchInfArrayWord( _
        ByRef aRtnAryAry As Variant, _
        ByVal aTarget As String, _
        Optional ByVal aSearch As String = "", Optional ByVal aSrchSpec As E_STRING_SPEC = E_STRING_SPEC_MATCH_MID, Optional ByVal aSrchPtn As String = "", _
        Optional ByVal aChkPtn As String = D_STRING_MATCH_CHECKWORD, Optional ByVal aChkSpec As E_STRING_SPEC = E_STRING_SPEC_MATCH_ALL) As Boolean
    Dim wkArgInf As T_STRING_ARG_SEARCH_INF: wkArgInf = G_String_InitArgSrchInf()
    
    '�������ɐݒ�
    With wkArgInf
        .Target = aTarget
        .Search = aSearch
        .SrchSpec = aSrchSpec
        
        .ChkPtn = aChkPtn
        .ChkSpec = aChkSpec Or E_STRING_SPEC_MATCH_WORD
    End With
    
    F_String_GetSearchInfArrayWord = F_String_GetSearchInfArray_Inf(aRtnAryAry, wkArgInf)
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �����񑶍݃`�F�b�N
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �����񑶍݃`�F�b�N�i�������w��j
'------------------------------------------------------------------------------
Public Function F_String_Check_Inf( _
        ByRef aArgInf As T_STRING_ARG_SEARCH_INF) As Boolean
    Dim wkTmpAryAry As Variant
    
    F_String_Check_Inf = F_String_GetSearchInfArray_Inf(wkTmpAryAry, aArgInf)
End Function

'------------------------------------------------------------------------------
' �����񑶍݃`�F�b�N�i�����w��j
'------------------------------------------------------------------------------
Public Function F_String_Check( _
        ByVal aTarget As String, _
        Optional ByVal aSearch As String = "", Optional ByVal aSrchSpec As E_STRING_SPEC = E_STRING_SPEC_MATCH_MID, Optional ByVal aSrchPtn As String = "") As Boolean
    Dim wkArgInf As T_STRING_ARG_SEARCH_INF: wkArgInf = G_String_InitArgSrchInf()
    
    With wkArgInf
        .Target = aTarget
        
        .Search = aSearch
        .SrchSpec = aSrchSpec
        .SrchPtn = aSrchPtn
    End With
    
    F_String_Check = F_String_Check_Inf(wkArgInf)
End Function

'------------------------------------------------------------------------------
' �P�ꑶ�݃`�F�b�N�i�����w��j
'------------------------------------------------------------------------------
Public Function F_String_CheckWord( _
        ByVal aTarget As String, _
        Optional ByVal aSearch As String = "", Optional ByVal aSrchSpec As E_STRING_SPEC = E_STRING_SPEC_MATCH_MID, Optional ByVal aSrchPtn As String = "", _
        Optional ByVal aChkPtn As String = D_STRING_MATCH_CHECKWORD, Optional ByVal aChkSpec As E_STRING_SPEC = E_STRING_SPEC_MATCH_ALL) As Boolean
    Dim wkTmpAryAry As Variant
    
    F_String_CheckWord = F_String_GetSearchInfArrayWord(wkTmpAryAry, aTarget, aSearch:=aSearch, aSrchSpec:=aSrchSpec, aSrchPtn:=aSrchPtn, _
                                                        aChkPtn:=aChkPtn, aChkSpec:=aChkSpec)

End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �����񕪊�
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_String_GetSplit( _
        ByRef aRtnAry As Variant, _
        ByVal aTarget As String, ByVal aDlmt As String, _
        Optional ByVal aIncChkFlg As Boolean = False) As Boolean
    Dim wkRtnAry As Variant
    
    '�����`�F�b�N
    If aTarget = "" Or aDlmt = "" Then
        Exit Function
    End If
    
    '��؂�ŕ���
    wkRtnAry = Split(aTarget, aDlmt)
    
    '��؂��܃`�F�b�N����̏ꍇ�A��؂薳����NG�ŏI��
    If aIncChkFlg = True Then
        If UBound(wkRtnAry) <= LBound(wkRtnAry) Then
            Exit Function
        End If
    End If
    
    aRtnAry = wkRtnAry
    F_String_GetSplit = True
End Function

'------------------------------------------------------------------------------
' �g���q�w�蕪��
'------------------------------------------------------------------------------
Public Function F_String_GetSplitExtension( _
        ByRef aRtnAry As Variant, _
        ByVal aExtSpec As String) As Boolean
    Dim wkRtnAry As Variant
    Dim wkTmpAry As Variant, wkTmp As Variant
    
    '�����񕪊��i�����`�F�b�N���p�j
    If F_String_GetSplit(wkTmpAry, aExtSpec, D_STRING_DLMT_EXTENSION) <> True Then
        Exit Function
    End If
    
    '�����S�ă��[�v
    For Each wkTmp In wkTmpAry
        '�󔒂łȂ���Βǉ�
        wkTmp = Trim(wkTmp)
        If wkTmp <> "" Then
            wkRtnAry = M_Common.F_ReturnArrayAdd(wkTmpAry, wkTmp)
        End If
    Next wkTmp
    
    If IsArray(wkRtnAry) = True Then
        aRtnAry = wkRtnAry
        F_String_GetSplitExtension = True
    End If
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �w�蕶����Ԏ擾
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ������񏉊���
'------------------------------------------------------------------------------
Public Property Get G_String_InitArgGetInf() As T_STRING_ARG_GET_INF
    Dim wkInf As T_STRING_ARG_GET_INF
    
    With wkInf
        .SrchInf = G_String_InitArgSrchInf
        
        .SttStr = ""
        .EndStr = ""
        
        .AddBefFlg = False
        .AddSrchFlg = False
    End With
    
    G_String_InitArgGetInf = wkInf
End Property

'------------------------------------------------------------------------------
' �w�蕶����Ԏ擾�i�������w��j
'------------------------------------------------------------------------------
Public Function F_String_GetMidStr_Inf( _
        ByRef aRtn As String, _
        ByRef aArgInf As T_STRING_ARG_GET_INF) As Boolean
    Dim wkRet As Boolean
    Dim wkRtn As String
    
    Dim wkSttInfAry As Variant, wkSttInf As Variant
    Dim wkEndInfAry As Variant, wkEndInf As Variant
    Dim wkArgSrchInf As T_STRING_ARG_SEARCH_INF
    
    Dim wkSttCnt As Long, wkSttNow As Long
    Dim wkChkCnt As Long
    
    Dim wkGetSttInf As Variant
    Dim wkGetEndInf As Variant
    Dim wkGetSttPos As Long, wkGetEndPos As Long
    
    With aArgInf
        wkArgSrchInf = aArgInf.SrchInf
        
        wkArgSrchInf.Search = .SttStr
        '�J�n�ʒu�擾�i�����`�F�b�N���p�j
        If F_String_GetSearchInfArray_Inf(wkSttInfAry, wkArgSrchInf) <> True Then
            Exit Function
        End If
        
        wkArgSrchInf.Search = .EndStr
        '�I���ʒu�擾�i�����`�F�b�N���p�j
        If F_String_GetSearchInfArray_Inf(wkEndInfAry, wkArgSrchInf) <> True Then
            Exit Function
        End If
    End With
    
    '�I���ʒu�őS���[�v
    wkSttNow = LBound(wkSttInfAry) - 1
    For Each wkEndInf In wkEndInfAry
        '�J�n�ʒu�Ń��[�v
        For wkSttCnt = wkSttNow + 1 To UBound(wkSttInfAry)
            wkSttInf = wkSttInfAry(wkSttCnt)
            
            '�J�n�ʒu���I���ʒu�ȏ�̏ꍇ�̓��[�v�I��
            If wkSttInf(E_STRING_IDX_SRCH_INF_POS_START) >= wkEndInf(E_STRING_IDX_SRCH_INF_POS_START) Then
                Exit For
            End If
            
            '�e��X�V
            wkSttNow = wkSttCnt
            wkChkCnt = wkChkCnt + 1
            
            '����̏ꍇ�͊J�n�ʒu�ێ�
            If IsArray(wkGetSttInf) <> True Then
                wkGetSttInf = wkSttInf
            End If
        Next wkSttCnt
        
        '�J�n�ʒu���������Ă���ꍇ�͏I���ʒu�`�F�b�N
        If IsArray(wkGetSttInf) = True Then
            wkChkCnt = wkChkCnt - 1
            
            '��������ꂽ�ꍇ�̓��[�v�I��
            If wkChkCnt <= 0 Then
                wkGetEndInf = wkEndInf
                wkRet = True
                Exit For
            End If
        End If
    Next wkEndInf
    
    '�擾���莞�͕�����ԋp
    If wkRet = True Then
        With aArgInf
            '�J�n�ʒu����
            '�����O������ǉ��w�肠��̏ꍇ
            If .AddBefFlg = True Then
                wkGetSttPos = .SrchInf.SttPos
            '����������ǉ��w��Ȃ��̏ꍇ
            ElseIf .AddSrchFlg <> True Then
                wkGetSttPos = wkGetSttInf(E_STRING_IDX_SRCH_INF_POS_START) + wkGetSttInf(E_STRING_IDX_SRCH_INF_LENGTH)
            End If
        
            '�I���ʒu����
            '�����O������ǉ��w�肠��A�܂��͌���������ǉ��w�肠��̏ꍇ
            If .AddBefFlg = True Or .AddSrchFlg = True Then
                wkGetEndPos = wkGetEndInf(E_STRING_IDX_SRCH_INF_POS_START) - wkGetEndInf(E_STRING_IDX_SRCH_INF_LENGTH) + 1
            '������ǉ��w��Ȃ��̏ꍇ
            Else
                wkGetEndPos = wkGetEndInf(E_STRING_IDX_SRCH_INF_POS_START) - 1
            End If
            
            If wkGetSttPos <= wkGetEndPos Then
                wkRtn = Mid(.SrchInf.Target, wkGetSttPos, (wkGetEndPos - wkGetSttPos + 1))
            End If
        End With
        
        If wkRtn <> "" Then
            aRtn = wkRtn
            F_String_GetMidStr_Inf = True
        End If
    End If
End Function

'------------------------------------------------------------------------------
' �w�蕶����Ԏ擾�i�����w��j
'------------------------------------------------------------------------------
Public Function F_String_GetMidStr( _
        ByRef aRtn As String, _
        ByVal aTarget As String, _
        ByVal aSttStr As String, ByVal aEndStr As String, _
        Optional ByVal aAddBefFlg As Boolean = False, _
        Optional ByVal aAddSrchFlg As Boolean = False) As Boolean
    Dim wkArgInf As T_STRING_ARG_GET_INF: wkArgInf = G_String_InitArgGetInf()
    
    With wkArgInf
        With .SrchInf
            .Target = aTarget
            .GetIdx = D_IDX_START
        End With
        
        .SttStr = aSttStr
        .EndStr = aEndStr
        .AddBefFlg = aAddBefFlg
        .AddSrchFlg = aAddSrchFlg
    End With
    
    F_String_GetMidStr = F_String_GetMidStr_Inf(aRtn, wkArgInf)
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �w�蕶����폜
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_String_ReturnDelete( _
        ByVal aTarget As String, _
        ByVal aDelete As String, ByVal aDelSpec As E_STRING_SPEC) As String
    Dim wkRtn As String: wkRtn = aTarget
    Dim wkLen As Long
    Dim wkDelLen As Long
    
    '������ƍ폜����������ꍇ
    If aTarget <> "" And aDelete <> "" Then
        '���Ԉʒu�폜�̏ꍇ
        If M_Common.F_CheckBitOn(aDelSpec, E_STRING_SPEC_POS_MID) = True Then
            wkRtn = Replace(aTarget, aDelete, "")
        Else
            wkLen = Len(aTarget)
            wkDelLen = Len(aDelete)
            
            '�J�n�ʒu�폜�w��̏ꍇ
            If M_Common.F_CheckBitOn(aDelSpec, E_STRING_SPEC_POS_START) = True Then
                '�J�n�ʒu�ɍ폜����������ԃ��[�v
                Do While StrComp(Left(wkRtn, wkDelLen), aDelete, vbBinaryCompare) = 0
                    wkLen = wkLen - wkDelLen
                    wkRtn = Right(wkRtn, wkLen)
                Loop
            End If
            
            '�I���ʒu�폜�w��̏ꍇ
            If M_Common.F_CheckBitOn(aDelSpec, E_STRING_SPEC_POS_END) = True Then
                '�I���ʒu�ɍ폜����������ԃ��[�v
                Do While StrComp(Right(wkRtn, wkDelLen), aDelete, vbBinaryCompare) = 0
                    wkLen = wkLen - wkDelLen
                    wkRtn = Left(wkRtn, wkLen)
                Loop
            End If
        End If
    End If
    
    F_String_ReturnDelete = wkRtn
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �w�蕶����ȑO�A�ȍ~�폜
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ������񏉊���
'------------------------------------------------------------------------------
Public Property Get G_String_InitArgDelInf() As T_STRING_ARG_DEL_INF
    Dim wkInf As T_STRING_ARG_DEL_INF
    
    With G_String_InitArgDelInf
        .SrchInf = G_String_InitArgSrchInf
        
        .DelPosSpec = E_STRING_SPEC_POS_START
        .AddDelFlg = False
    End With
    
    G_String_InitArgDelInf = wkInf
End Property

'------------------------------------------------------------------------------
' �w�蕶����ȍ~�A�ȑO�폜�i�������w��j
'------------------------------------------------------------------------------
Public Function F_String_ReturnDeleteStr_Inf( _
        ByRef aArgInf As T_STRING_ARG_DEL_INF) As String
    Dim wkRtn As String
    
    Dim wkArgInf As T_STRING_ARG_DEL_INF: wkArgInf = aArgInf
    Dim wkInfAryAry As Variant, wkInfAry As Variant
    Dim wkDelStt As Long, wkDelEnd As Long
    Dim wkStrLen As Long
    
    '������ʒu�擾
    With wkArgInf
        wkRtn = .SrchInf.Target
        
        '����������쐬
        If F_String_GetSearchInfArray_Inf(wkInfAryAry, .SrchInf) <> True Then
            '����������������Ȃ������ꍇ�͖���
        Else
            wkStrLen = Len(.SrchInf.Target)
            
            '�J�n�`������ʒu�܂ō폜
            If M_Common.F_CheckBitOn(.DelPosSpec, E_STRING_SPEC_POS_START) = True Then
                '�폜�J�n�ʒu��ݒ�
                wkDelStt = .SrchInf.SttPos
                
                '�폜�I���ʒu��ݒ�
                wkDelEnd = wkInfAryAry(LBound(wkInfAryAry))(E_STRING_IDX_SRCH_INF_POS_START)
                '�폜������ǉ��w�肠��̏ꍇ
                If .AddDelFlg = True Then
                    wkDelEnd = wkDelEnd - 1
                '�폜������ǉ��w��Ȃ��̏ꍇ
                Else
                    wkDelEnd = wkDelEnd + wkInfAryAry(LBound(wkInfAryAry))(E_STRING_IDX_SRCH_INF_LENGTH) - 1
                End If
            '������ʒu�`�I���܂ō폜
            Else
                '�폜�J�n�ʒu��ݒ�
                wkDelStt = wkInfAryAry(LBound(wkInfAryAry))(E_STRING_IDX_SRCH_INF_POS_START)
                '�폜������ǉ��w�肠��̏ꍇ
                If .AddDelFlg = True And wkDelStt > D_POS_START Then
                    wkDelStt = wkDelStt + wkInfAryAry(LBound(wkInfAryAry))(E_STRING_IDX_SRCH_INF_LENGTH)
                End If
                
                '�폜�I���ʒu��ݒ�
                wkDelEnd = .SrchInf.EndPos
                If wkDelEnd < D_POS_START Then
                    PF_String_GetPosEndAdjust wkDelEnd, wkStrLen, .SrchInf.SttPos, .SrchInf.Length
                End If
            End If
            
            '�폜�ʒu�ɖ��Ȃ��ꍇ�A�폜���{
            If wkDelStt <= wkDelEnd Then
                wkRtn = ""
                If wkDelStt > 1 Then
                    wkRtn = Left(.SrchInf.Target, wkDelStt - 1)
                End If
                If wkDelEnd < wkStrLen Then
                    wkRtn = wkRtn & Right(.SrchInf.Target, (wkStrLen - wkDelEnd))
                End If
            End If
        End If
    End With
    
    F_String_ReturnDeleteStr_Inf = wkRtn
End Function

'------------------------------------------------------------------------------
' �w�蕶����ȍ~�A�ȑO�폜�i�����w��j
'------------------------------------------------------------------------------
Public Function F_String_ReturnDeleteStr( _
        ByVal aTarget As String, _
        ByVal aDelete As String, _
        Optional ByVal aDelPosSpec As E_STRING_SPEC = E_STRING_SPEC_POS_START, _
        Optional ByVal aDelIdx As Long = D_IDX_START) As String
    Dim wkArgInf As T_STRING_ARG_DEL_INF: wkArgInf = G_String_InitArgDelInf
    
    With wkArgInf
        .SrchInf.Target = aTarget
        .SrchInf.Search = aDelete
        .SrchInf.GetIdx = aDelIdx
        
        .DelPosSpec = aDelPosSpec
    End With
    
    F_String_ReturnDeleteStr = F_String_ReturnDeleteStr_Inf(wkArgInf)
End Function

'==============================================================================
' ��������
'==============================================================================
' �p�^�[���쐬
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �����p�^�[���쐬
'------------------------------------------------------------------------------
Private Function PF_String_ReturnSearchPattern( _
        ByVal aSearch As String, ByVal aSrchSpec As E_STRING_SPEC) As String
    Dim wkRtn As String
    
    '�G�X�P�[�v�V�[�P���X�ϊ��i�K�v���̂݁j
    wkRtn = PF_String_ReturnChangeEscSeq(aSearch, Array("\", "{", "}", "(", ")", "[", "]"))
    
    '�J�n���p�^�[���ݒ�
    If M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_POS_START) = True Then
        wkRtn = F_String_ReturnAdd(wkRtn, "^", aAddSpec:=E_STRING_SPEC_POS_START)
    End If
            
    '�I�����p�^�[���ݒ�
    If M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_POS_END) = True Then
        wkRtn = F_String_ReturnAdd(wkRtn, "$", aExcluded:="\$")
    End If
    
    PF_String_ReturnSearchPattern = wkRtn
End Function

' �G�X�P�[�v�V�[�P���X�u��
Private Function PF_String_ReturnChangeEscSeq( _
        ByVal aSearch As String, _
        ByVal aEscSeqAry As Variant) As String
    Dim wkRtn As String
    Dim wkTmp As Variant
    
    '������
    wkRtn = aSearch
    
    '�G�X�P�[�v�V�[�P���X�ɒu������
    For Each wkTmp In aEscSeqAry
        wkRtn = Replace(wkRtn, wkTmp, "\" & wkTmp)
    Next wkTmp
    
    '������
    PF_String_ReturnChangeEscSeq = wkRtn
End Function

'------------------------------------------------------------------------------
' �`�F�b�N�p�^�[���쐬
'------------------------------------------------------------------------------
Private Function PF_String_ReturnCheckPattern( _
        ByVal aSrchPtn As String, ByVal aSrchSpec As E_STRING_SPEC, _
        ByVal aChkPtn As String, ByVal aChkSpec As E_STRING_SPEC) As String
    Dim wkRtn As String
    Dim wkPattern As String
    
    wkRtn = "(" & aSrchPtn & ")"
    
    '��v�p�^�[���ǉ�����̏ꍇ
    If aChkPtn <> "" And aChkSpec <> E_STRING_SPEC_NONE Then
        '�J�n���`�F�b�N�w�肠��A���J�n����v�w��Ȃ��̏ꍇ�A����������J�n���Ƀ`�F�b�N�p�^�[���ǉ�
        If M_Common.F_CheckBitOn(aChkSpec, E_STRING_SPEC_MATCH_START) = True And _
                M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_MATCH_START) <> True Then
            wkRtn = aChkPtn & wkRtn
        End If
        '�I�[���`�F�b�N�w�肠��A���I�[����v�w��Ȃ��̏ꍇ�A����������I�[���Ƀ`�F�b�N�p�^�[���ǉ�
        If M_Common.F_CheckBitOn(aChkSpec, E_STRING_SPEC_MATCH_END) = True And _
                M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_MATCH_END) <> True Then
            wkRtn = wkRtn & aChkPtn
        End If
    End If
    
    '�J�n���p�^�[���ݒ�
    If M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_POS_START) = True Then
        wkRtn = F_String_ReturnAdd(wkRtn, "^", aAddSpec:=E_STRING_SPEC_POS_START)
    End If
            
    '�I�����p�^�[���ݒ�
    If M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_POS_END) = True Then
        wkRtn = F_String_ReturnAdd(wkRtn, "$", aExcluded:="\$")
    End If
    
    PF_String_ReturnCheckPattern = wkRtn
End Function

'------------------------------------------------------------------------------
' �P�ꌟ���p�^�[���쐬
'------------------------------------------------------------------------------
Private Function PF_String_ReturnCheckPatternWord( _
        ByVal aSrchPtn As String, ByVal aSrchSpec As E_STRING_SPEC, _
        ByVal aChkPtn As String, ByVal aChkSpec As E_STRING_SPEC) As String
    Dim wkRtn As String
    Dim wkTmpStr As String
    
    '���Ԍ����w�肠��̏ꍇ�A���������w��
    If M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_MATCH_MID) = True Then
        wkRtn = PF_String_ReturnCheckPattern(aSrchPtn, aSrchSpec, aChkPtn, aChkSpec)
    End If
        
    '�擪�����w��܂��͒��Ԍ����w�肠��A���p�^�[���w��ɏI�[����̏ꍇ
    If M_Common.F_CheckBitOn(aSrchSpec, (E_STRING_SPEC_MATCH_START Or E_STRING_SPEC_MATCH_MID)) = True And _
            M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_MATCH_END) = True Then
        '�J�n�ʒu��v�m�F��ǉ�
        wkTmpStr = PF_String_ReturnCheckPattern(aSrchPtn, E_STRING_SPEC_MATCH_START, aChkPtn, E_STRING_SPEC_MATCH_END)
        wkRtn = F_String_ReturnAdd(wkRtn, wkTmpStr, aDlmt:="|")
    End If
        
    '�I�[�����w��܂��͒��Ԍ����w�肠��A���p�^�[���w��ɐ擪����̏ꍇ
    If M_Common.F_CheckBitOn(aSrchSpec, (E_STRING_SPEC_MATCH_END Or E_STRING_SPEC_MATCH_MID)) = True And _
            M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_MATCH_START) = True Then
        '�I�[�ʒu��v�m�F��ǉ�
        wkTmpStr = PF_String_ReturnCheckPattern(aSrchPtn, E_STRING_SPEC_MATCH_END, aChkPtn, E_STRING_SPEC_MATCH_START)
        wkRtn = F_String_ReturnAdd(wkRtn, wkTmpStr, aDlmt:="|")
    End If
    
    If wkRtn = "" Then
        wkRtn = aSrchPtn
    End If
    
    PF_String_ReturnCheckPatternWord = wkRtn
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �����񒲐�
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �I���ʒu����
'------------------------------------------------------------------------------
Private Function PF_String_GetPosEndAdjust( _
        ByRef aRtn As Long, _
        ByVal aLenMax As Long, _
        ByVal aSttPos As Long, _
        ByVal aLength As Long) As Boolean
    Dim wkRtn As Long: wkRtn = aRtn
    
    '�I���ʒu���J�n�ʒu��菬�����ꍇ
    If wkRtn < aSttPos Then
        '�����w�肠��Ȃ甽�f
        If aLength >= D_POS_START Then
            wkRtn = aSttPos + aLength - 1
        '�����w��Ȃ��Ȃ�ő咷���f
        Else
            wkRtn = aLenMax
        End If
    End If
    If wkRtn > aLenMax Then
        wkRtn = aLenMax
    End If
    
    If wkRtn >= D_POS_START Then
        aRtn = wkRtn
        PF_String_GetPosEndAdjust = True
    End If
End Function

'------------------------------------------------------------------------------
' �����񒷒���
'------------------------------------------------------------------------------
Private Function PF_String_GetLengthAdjust( _
        ByRef aRtn As Long, _
        ByVal aLenMax As Long, _
        ByVal aSttPos As Long, _
        ByVal aEndPos As Long) As Boolean
    Dim wkRtn As Long: wkRtn = aRtn
    
    Dim wkLenMax As Long: wkLenMax = aLenMax - aSttPos + 1
    
    '�����񒷂����͈͊O�̏ꍇ
    If wkRtn < D_POS_START Then
        '�I���w�肠��Ȃ甽�f
        If aEndPos >= aSttPos Then
            wkRtn = aEndPos - aSttPos + 1
        '�I���w��Ȃ��Ȃ�ő咷���f
        Else
            wkRtn = wkLenMax
        End If
    End If
    If wkRtn > wkLenMax Then
        wkRtn = wkLenMax
    End If
    
    If wkRtn >= D_POS_START Then
        aRtn = wkRtn
        PF_String_GetLengthAdjust = True
    End If
End Function
