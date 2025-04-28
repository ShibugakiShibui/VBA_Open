Attribute VB_Name = "M_String"
Option Explicit
'##############################################################################
' �����񏈗�
'##############################################################################
' �Q�Ɛݒ�          |   Microsoft VBScript Regular Expressions 5.5
'------------------------------------------------------------------------------
' �Q�ƃ��W���[��    |   �\
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
    E_STRING_SPEC_POS_MATCH = E_STRING_SPEC_POS_START Or E_STRING_SPEC_POS_END
    E_STRING_SPEC_POS_FULL = E_STRING_SPEC_POS_START Or E_STRING_SPEC_POS_END Or E_STRING_SPEC_POS_MID
    E_STRING_SPEC_WORD_NOTWORD = &H10
    E_STRING_SPEC_WORD_WORD = &H20
    E_STRING_SPEC_WORD_MASK = E_STRING_SPEC_WORD_NOTWORD Or E_STRING_SPEC_WORD_WORD
End Enum

Public Enum E_STRING_IDX_SRCH_INF
    E_STRING_IDX_SRCH_INF_NONE = D_IDX_START - 1
    E_STRING_IDX_SRCH_INF_POS_START
    E_STRING_IDX_SRCH_INF_LENGTH
    E_STRING_IDX_SRCH_INF_MAX
    E_STRING_IDX_SRCH_INF_EEND = E_STRING_IDX_SRCH_INF_MAX - 1
End Enum

Public Const D_STRING_CHECKWORD As String = "A-Za-z0-9_"

'------------------------------------------------------------------------------
' �\���̒�`
'------------------------------------------------------------------------------
Public Type T_STRING_ARG_ADD_INF
    '������
    Str As String
    
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

Public Type T_STRING_ARG_CHK_INF
    '������
    Str As String
    
    '����������
    Search As String
    SrchSpec As E_STRING_SPEC
    SrchPtn As String
    
    '�`�F�b�N�p�^�[��
    ChkPtn As String
    ChkPtnSpec As E_STRING_SPEC
    ChkPtnOfs As Long
    ChkWordFlg As Boolean
    
    '�����ʒu�w��
    SttPos As Long
    EndPos As Long
    Length As Long
End Type

Public Type T_STRING_ARG_GET_INF
    ChkInf As T_STRING_ARG_CHK_INF
    
    SttStr As String
    EndStr As String
    
    AddBefFlg As Boolean
    AddSrchFlg As Boolean
End Type

Public Type T_STRING_ARG_DEL_INF
    ChkInf As T_STRING_ARG_CHK_INF
    
    AddDelFlg As Boolean
    DelSpec As E_STRING_SPEC
    DelPosCnt As Long
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
    With G_String_InitArgAddInf
        .Str = ""
        
        .Dlmt = ""
        .DlmtChkFlg = True
        
        .Add = ""
        .AddChkFlg = False
        .AddSpec = E_STRING_SPEC_POS_END
        
        .Excluded = ""
    End With
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
        wkRtn = .Str
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
        ByVal aStr As String, ByVal aAdd As String, _
        Optional ByVal aDlmt As String = "", _
        Optional ByVal aAddSpec As E_STRING_SPEC = E_STRING_SPEC_POS_END, _
        Optional ByVal aExcluded As String = "") As String
    Dim wkArgInf As T_STRING_ARG_ADD_INF: wkArgInf = G_String_InitArgAddInf()
    
    With wkArgInf
        .Str = aStr
        .Dlmt = aDlmt
        .Add = aAdd
        .AddSpec = aAddSpec
        .Excluded = aExcluded
    End With
    
    F_String_ReturnAdd = F_String_ReturnAdd_Inf(wkArgInf)
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �����񌟍����z��擾
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ������񏉊���
'------------------------------------------------------------------------------
Public Property Get G_String_InitArgChkInf() As T_STRING_ARG_CHK_INF
    With G_String_InitArgChkInf
        .Str = ""
        
        .Search = ""
        .SrchSpec = E_STRING_SPEC_POS_MID
        .SrchPtn = ""
        
        .ChkPtn = ""
        .ChkPtnSpec = E_STRING_SPEC_POS_MATCH
        .ChkPtnOfs = 1
        .ChkWordFlg = False
        
        .SttPos = D_POS_START
        .EndPos = D_POS_END
        .Length = D_POS_END
    End With
End Property

'------------------------------------------------------------------------------
' �����񌟍����z��擾�i�������w��j
'------------------------------------------------------------------------------
Public Function F_String_GetSearchInfArray_Inf( _
        ByRef aRtnAryAry As Variant, _
        ByRef aArgInf As T_STRING_ARG_CHK_INF) As Boolean
    Dim wkRtnAryAry As Variant
    Dim wkRtnAry(D_IDX_START To E_STRING_IDX_SRCH_INF_EEND) As Variant
    Dim wkAddFlg As Boolean
    
    Dim wkArgInf As T_STRING_ARG_CHK_INF: wkArgInf = aArgInf
    Dim wkSttPos As Long, wkEndPos As Long, wkMaxPos As Long
    
    Dim wkSrchPtn As String
    Dim wkSrchRegExp As RegExp
    Dim wkSrchMatch As Variant
    Dim wkSrchCnt As Long
    
    Dim wkChkPtn As String
    Dim wkChkRegExp As RegExp
    Dim wkChkStt As Long, wkChkEnd As Long
    
    With aArgInf
        '�����`�F�b�N
        If .Str = "" Or (.Search = "" And .SrchPtn = "") Or .SttPos < D_POS_START Then
            Exit Function
        End If
        
        '�I�[�ʒu����
        wkSttPos = .SttPos
        wkEndPos = .EndPos
        wkMaxPos = Len(.Str)
        If PF_String_GetPosEndAdjust(wkEndPos, Len(.Str), wkSttPos, .Length) <> True Then
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
        Set wkSrchMatch = wkSrchRegExp.Execute(Mid(.Str, wkSttPos, (wkEndPos - wkSttPos + 1)))
        '�������ʂ�������ΏI��
        If wkSrchMatch.Count <= 0 Then
            Exit Function
        End If
        
        '�`�F�b�N�p�^�[��������ꍇ
        wkChkPtn = .ChkPtn
        If .ChkWordFlg = True And wkChkPtn = "" Then
            wkChkPtn = D_STRING_CHECKWORD
        End If
        If wkChkPtn <> "" And .ChkPtnOfs > 0 Then
            '�`�F�b�N�p�^�[���w�肪����ꍇ
            If M_Common.F_CheckBitOn(.ChkPtnSpec, E_STRING_SPEC_POS_MATCH) = True Then
                '�`�F�b�N�p�^�[���쐬
                If .ChkWordFlg <> True Then
                    wkChkPtn = PF_String_ReturnCheckPattern(wkSrchPtn, .SrchSpec, wkChkPtn, .ChkPtnSpec)
                Else
                    wkChkPtn = PF_String_ReturnCheckPatternWord(wkSrchPtn, .SrchSpec, wkChkPtn, .ChkPtnSpec)
                End If
            '�`�F�b�N�p�^�[���w�肪�Ȃ��ꍇ
            Else
                '�`�F�b�N�p�^�[�������̂܂ܐݒ�
                wkChkPtn = .ChkPtn
            End If
            
            '�`�F�b�N�����ݒ�
            Set wkChkRegExp = New RegExp
            With wkChkRegExp
                .IgnoreCase = False
                .Global = False
                .Pattern = wkChkPtn
            End With
        End If
    End With
        
    '�����q�b�g�����A�ʒu���𒊏o
    For wkSrchCnt = 0 To wkSrchMatch.Count - 1
        '������
        wkAddFlg = True
        wkRtnAry(E_STRING_IDX_SRCH_INF_POS_START) = wkSttPos + wkSrchMatch.Item(wkSrchCnt).FirstIndex
        wkRtnAry(E_STRING_IDX_SRCH_INF_LENGTH) = wkSrchMatch.Item(wkSrchCnt).Length
            
        If Not wkChkRegExp Is Nothing Then
            '�`�F�b�N�ʒu����
            wkChkStt = wkRtnAry(E_STRING_IDX_SRCH_INF_POS_START)
            wkChkEnd = wkChkStt + wkRtnAry(E_STRING_IDX_SRCH_INF_LENGTH) - 1

            '�J�n�ʒu����
            If wkChkStt > 1 Then
                wkChkStt = wkChkStt - aArgInf.ChkPtnOfs
            End If
            '�I���ʒu����
            If (wkChkEnd + aArgInf.ChkPtnOfs) <= wkMaxPos Then
                wkChkEnd = wkChkEnd + aArgInf.ChkPtnOfs
            End If
                
            wkAddFlg = wkChkRegExp.Test(Mid(aArgInf.Str, wkChkStt, (wkChkEnd - wkChkStt + 1)))
        End If
            
        If wkAddFlg = True Then
            '�擾���ʂ�z��ɓo�^
            wkRtnAryAry = M_Common.F_ReturnArrayAdd(wkRtnAryAry, wkRtnAry)
        End If
    Next wkSrchCnt
    
    '�w��p�^�[�������������ꍇ
    If IsArray(wkRtnAryAry) = True Then
        aRtnAryAry = wkRtnAryAry
        F_String_GetSearchInfArray_Inf = True
    End If
    
    Set wkSrchRegExp = Nothing
    Set wkSrchMatch = Nothing
    Set wkChkRegExp = Nothing
End Function

'------------------------------------------------------------------------------
' �����񌟍����z��擾�i�����w��j
'------------------------------------------------------------------------------
Public Function F_String_GetSearchInfArray( _
        ByRef aRtnAryAry As Variant, _
        ByVal aStr As String, ByVal aSearch As String, _
        Optional ByVal aSrchPtnFlg As Boolean = False) As Boolean
    Dim wkArgInf As T_STRING_ARG_CHK_INF: wkArgInf = G_String_InitArgChkInf()
    
    With wkArgInf
        .Str = aStr
        If aSrchPtnFlg = True Then
            .SrchPtn = aSearch
        Else
            .Search = aSearch
        End If
    End With
    
    F_String_GetSearchInfArray = F_String_GetSearchInfArray_Inf(aRtnAryAry, wkArgInf)
End Function

'------------------------------------------------------------------------------
' �P�ꌟ�����z��擾�i�����w��j
'------------------------------------------------------------------------------
Public Function F_String_GetSearchInfArrayWord( _
        ByRef aRtnAryAry As Variant, _
        ByVal aStr As String, ByVal aSearch As String, _
        Optional ByVal aChkPtn As String = D_STRING_CHECKWORD, _
        Optional ByVal aSrchSpec As E_STRING_SPEC = E_STRING_SPEC_POS_MID, _
        Optional ByVal aChkPtnSpec As E_STRING_SPEC = E_STRING_SPEC_POS_MATCH) As Boolean
    Dim wkArgInf As T_STRING_ARG_CHK_INF: wkArgInf = G_String_InitArgChkInf()
    
    '�������ɐݒ�
    With wkArgInf
        .Str = aStr
        .Search = aSearch
        .SrchSpec = aSrchSpec
        
        .ChkWordFlg = True
        .ChkPtn = aChkPtn
        .ChkPtnSpec = aChkPtnSpec
    End With
    
    F_String_GetSearchInfArrayWord = F_String_GetSearchInfArray_Inf(aRtnAryAry, wkArgInf)
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �����񑶍݃`�F�b�N
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �����񑶍݃`�F�b�N�i�������w��j
'------------------------------------------------------------------------------
Public Function F_String_Check_Inf( _
        ByRef aArgInf As T_STRING_ARG_CHK_INF) As Boolean
    Dim wkRtn As Boolean
    
    Dim wkTmpAryAry As Variant
    Dim wkTmpPos As Long
    
    If F_String_GetSearchInfArray_Inf(wkTmpAryAry, aArgInf) <> True Then
        Exit Function
    End If
    
    With aArgInf
        '�擪��v�w��A���̏ꍇ
        If M_Common.F_CheckBitOn(.SrchSpec, E_STRING_SPEC_POS_START) = True Then
            wkTmpPos = LBound(wkTmpAryAry)
            
            '�擪�łȂ���ΏI��
            If wkTmpAryAry(wkTmpPos)(E_STRING_IDX_SRCH_INF_POS_START) > 1 Then
                Exit Function
            End If
        End If
        '�I�[��v�w��A���̏ꍇ
        If M_Common.F_CheckBitOn(.SrchSpec, E_STRING_SPEC_POS_END) = True Then
            wkTmpPos = UBound(wkTmpAryAry)
            
            '�I�[�łȂ���ΏI��
            If (wkTmpAryAry(wkTmpPos)(E_STRING_IDX_SRCH_INF_POS_START) + wkTmpAryAry(wkTmpPos)(E_STRING_IDX_SRCH_INF_LENGTH) - 1) < _
                    Len(.Str) Then
                Exit Function
            End If
        End If
    End With
    
    F_String_Check_Inf = True
End Function

'------------------------------------------------------------------------------
' �����񑶍݃`�F�b�N�i�����w��j
'------------------------------------------------------------------------------
Public Function F_String_Check( _
        ByVal aStr As String, ByVal aSearch As String, _
        Optional ByVal aSrchSpec As E_STRING_SPEC = E_STRING_SPEC_POS_MID) As Boolean
    Dim wkArgInf As T_STRING_ARG_CHK_INF: wkArgInf = G_String_InitArgChkInf()
    
    With wkArgInf
        .Str = aStr
        .Search = aSearch
        .SrchSpec = aSrchSpec
    End With
    
    F_String_Check = F_String_Check_Inf(wkArgInf)
End Function

'------------------------------------------------------------------------------
' �P�ꑶ�݃`�F�b�N�i�����w��j
'------------------------------------------------------------------------------
Public Function F_String_CheckWord( _
        ByVal aStr As String, ByVal aSearch As String, _
        Optional ByVal aChkPtn As String = D_STRING_CHECKWORD, _
        Optional ByVal aSrchSpec As E_STRING_SPEC = E_STRING_SPEC_POS_MID, _
        Optional ByVal aChkPtnSpec As E_STRING_SPEC = E_STRING_SPEC_POS_MATCH) As Boolean
    Dim wkRtn As Boolean
    
    Dim wkArgInf As T_STRING_ARG_CHK_INF: wkArgInf = G_String_InitArgChkInf()
    Dim wkTmpAryAry As Variant
    
    With wkArgInf
        .Str = aStr
        .Search = aSearch
        .SrchSpec = aSrchSpec
        
        .ChkPtn = aChkPtn
        .ChkPtnSpec = aChkPtnSpec
    End With
    
    F_String_CheckWord = F_String_Check_Inf(wkArgInf)
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �����񕪊�
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_String_GetSplit( _
        ByRef aRtnAry As Variant, _
        ByVal aStr As String, ByVal aDlmt As String, _
        Optional ByVal aIncChkFlg As Boolean = False) As Boolean
    Dim wkRtnAry As Variant
    
    '�����`�F�b�N
    If aStr = "" Or aDlmt = "" Then
        Exit Function
    End If
    
    '��؂�ŕ���
    wkRtnAry = Split(aStr, aDlmt)
    
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
    If F_String_GetSplit(wkTmpAry, aExtSpec, ";") <> True Then
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
    With G_String_InitArgGetInf
        .ChkInf = G_String_InitArgChkInf
        
        .SttStr = ""
        .EndStr = ""
        
        .AddBefFlg = False
        .AddSrchFlg = False
    End With
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
    Dim wkArgChkInf As T_STRING_ARG_CHK_INF
    
    Dim wkSttCnt As Long, wkSttNow As Long
    Dim wkChkCnt As Long
    
    Dim wkGetSttInf As Variant
    Dim wkGetEndInf As Variant
    Dim wkGetSttPos As Long, wkGetEndPos As String
    
    With aArgInf
        wkArgChkInf = aArgInf.ChkInf
        
        wkArgChkInf.Search = .SttStr
        '�J�n�ʒu�擾�i�����`�F�b�N���p�j
        If F_String_GetSearchInfArray_Inf(wkSttInfAry, wkArgChkInf) <> True Then
            Exit Function
        End If
        
        wkArgChkInf.Search = .EndStr
        '�I���ʒu�擾�i�����`�F�b�N���p�j
        If F_String_GetSearchInfArray_Inf(wkEndInfAry, wkArgChkInf) <> True Then
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
                wkGetSttPos = .ChkInf.SttPos
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
                wkRtn = Mid(.ChkInf.Str, wkGetSttPos, (wkGetEndPos - wkGetSttPos + 1))
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
        ByVal aStr As String, _
        ByVal aSttStr As String, ByVal aEndStr As String, _
        Optional ByVal aAddBefFlg As Boolean = False, _
        Optional ByVal aAddSrchFlg As Boolean = False) As Boolean
    Dim wkArgInf As T_STRING_ARG_GET_INF: wkArgInf = G_String_InitArgGetInf()
    
    With wkArgInf
        .ChkInf.Str = aStr
        
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
        ByVal aStr As String, _
        ByVal aDelete As String, ByVal aDelSpec As E_STRING_SPEC) As String
    Dim wkRtn As String: wkRtn = aStr
    Dim wkLen As Long
    Dim wkDelLen As Long
    
    '������ƍ폜����������ꍇ
    If aStr <> "" And aDelete <> "" Then
        '���Ԉʒu�폜�̏ꍇ
        If M_Common.F_CheckBitOn(aDelSpec, E_STRING_SPEC_POS_MID) = True Then
            wkRtn = Replace(aStr, aDelete, "")
        Else
            wkLen = Len(aStr)
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
    With G_String_InitArgDelInf
        .ChkInf = G_String_InitArgChkInf
        
        .DelSpec = E_STRING_SPEC_POS_START
        .DelPosCnt = D_IDX_START
    End With
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
        wkRtn = .ChkInf.Str
        
        '����������쐬
        If F_String_GetSearchInfArray_Inf(wkInfAryAry, .ChkInf) <> True Then
            '����������������Ȃ������ꍇ�͖���
        ElseIf .DelPosCnt > UBound(wkInfAryAry) Then
            '�폜�w��J�E���^���z��𒴂��Ă���ꍇ�͖���
        Else
            '�폜�w��J�E���^����
            If .DelPosCnt < D_IDX_START Then
                .DelPosCnt = UBound(wkInfAryAry)
            End If
            wkStrLen = Len(.ChkInf.Str)
            
            '�J�n�`������ʒu�܂ō폜
            If M_Common.F_CheckBitOn(.DelSpec, E_STRING_SPEC_POS_START) = True Then
                '�폜�J�n�ʒu��ݒ�
                wkDelStt = .ChkInf.SttPos
                
                '�폜�I���ʒu��ݒ�
                wkDelEnd = wkInfAryAry(.DelPosCnt)(E_STRING_IDX_SRCH_INF_POS_START)
                '�폜������ǉ��w�肠��̏ꍇ
                If .AddDelFlg = True Then
                    wkDelEnd = wkDelEnd - 1
                '�폜������ǉ��w��Ȃ��̏ꍇ
                Else
                    wkDelEnd = wkDelEnd + wkInfAryAry(.DelPosCnt)(E_STRING_IDX_SRCH_INF_LENGTH) - 1
                End If
            '������ʒu�`�I���܂ō폜
            Else
                '�폜�J�n�ʒu��ݒ�
                wkDelStt = wkInfAryAry(.DelPosCnt)(E_STRING_IDX_SRCH_INF_POS_START)
                '�폜������ǉ��w�肠��̏ꍇ
                If .AddDelFlg = True And wkDelStt > D_POS_START Then
                    wkDelStt = wkDelStt + wkInfAryAry(.DelPosCnt)(E_STRING_IDX_SRCH_INF_LENGTH)
                End If
                
                '�폜�I���ʒu��ݒ�
                wkDelEnd = .ChkInf.EndPos
                If wkDelEnd < D_POS_START Then
                    PF_String_GetPosEndAdjust wkDelEnd, wkStrLen, .ChkInf.SttPos, .ChkInf.Length
                End If
            End If
            
            '�폜�ʒu�ɖ��Ȃ��ꍇ�A�폜���{
            If wkDelStt <= wkDelEnd Then
                wkRtn = ""
                If wkDelStt > 1 Then
                    wkRtn = Left(.ChkInf.Str, wkDelStt - 1)
                End If
                If wkDelEnd < wkStrLen Then
                    wkRtn = wkRtn & Right(.ChkInf.Str, (wkStrLen - wkDelEnd))
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
        ByVal aStr As String, _
        ByVal aDelete As String, _
        Optional ByVal aDelSpec As E_STRING_SPEC = E_STRING_SPEC_POS_START, _
        Optional ByVal aDelPosCnt As Long = D_IDX_START) As String
    Dim wkArgInf As T_STRING_ARG_DEL_INF: wkArgInf = G_String_InitArgDelInf
    
    With wkArgInf
        .ChkInf.Str = aStr
        .ChkInf.Search = aDelete
        .DelSpec = aDelSpec
        .DelPosCnt = aDelPosCnt
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
        ByVal aChkPtn As String, ByVal aChkPtnSpec As E_STRING_SPEC) As String
    Dim wkRtn As String
    Dim wkPattern As String
    
    wkRtn = "(" & aSrchPtn & ")"
    
    '�p�^�[���Z�b�g
    If aChkPtnSpec <> E_STRING_SPEC_NONE And aChkPtn <> "" Then
        If M_Common.F_CheckBitOn(aChkPtnSpec, E_STRING_SPEC_WORD_NOTWORD) = True Then
            wkPattern = "[^" & aChkPtn & "]+"
        Else
            wkPattern = "[" & aChkPtn & "]+"
        End If
        '�J�n���`�F�b�N�w�肠�肩�J�n�������w��Ȃ��̏ꍇ
        If M_Common.F_CheckBitOn(aChkPtnSpec, E_STRING_SPEC_POS_START) = True And _
                M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_POS_START) <> True Then
            wkRtn = wkPattern & wkRtn
        End If
        '�I�����`�F�b�N�w�肠��
        If M_Common.F_CheckBitOn(aChkPtnSpec, E_STRING_SPEC_POS_END) = True And _
                M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_POS_END) <> True Then
            wkRtn = wkRtn & wkPattern
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
        ByVal aChkPtn As String, ByVal aChkPtnSpec As E_STRING_SPEC) As String
    Dim wkRtn As String
    Dim wkTmpStr As String
    
    '���Ԍ����w�肠��̏ꍇ�A���������w��
    If M_Common.F_CheckBitOn(aSrchSpec, E_STRING_SPEC_POS_MID) = True Then
        wkRtn = PF_String_ReturnCheckPattern(aSrchPtn, aSrchSpec, aChkPtn, (aChkPtnSpec Or E_STRING_SPEC_WORD_NOTWORD))
    End If
    
    '�擪�����w��܂��͒��Ԍ����w�肠��A���p�^�[���w��ɏI�[����̏ꍇ
    If M_Common.F_CheckBitOn(aSrchSpec, (E_STRING_SPEC_POS_START Or E_STRING_SPEC_POS_MID)) = True And _
            M_Common.F_CheckBitOn(aChkPtnSpec, E_STRING_SPEC_POS_END) = True Then
        '�I�[��v�m�F��ǉ�
        wkTmpStr = PF_String_ReturnCheckPattern(aSrchPtn, E_STRING_SPEC_POS_START, aChkPtn, (E_STRING_SPEC_POS_END Or E_STRING_SPEC_WORD_NOTWORD))
        wkRtn = F_String_ReturnAdd(wkRtn, wkTmpStr, aDlmt:="|")
    End If
    
    '�I�[�����w��܂��͒��Ԍ����w�肠��A���p�^�[���w��ɐ擪����̏ꍇ
    If M_Common.F_CheckBitOn(aSrchSpec, (E_STRING_SPEC_POS_END Or E_STRING_SPEC_POS_MID)) = True And _
            M_Common.F_CheckBitOn(aChkPtnSpec, E_STRING_SPEC_POS_START) = True Then
        '�I�[��v�m�F��ǉ�
        wkTmpStr = PF_String_ReturnCheckPattern(aSrchPtn, E_STRING_SPEC_POS_END, aChkPtn, (E_STRING_SPEC_POS_START Or E_STRING_SPEC_WORD_NOTWORD))
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
