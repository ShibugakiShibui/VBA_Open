Attribute VB_Name = "M_List"
Option Explicit
'##############################################################################
' ���X�g����
'##############################################################################
' �Q�Ɛݒ�          |   Microsoft Scripting Runtime
'------------------------------------------------------------------------------
' �Q�ƃ��W���[��    |   �\
'------------------------------------------------------------------------------
' ���ʃo�[�W����    |   250427
'------------------------------------------------------------------------------

'==============================================================================
' ���J��`
'==============================================================================

'==============================================================================
' ������`
'==============================================================================

'==============================================================================
' ���J����
'==============================================================================
' ���X�g�ǉ�
'------------------------------------------------------------------------------
Public Function F_List_GetAdd( _
        ByRef aRtn As Dictionary, _
        ByVal aKey As Variant, _
        ByVal aItem As Variant, _
        Optional ByVal aUpdtFlg As Boolean = True) As Boolean
    Dim wkRtn As Dictionary: Set wkRtn = aRtn
    
    '�����`�F�b�N
    Select Case VarType(aKey)
        Case vbEmpty, vbNull
            '��L�ȊO�ɂ����邪���������w�肵�Ȃ��̂Ŗ��Ȃ�
            Exit Function
        Case Else
            '���Ȃ�
    End Select
    
    '���X�g����
    If wkRtn Is Nothing Then
        Set wkRtn = New Dictionary
    End If
    
    With wkRtn
        '�L�[�����݂��Ȃ��ꍇ
        If .Exists(aKey) <> True Then
            .Add aKey, aItem
        '�L�[�����݂���ꍇ
        Else
            '�X�V���Ȃ�Ώ㏑��
            If aUpdtFlg = True Then
                M_Common.S_SetItem .Item(aKey), aItem
            Else
                Exit Function
            End If
        End If
    End With
    
    Set aRtn = wkRtn
    F_List_GetAdd = True
End Function
