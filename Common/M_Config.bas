Attribute VB_Name = "M_Config"
Option Explicit
'##############################################################################
' �R���t�B�O�ݒ�
'##############################################################################
' ���ʃo�[�W����    |   250504
'------------------------------------------------------------------------------

'==============================================================================
' ���J��`
'==============================================================================
' �萔��`
'------------------------------------------------------------------------------
'�N���X���
Public Enum E_CONFIG_IDX_CLASS_INF
    E_CONFIG_IDX_CLASS_INF_NONE = D_IDX_START - 1
    E_CONFIG_IDX_CLASS_INF_FULLPATH
    E_CONFIG_IDX_CLASS_INF_BOOK
    E_CONFIG_IDX_CLASS_INF_BOOK_NAME
    E_CONFIG_IDX_CLASS_INF_SHEET
    E_CONFIG_IDX_CLASS_INF_SHEET_NAME
    E_CONFIG_IDX_CLASS_INF_RANGE
    E_CONFIG_IDX_CLASS_INF_ARG_TYPE
    E_CONFIG_IDX_CLASS_INF_MAX
    E_CONFIG_IDX_CLASS_INF_EEND = E_CONFIG_IDX_CLASS_INF_MAX - 1
End Enum

'�N���X�w��
Public Enum E_CONFIG_IDX_CLASS
    E_CONFIG_IDX_CLASS_NONE = D_IDX_START - 1
    E_CONFIG_IDX_CLASS_UNIQUE
    E_CONFIG_IDX_CLASS_MAX
    E_CONFIG_IDX_CLASS_EEND = E_CONFIG_IDX_CLASS_MAX - 1
End Enum

' �s�w��
Public Enum E_CONFIG_IDX_CLASS_ARG_TYPE
    E_CONFIG_IDX_CLASS_ARG_TYPE_NONE = D_IDX_START - 1
    E_CONFIG_IDX_CLASS_ARG_TYPE_SHEET
    E_CONFIG_IDX_CLASS_ARG_TYPE_RANGE
    E_CONFIG_IDX_CLASS_ARG_TYPE_MAX
    E_CONFIG_IDX_CLASS_ARG_TYPE_EEND = E_CONFIG_IDX_CLASS_ARG_TYPE_MAX - 1
End Enum

' �s�w��
Public Enum E_CONFIG_ROW
    E_CONFIG_ROW_NONE = D_EXCEL_ROW_START - 1
    E_CONFIG_ROW_MAX
    E_CONFIG_ROW_EEND = E_CONFIG_ROW_MAX - 1
End Enum

'==============================================================================
' ������`
'==============================================================================
' �\���̒�`
'------------------------------------------------------------------------------
Private Type PT_CONFIG_CLASS_INF
    Cls As Object
End Type

Private Type PT_INF
    EventFlg As Boolean
    
    ClsInf(D_IDX_START To E_CONFIG_IDX_CLASS_EEND) As PT_CONFIG_CLASS_INF
End Type

'------------------------------------------------------------------------------
' �ϐ���`
'------------------------------------------------------------------------------
Private pgInf As PT_INF

'==============================================================================
' ���J����
'==============================================================================
' ����������
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Sub S_Config_Init()
    Dim wkCnt As Long
    
    With pgInf
        .EventFlg = False
        
        For wkCnt = D_IDX_START To E_CONFIG_IDX_CLASS_EEND
            With .ClsInf(wkCnt)
                Set .Cls = Nothing
            End With
        Next wkCnt
    End With
End Sub
 
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �N���X����
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �N���X�擾
'------------------------------------------------------------------------------
Public Function F_Config_GetClass( _
        ByRef aRtn As Object, _
        ByRef aClsInf As Variant) As Boolean
    Dim wkRet As Boolean
    
    Dim wkICnt As Long
    
    Dim wkCls As Object
    
    '�C�x���g���͎擾�s�v
    If pgInf.EventFlg = True Then
        Exit Function
    End If
    
    '�N���X���`�F�b�N
    If PF_Config_CheckClsInf(aClsInf) <> True Then
        Exit Function
    End If
    
    For wkICnt = D_IDX_START To E_CONFIG_IDX_CLASS_EEND
        Select Case wkICnt
            Case E_CONFIG_IDX_CLASS_UNIQUE
                Set wkCls = New C_Unique
        End Select
        
        '�������������̓N���X��ێ�
        If wkCls.F_Init(aClsInf) = True Then
            With pgInf.ClsInf(wkICnt)
                Set .Cls = wkCls
                
                Set aRtn = wkCls
                wkRet = True
                Exit For
            End With
        End If
    Next wkICnt
    F_Config_GetClass = wkRet
End Function

' �N���X�������`�F�b�N����
Private Function PF_Config_CheckClsInf( _
        ByRef aRtn As Variant) As Boolean
    Dim wkRtn As Variant: wkRtn = aRtn
    
    Dim wkTmpSh As Worksheet
    
    '�Z���͈͂�����ꍇ
    If Not wkRtn(E_CONFIG_IDX_CLASS_INF_RANGE) Is Nothing Then
        Set wkTmpSh = wkRtn(E_CONFIG_IDX_CLASS_INF_RANGE).Worksheet
        
        Set wkRtn(E_CONFIG_IDX_CLASS_INF_SHEET) = wkTmpSh
        wkRtn(E_CONFIG_IDX_CLASS_INF_ARG_TYPE) = E_CONFIG_IDX_CLASS_ARG_TYPE_RANGE
    Else
        '�V�[�g������ꍇ
        If Not wkRtn(E_CONFIG_IDX_CLASS_INF_SHEET) Is Nothing Then
            Set wkTmpSh = wkRtn(E_CONFIG_IDX_CLASS_INF_SHEET)
            '�������s
        '�V�[�g��������ꍇ
        ElseIf M_Excel.F_Excel_GetSheetName2Sheet(wkTmpSh, _
                                                    wkRtn(E_CONFIG_IDX_CLASS_INF_SHEET_NAME), _
                                                    wkRtn(E_CONFIG_IDX_CLASS_INF_BOOK)) = True Then
            '�������s
            Set wkRtn(E_CONFIG_IDX_CLASS_INF_SHEET) = wkTmpSh
        '�o���Ȃ��ꍇ�ُ͈�I��
        Else
            Exit Function
        End If
        
        wkRtn(E_CONFIG_IDX_CLASS_INF_RANGE) = wkTmpSh.Cells(D_EXCEL_ROW_START, D_EXCEL_CLM_START)
        wkRtn(E_CONFIG_IDX_CLASS_INF_ARG_TYPE) = E_CONFIG_IDX_CLASS_ARG_TYPE_SHEET
    End If
        
    If wkRtn(E_CONFIG_IDX_CLASS_INF_ARG_TYPE) <> E_CONFIG_IDX_CLASS_ARG_TYPE_NONE Then
        With wkTmpSh
            wkRtn(E_CONFIG_IDX_CLASS_INF_SHEET_NAME) = .Name
            
            Set wkRtn(E_CONFIG_IDX_CLASS_INF_BOOK) = .Parent
            With .Parent
                wkRtn(E_CONFIG_IDX_CLASS_INF_FULLPATH) = .FullName
                wkRtn(E_CONFIG_IDX_CLASS_INF_BOOK_NAME) = .Name
            End With
        End With
    
        aRtn = wkRtn
        PF_Config_CheckClsInf = True
    End If
End Function

'------------------------------------------------------------------------------
' �C�x���g�t���O�ݒ�
'------------------------------------------------------------------------------
Public Property Let L_Config_EventFlg( _
            ByVal aFlg As Boolean)
    pgInf.EventFlg = aFlg
End Property

'==============================================================================
' ��������
'==============================================================================
