Attribute VB_Name = "M_Excel"
Option Explicit
'##############################################################################
' Excel����
'##############################################################################
' �Q�Ɛݒ�          |   �\
'------------------------------------------------------------------------------
' ���ʃo�[�W����    |   250504
'------------------------------------------------------------------------------
' �ʃo�[�W����    |   �\
'------------------------------------------------------------------------------
' �捞����          |   �t�@�C���ꗗ�쐬_250504
'------------------------------------------------------------------------------

'==============================================================================
' ���J��`
'==============================================================================
' �萔��`
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
' �\���̒�`
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
' ������`
'==============================================================================

'==============================================================================
' ���J����
'==============================================================================
' ���[�N�V�[�g�֘A����
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ���[�N�V�[�g�������[�N�V�[�g�ϊ�
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
' �S�\��
'------------------------------------------------------------------------------
Public Sub S_Excel_ShowAll( _
        ByVal aSh As Worksheet)
    If aSh Is Nothing Then
        Exit Sub
    End If
    
    With aSh
        .Cells.Rows.Hidden = False
        .Cells.Columns.Hidden = False
        
        '�I�[�g�t�B���^���ݒ肳��Ă���ꍇ�͑S�\��
        If .AutoFilterMode <> True Then
            '�t�B���^�ݒ薳���̏ꍇ�͖���
        ElseIf .FilterMode <> True Then
            '�i�荞�݂���Ă��Ȃ��ꍇ�͒�
        Else
            .ShowAllData
        End If
    End With
End Sub

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �Z���֘A����
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �ʒu��񏉊���
'------------------------------------------------------------------------------
' �s�ʒu������
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

' ��ʒu������
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
' �Z���͈͒l�擾
'------------------------------------------------------------------------------
Public Function F_Excel_ReturnRangeValueArray( _
        ByVal aRg As Range) As Variant
    Dim wkRtnAry As Variant
    
    If aRg Is Nothing Then
        Exit Function
    End If
    
    wkRtnAry = aRg.Value
    '�z��łȂ��ꍇ
    If IsArray(wkRtnAry) <> True Then
        ReDim wkRtnAry(D_EXCEL_ROW_START To D_EXCEL_CLM_START)
        wkRtnAry(D_EXCEL_ROW_START, D_EXCEL_CLM_START) = aRg.Value
    End If
    
    F_Excel_ReturnRangeValueArray = wkRtnAry
End Function

'------------------------------------------------------------------------------
' �I�[�g�t�B���^�ݒ�
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
    
    '�I�[�g�t�B���^���ݒ肳��Ă���ꍇ�͈�U����
    If wkSh.AutoFilterMode = True Then
        wkSh.Cells.AutoFilter
    End If
    
    '�I�[�g�t�B���^�ݒ�
    wkRg.AutoFilter
End Sub

'==============================================================================
' ��������
'==============================================================================
