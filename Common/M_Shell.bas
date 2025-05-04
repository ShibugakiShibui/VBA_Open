Attribute VB_Name = "M_Shell"
Option Explicit
'##############################################################################
' �O���v���O��������
'##############################################################################
' �Q�Ɛݒ�          |   Windows Script Host Object Model
'------------------------------------------------------------------------------
' �Q�ƃ��W���[��    |   M_String
'------------------------------------------------------------------------------
' ���ʃo�[�W����    |   �\
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
' ���s���ʎ擾
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �ꊇ�ԋp
'------------------------------------------------------------------------------
Public Function F_Shell_GetResult( _
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
    Set wkExec = wkShell.Exec(aCmd)
    
    wkRtn = wkExec.StdOut.ReadAll
    On Error GoTo 0

    If wkRtn <> "" Then
        aRtn = wkRtn
        wkRet = True
    End If

PROC_ERROR:
    F_Shell_GetResult = wkRet
End Function

'------------------------------------------------------------------------------
' �z��ԋp
'------------------------------------------------------------------------------
Public Function F_Shell_GetResultArray( _
        ByRef aRtnAry As Variant, _
        ByVal aCmd As String) As Boolean
    Dim wkRtnAry As Variant
    Dim wkRtnStr As String
    
    Dim wkTmpAry As Variant, wkTmp As Variant
    Dim wkTmpAry2 As Variant, wkTmp2 As Variant
    
    '�R�}���h�̌��ʂ�NG�ł���ΏI��
    If F_Shell_GetResult(wkRtnStr, aCmd) <> True Then
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
        F_Shell_GetResultArray = True
    End If
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' ���s
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_Shell_Run( _
        ByVal aCmd As String, _
        Optional ByVal aWindowStyle As Long = 0, _
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
    wkCmdRet = wkShell.Run(aCmd, WindowStyle:=aWindowStyle, WaitOnReturn:=aWaitOnReturn)
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
    F_Shell_Run = wkRet
End Function

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
' �w��v���O�������s
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Public Function F_Shell_OpenFile( _
        ByVal aFilePath As String, _
        ByVal aExePath As String, _
        Optional ByVal aExeArg As String = "", _
        Optional ByVal aWindowStyle As VbAppWinStyle = vbNormalFocus) As Boolean
    Dim wkRet As Boolean: wkRet = False
    Dim wkCmd As String
    
    '�����`�F�b�N
    If Dir(aFilePath) = "" Or Dir(aExePath) = "" Then
        Exit Function
    End If
    
    wkCmd = aExePath
    wkCmd = M_String.F_String_ReturnAdd(wkCmd, aExeArg, aDlmt:=" ")
    wkCmd = M_String.F_String_ReturnAdd(wkCmd, aFilePath, aDlmt:=" ")
    
    Shell wkCmd, aWindowStyle
    
    F_Shell_OpenFile = True
End Function
