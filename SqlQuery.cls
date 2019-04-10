VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SqlQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'##################################################################################################
' Class: SqlQuery
' Description:
'   �p�����[�^�[�t��SQL�N�G���̕�����𐶐�����N���X�ł��B
'   ���O�t���o�C���h�ϐ��AIN��ւ̔z��p�����[�^�[��tuple�̔z��p�����[�^�[�}�����ł��܂��B
'
'##################################################################################################

Private s_sqlBody As String
Private dict_params As Object
Private e_dbType As DB_TYPE

' RDBMS�̎��
' ���t�����̃t�H�[�}�b�g�w��q�؂�ւ���ASQL�̕����̐؂�ւ��Ɏg��
' ���̂Ƃ���Oracle�����Ή����ĂȂ�
Public Enum DB_TYPE
    ORACLE
End Enum

' AddParam���\�b�h�œn���p�����[�^�[�l�̌^��\��
' AddParam�̈����Ƃ��Ďg�����肾�������ǁA�p�����[�^�[��tuple�̏ꍇ�ɂǂ����邩�v�����Ă��Ȃ��̂ŁA
' Private�Ŏg���Ă��邾��
Private Enum PARAMETER_TYPE
    VarChar
    Number
    Date
    Unknown
End Enum

'------ �R���X�g���N�^ ------
Private Sub Class_Initialize()
    s_sqlBody = ""
    Set dict_params = CreateObject("Scripting.Dictionary")
    e_dbType = DB_TYPE.ORACLE
End Sub

'------ �v���p�e�B -----
Public Property Let SqlBody(v As String)
    s_sqlBody = v
End Property

Public Property Get SqlBody() As String
    SqlBody = s_sqlBody
End Property

Public Property Get ParamKeys() As Variant()
    ParamKeys = dict_params.Keys()
End Property

Public Property Get Param(key As Variant) As Variant
    Param = dict_params(key)
End Property

Public Property Let DbType(v As DB_TYPE)
    e_dbType = v
End Property

Public Property Get DbType() As DB_TYPE
    DbType = e_dbType
End Property

' SQL�N�G���Ƀo�C���h�ϐ���K�p�������������Ԃ�
Public Property Get SqlQueryString() As String
    Dim sqlStr As String
    sqlStr = s_sqlBody

    ' SQL�N�G�����̃o�C���h�ϐ����AAddParam�ŃZ�b�g�����l�̕�����\���ɒu������
    Dim key As Variant
    For Each key In dict_params.Keys
        sqlStr = bindParamToSql(sqlStr, CStr(key), dict_params(key))
    Next key

    SqlQueryString = sqlStr
End Property

' �o�C���h�ϐ���ǉ�
' bindName�ɂ�SQL�ɖ��ߍ��񂾃v���[�X�z���_�[(:�Ŏn�܂镶����)�̃R��������������������w��(��: where userid = :USERID �Ȃ�"USERID")
Public Sub AddParam(bindName As String, value As Variant)
    Dim v As Variant

    ' Range�I�u�W�F�N�g�̒l�����o��
    ' Range�������͈͂��Q�Ƃ��Ă���ꍇ�́AVariant��2�����z�񂪕Ԃ��Ă���
    If TypeName(value) = "Range" Then
        v = value.value
    Else
        v = value
    End If

    ' value��2�����z�񂾂�����Atuple��\��Variant�z��ɒu������
    If is2DArray(v) Then
        v = MakeTupleParamList(v)
    End If

    ' �o�C���h�ϐ������ǉ��ς̏ꍇ��value���㏑������
    If dict_params.Exists(bindName) Then
        dict_params.Item(bindName) = v
    Else
        dict_params.Add bindName, v
    End If

End Sub

' �o�C���h�ϐ��A�l���N���A
Public Sub CleanParams(bindName As String)
    dict_params.RemoveAll
End Sub


' 2�����z���tuple��\������Variant�z��̃l�X�g�ɕϊ�����
' ��: SQL���tuple parameters��
'   where (col1, col2, col3) in ( (1, 'a', '��'), (2, 'b', '��'), (3, 'c', '��') )
'
'   �̏ꍇ�����Function��
'
'   Array(
'     Array(1, "a", "��"),
'     Array(2, "b", "��"),
'     Array(3, "c", "��")
'   )
'
'   ��Ԃ�
Private Function MakeTupleParamList(v As Variant) As Variant

    Dim i As Long, j As Long
    Dim i_min As Long, i_max As Long, j_min As Long, j_max As Long
    i_min = LBound(v, 1)
    i_max = UBound(v, 1)
    j_min = LBound(v, 2)
    j_max = UBound(v, 2)

    Dim paramList() As Variant
    ReDim paramList(i_min To i_max)
    For i = i_min To i_max
        Dim tuple() As Variant
        ReDim tuple(j_min To j_max)
        For j = j_min To j_max
            tuple(j) = v(i, j)
        Next j
        paramList(i) = tuple
    Next i

    MakeTupleParamList = paramList
End Function

' �����œn���ꂽ�l��2�����z�񂩔ۂ���Ԃ�
Private Function is2DArray(v As Variant) As Boolean

    ' �z��̎��������擾����d�g��
    ' �z��array��i�����̗v�f����UBound(array, i)�Ŏ擾����ہA
    ' i��array�̎������𒴂���ƃG���[�ɂȂ邱�Ƃ𗘗p����
    On Error Resume Next
    Dim tmp As Variant
    Dim i As Long: i = 0
    Do While Err.Number = 0
        i = i + 1
        tmp = UBound(v, i + 1)
    Loop

    is2DArray = (i = 2)

End Function

' SQL���̃v���[�X�z���_�[��l�Œu������
' TODO: SQL�C���W�F�N�V�����΍�
Private Function bindParamToSql(sql As String, bindName As String, value As Variant) As String
    With CreateObject("VBScript.RegExp")
        .Pattern = "\s*:" & bindName & "([\s,\)]?)"
        .IgnoreCase = False
        .Global = True
        bindParamToSql = .Replace(sql, " " & paramValueString(value) & "$1")
    End With
End Function

' �v���[�X�z���_�[��u��������l��\�����镶����𐶐�
' �l���z�񂾂�����A�J���}��؂�ŘA�������������Ԃ�
Private Function paramValueString(value As Variant) As String

    If IsArray(value) Then
        Dim a() As Variant
        ReDim a(LBound(value) To UBound(value))
        Dim i As Long
        For i = LBound(value) To UBound(value)
            a(i) = paramExpressionByType(value(i))
        Next i
        paramValueString = Join(a, ", ")
    Else
        paramValueString = paramExpressionByType(value)
    End If

End Function

' �l�̌^�ɉ����āASQL�ɖ��ߍ��ނׂ�������𐶐�
'  - ���l�^:   �l�����̂܂�
'  - ������^: �V���O���N�H�[�e�[�V�����ň͂�
'  - ���t�^:   ���t�^��Ԃ�SQL�֐��̌`���ŕԂ�(TO_DATE�Ƃ�)
'  - �^�s��:   ������^�Ƃ��Ĉ���(�V���O���N�H�[�e�[�V�����ň͂�)
' �l��tuple�̏ꍇ�̓J�b�R�ł��������������Ԃ�
Private Function paramExpressionByType(value As Variant) As String
    Dim v As Variant
    If Not IsArray(value) Then
        v = value
    ElseIf LBound(value) = UBound(value) Then
        ' �v�f��1�̔z���tuple�ɂ���Ə璷�Ȃ̂�
        ' ('a') -> 'a'
        ' �Ƃ���
        v = value(LBound(value))
    Else
        ' tuple�̏ꍇ�̓J�b�R�ł��������������Ԃ�
        ' tuple���̊e�l��paramValueString�œK�؂ȕ�����\���ɒu��
        paramExpressionByType = "(" & paramValueString(value) & ")"
        Exit Function
    End If

    ' �^�ɉ������������Ԃ�
    Select Case getParameterType(v)
        Case PARAMETER_TYPE.VarChar
            paramExpressionByType = "'" & v & "'"
        Case PARAMETER_TYPE.Number
            paramExpressionByType = v
        Case PARAMETER_TYPE.Date
            paramExpressionByType = getDateTimeFunc(v)
        Case Else
            paramExpressionByType = "'" & v & "'"
    End Select
End Function

' Excel�̒l�^��SQL�̌^�ŕ��ށE����
Private Function getParameterType(value As Variant) As PARAMETER_TYPE
    Select Case VBA.VarType(value)
        Case vbByte, vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDecimal
            getParameterType = PARAMETER_TYPE.Number
#If Win64 Then
        Case vbLongLong
            getParameterType = PARAMETER_TYPE.Number
#End If
        Case vbDate
            getParameterType = PARAMETER_TYPE.Date
        Case vbString
            getParameterType = PARAMETER_TYPE.VarChar
        Case Else
            getParameterType = PARAMETER_TYPE.Unknown
    End Select
End Function

' ���t�^�͕�����->���t�^�ɕϊ�����֐��̌`���ɕϊ�����
Private Function getDateTimeFunc(v As Variant) As String
    Dim datetimeStr As String
    datetimeStr = Format(v, "yyyy-mm-dd hh:nn:ss")

    ' ������->���t�^�ϊ��̊֐���RDBMS�ɂ���ĈقȂ�
    Select Case e_dbType
        Case DB_TYPE.ORACLE
            getDateTimeFunc = "TO_DATE('" & datetimeStr & "', 'YYYY-MM-DD HH24:MI:SS')"
    End Select
End Function