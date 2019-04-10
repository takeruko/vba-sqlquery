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
'   パラメーター付きSQLクエリの文字列を生成するクラスです。
'   名前付きバインド変数、IN句への配列パラメーターやtupleの配列パラメーター挿入ができます。
'
'##################################################################################################

Private s_sqlBody As String
Private dict_params As Object
Private e_dbType As DB_TYPE

' RDBMSの種類
' 日付時刻のフォーマット指定子切り替えや、SQLの方言の切り替えに使う
' 今のところOracleしか対応してない
Public Enum DB_TYPE
    ORACLE
End Enum

' AddParamメソッドで渡すパラメーター値の型を表す
' AddParamの引数として使うつもりだったけど、パラメーターがtupleの場合にどうするか思いついていないので、
' Privateで使っているだけ
Private Enum PARAMETER_TYPE
    VarChar
    Number
    Date
    Unknown
End Enum

'------ コンストラクタ ------
Private Sub Class_Initialize()
    s_sqlBody = ""
    Set dict_params = CreateObject("Scripting.Dictionary")
    e_dbType = DB_TYPE.ORACLE
End Sub

'------ プロパティ -----
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

' SQLクエリにバインド変数を適用させた文字列を返す
Public Property Get SqlQueryString() As String
    Dim sqlStr As String
    sqlStr = s_sqlBody

    ' SQLクエリ中のバインド変数を、AddParamでセットした値の文字列表現に置換する
    Dim key As Variant
    For Each key In dict_params.Keys
        sqlStr = bindParamToSql(sqlStr, CStr(key), dict_params(key))
    Next key

    SqlQueryString = sqlStr
End Property

' バインド変数を追加
' bindNameにはSQLに埋め込んだプレースホルダー(:で始まる文字列)のコロンを除いた文字列を指定(例: where userid = :USERID なら"USERID")
Public Sub AddParam(bindName As String, value As Variant)
    Dim v As Variant

    ' Rangeオブジェクトの値を取り出す
    ' Rangeが複数範囲を参照している場合は、Variantの2次元配列が返ってくる
    If TypeName(value) = "Range" Then
        v = value.value
    Else
        v = value
    End If

    ' valueが2次元配列だったら、tupleを表すVariant配列に置換する
    If is2DArray(v) Then
        v = MakeTupleParamList(v)
    End If

    ' バインド変数名が追加済の場合はvalueを上書きする
    If dict_params.Exists(bindName) Then
        dict_params.Item(bindName) = v
    Else
        dict_params.Add bindName, v
    End If

End Sub

' バインド変数、値をクリア
Public Sub CleanParams(bindName As String)
    dict_params.RemoveAll
End Sub


' 2次元配列をtupleを表現するVariant配列のネストに変換する
' 例: SQL上のtuple parametersが
'   where (col1, col2, col3) in ( (1, 'a', 'あ'), (2, 'b', 'い'), (3, 'c', 'う') )
'
'   の場合､このFunctionは
'
'   Array(
'     Array(1, "a", "あ"),
'     Array(2, "b", "い"),
'     Array(3, "c", "う")
'   )
'
'   を返す
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

' 引数で渡された値が2次元配列か否かを返す
Private Function is2DArray(v As Variant) As Boolean

    ' 配列の次元数を取得する仕組み
    ' 配列arrayのi次元の要素数をUBound(array, i)で取得する際、
    ' iがarrayの次元数を超えるとエラーになることを利用する
    On Error Resume Next
    Dim tmp As Variant
    Dim i As Long: i = 0
    Do While Err.Number = 0
        i = i + 1
        tmp = UBound(v, i + 1)
    Loop

    is2DArray = (i = 2)

End Function

' SQL中のプレースホルダーを値で置換する
' TODO: SQLインジェクション対策
Private Function bindParamToSql(sql As String, bindName As String, value As Variant) As String
    With CreateObject("VBScript.RegExp")
        .Pattern = "\s*:" & bindName & "([\s,\)]?)"
        .IgnoreCase = False
        .Global = True
        bindParamToSql = .Replace(sql, " " & paramValueString(value) & "$1")
    End With
End Function

' プレースホルダーを置き換える値を表現する文字列を生成
' 値が配列だったら、カンマ区切りで連結した文字列を返す
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

' 値の型に応じて、SQLに埋め込むべき文字列を生成
'  - 数値型:   値をそのまま
'  - 文字列型: シングルクォーテーションで囲む
'  - 日付型:   日付型を返すSQL関数の形式で返す(TO_DATEとか)
'  - 型不明:   文字列型として扱う(シングルクォーテーションで囲む)
' 値がtupleの場合はカッコでくくった文字列を返す
Private Function paramExpressionByType(value As Variant) As String
    Dim v As Variant
    If Not IsArray(value) Then
        v = value
    ElseIf LBound(value) = UBound(value) Then
        ' 要素数1の配列はtupleにすると冗長なので
        ' ('a') -> 'a'
        ' とする
        v = value(LBound(value))
    Else
        ' tupleの場合はカッコでくくった文字列を返す
        ' tuple内の各値はparamValueStringで適切な文字列表現に置換
        paramExpressionByType = "(" & paramValueString(value) & ")"
        Exit Function
    End If

    ' 型に応じた文字列を返す
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

' Excelの値型をSQLの型で分類・判別
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

' 日付型は文字列->日付型に変換する関数の形式に変換する
Private Function getDateTimeFunc(v As Variant) As String
    Dim datetimeStr As String
    datetimeStr = Format(v, "yyyy-mm-dd hh:nn:ss")

    ' 文字列->日付型変換の関数はRDBMSによって異なる
    Select Case e_dbType
        Case DB_TYPE.ORACLE
            getDateTimeFunc = "TO_DATE('" & datetimeStr & "', 'YYYY-MM-DD HH24:MI:SS')"
    End Select
End Function
