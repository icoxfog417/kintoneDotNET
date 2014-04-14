Imports System.Linq.Expressions

Namespace API

    ''' <summary>
    ''' kintoneのクエリを作成するための中間オブジェクト
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneQuery
        Public Property App As String = ""
        Private _query As String = ""
        Private _offset As Integer = -1
        Private _limit As Integer = -1
        Private _orderBy As New Dictionary(Of String, Boolean)
        Private _fields As New List(Of String)

        Public Sub New()
        End Sub

        Public Sub New(ByVal query As String)
            Me._query = query
        End Sub

        Public Sub New(ByVal app As String, ByVal query As String)
            Me.App = app
            Me._query = query
        End Sub

        Public Shared Function Make() As kintoneQuery
            Return New kintoneQuery()
        End Function

        Public Shared Function Make(ByVal query As String) As kintoneQuery
            Return New kintoneQuery(query)
        End Function

        Public Shared Function Make(ByVal app As String, ByVal query As String) As kintoneQuery
            Return New kintoneQuery(app, query)
        End Function

        Public Function Clone() As kintoneQuery
            Dim c As New kintoneQuery(App, _query)
            c._offset = Me._offset
            c._limit = Me._limit
            c._orderBy = New Dictionary(Of String, Boolean)(Me._orderBy)
            c._fields = New List(Of String)(Me._fields)
            Return c

        End Function


        ''' <summary>
        ''' 抽出条件の作成を行う
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="expression"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Where(Of T As AbskintoneModel)(ByVal expression As Expression(Of Func(Of T, Boolean))) As kintoneQuery
            _query = kintoneQueryExpression.Eval(expression)
            Return Me
        End Function

        ''' <summary>
        ''' 抽出条件の作成を行う
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="expression"></param>
        ''' <param name="nameConvertor"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Where(Of T As AbskintoneModel)(ByVal expression As Expression(Of Func(Of T, Boolean)), _
                                                             ByVal nameConvertor As Dictionary(Of String, String)) As kintoneQuery
            _query = kintoneQueryExpression.Eval(expression, nameConvertor)
            Return Me
        End Function

        ''' <summary>
        ''' レコードの読み出し位置(Offset)の指定を行う
        ''' </summary>
        ''' <param name="index"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Offset(ByVal index As Integer) As kintoneQuery
            _offset = index
            Return Me
        End Function

        ''' <summary>
        ''' 抽出するレコード件数(Limit)の指定を行う
        ''' </summary>
        ''' <param name="size"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Limit(ByVal size As Integer) As kintoneQuery
            _limit = size
            Return Me
        End Function

        ''' <summary>
        ''' 抽出対象のフィールドを指定する
        ''' </summary>
        ''' <param name="names"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Fields(ParamArray names As String()) As kintoneQuery
            For Each n As String In names
                _fields.Add(n)
            Next
            Return Me
        End Function

        ''' <summary>
        ''' 昇順のオーダー設定を行う
        ''' </summary>
        ''' <param name="names"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Ascending(ParamArray names As String()) As kintoneQuery
            For Each n As String In names
                If _orderBy.ContainsKey(n) Then
                    _orderBy(n) = True
                Else
                    _orderBy.Add(n, True)
                End If
            Next

            Return Me

        End Function

        ''' <summary>
        ''' 降順のオーダー設定を行う
        ''' </summary>
        ''' <param name="names"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Descending(ParamArray names As String()) As kintoneQuery
            For Each n As String In names
                If _orderBy.ContainsKey(n) Then
                    _orderBy(n) = False
                Else
                    _orderBy.Add(n, False)
                End If
            Next

            Return Me

        End Function

        ''' <summary>
        ''' 与えられたパラメーターを元にクエリ式を組み立てる
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Build(Optional ByVal isEncode As Boolean = False) As String
            Dim result As String = ""

            result = _query

            '順序
            If _orderBy.Count > 0 Then
                Dim os As New List(Of String)
                For Each o As KeyValuePair(Of String, Boolean) In _orderBy
                    os.Add(o.Key + " " + If(o.Value, " asc", " desc"))
                Next
                result += " order by " + String.Join(",", os)
            End If

            'Limit
            If _limit >= 0 Then result += " limit " + _limit.ToString

            'Offset
            If _offset >= 0 Then result += " offset " + _offset.ToString

            'クエリ条件を作成
            If result.Length > 0 Then
                If Not isEncode Then
                    result = "query=" + Trim(result)
                Else
                    result = "query=" + HttpUtility.UrlEncode(Trim(result))
                End If
            End If

            'Field
            If _fields.Count > 0 Then
                Dim fs As New List(Of String)
                For i As Integer = 0 To _fields.Count - 1
                    If Not isEncode Then
                        fs.Add("fields[" + i.ToString + "]=" + _fields(i))
                    Else
                        fs.Add(HttpUtility.UrlEncode("fields[" + i.ToString + "]") + "=" + HttpUtility.UrlEncode(_fields(i)))
                    End If

                Next

                result = String.Join("&", fs) + If(result.Length > 0, "&" + result, "")

            End If

            'アプリケーション番号がある場合、付与
            If Not String.IsNullOrEmpty(App) Then
                result = "app=" + App + If(result.Length > 0, "&" + result, "")
            End If

            Return result

        End Function

        Public Overrides Function ToString() As String
            Return Build()
        End Function


    End Class

End Namespace


