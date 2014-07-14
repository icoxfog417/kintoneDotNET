Imports System.Linq.Expressions

Namespace API

    ''' <summary>
    ''' kintoneのクエリを作成するための中間オブジェクト
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneQuery(Of T As AbskintoneModel)

        Private _model As T = Nothing
        Private _rawquery As String = ""
        Private _expression As Expression(Of Func(Of T, Boolean)) = Nothing
        Private _convertDictionary As New Dictionary(Of String, String)
        Private _offset As Integer = -1
        Private _limit As Integer = -1
        Private _orderBy As New Dictionary(Of String, Boolean)
        Private _fields As New List(Of String)
        Private _isAll As Boolean = False

        Public Sub New()
            _model = Activator.CreateInstance(Of T)()
            _convertDictionary = _model.GetToItemNameDic
        End Sub

        Public Sub New(ByVal query As String)
            Me.New()
            Me._rawquery = query
        End Sub

        Public Sub New(ByVal isAll As Boolean)
            Me.New()
            Me._isAll = isAll
        End Sub

        Public Shared Function Make() As kintoneQuery(Of T)
            Return New kintoneQuery(Of T)()
        End Function

        Public Shared Function Make(ByVal query As String) As kintoneQuery(Of T)
            Return New kintoneQuery(Of T)(query)
        End Function

        Public Function Clone() As kintoneQuery(Of T)
            Dim c As New kintoneQuery(Of T)(_rawquery)
            c._offset = Me._offset
            c._limit = Me._limit
            c._orderBy = New Dictionary(Of String, Boolean)(Me._orderBy)
            c._fields = New List(Of String)(Me._fields)
            Return c

        End Function

        ''' <summary>
        ''' 抽出条件の設定を行う
        ''' </summary>
        ''' <param name="expression"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Where(ByVal expression As Expression(Of Func(Of T, Boolean))) As kintoneQuery(Of T)
            _expression = expression
            Return Me
        End Function

        ''' <summary>
        ''' 項目変換用のディクショナリを設定
        ''' </summary>
        ''' <param name="dic"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ConvertBy(ByVal dic As Dictionary(Of String, String)) As kintoneQuery(Of T)
            _convertDictionary = dic
            Return Me
        End Function

        ''' <summary>
        ''' 項目変換をオフにする
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ConvertOff() As kintoneQuery(Of T)
            _convertDictionary.Clear()
            Return Me
        End Function

        ''' <summary>
        ''' 項目変換をオンにする
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function ConvertOn() As kintoneQuery(Of T)
            _convertDictionary = _model.GetToItemNameDic
            Return Me
        End Function

        ''' <summary>
        ''' レコードの読み出し位置(Offset)の指定を行う
        ''' </summary>
        ''' <param name="index"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Offset(ByVal index As Integer) As kintoneQuery(Of T)
            _offset = index
            Return Me
        End Function

        ''' <summary>
        ''' 抽出するレコード件数(Limit)の指定を行う
        ''' </summary>
        ''' <param name="size"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Limit(ByVal size As Integer) As kintoneQuery(Of T)
            _limit = size
            Return Me
        End Function

        ''' <summary>
        ''' 抽出対象のフィールドを指定する
        ''' </summary>
        ''' <param name="names"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Fields(ParamArray names As String()) As kintoneQuery(Of T)
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
        Public Function Ascending(ParamArray names As String()) As kintoneQuery(Of T)
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
        Public Function Descending(ParamArray names As String()) As kintoneQuery(Of T)
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

            result = _rawquery

            '条件
            If _expression IsNot Nothing Then
                result += kintoneQueryExpression.Eval(Of T)(_expression, _convertDictionary)
            End If

            '順序
            If _orderBy.Count > 0 Then
                Dim os As New List(Of String)
                For Each o As KeyValuePair(Of String, Boolean) In _orderBy
                    Dim key As String = If(_convertDictionary.ContainsKey(o.Key), _convertDictionary(o.Key), o.Key)
                    os.Add(key + " " + If(o.Value, " asc", " desc"))
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
                    Dim fName As String = If(_convertDictionary.ContainsKey(_fields(i)), _convertDictionary(_fields(i)), _fields(i))

                    If Not isEncode Then
                        fs.Add("fields[" + i.ToString + "]=" + fName)
                    Else
                        fs.Add(HttpUtility.UrlEncode("fields[" + i.ToString + "]") + "=" + HttpUtility.UrlEncode(fName))
                    End If

                Next

                result = String.Join("&", fs) + If(result.Length > 0, "&" + result, "")

            End If

            'アプリケーション番号を付与
            result = "app=" + _model.app + If(result.Length > 0, "&" + result, "")

            Return result

        End Function

        Public Overrides Function ToString() As String
            Return Build()
        End Function

        ''' <summary>
        ''' リスト型への暗黙型変換
        ''' </summary>
        ''' <param name="q"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Widening Operator CType(ByVal q As kintoneQuery(Of T)) As List(Of T)
            Dim model As T = Activator.CreateInstance(Of T)()
            Dim api As New kintoneAPI(model.app)
            If Not q._isAll Then
                Return api.Find(q)
            Else
                Return api.FindAll(q)
            End If
        End Operator

    End Class

End Namespace


