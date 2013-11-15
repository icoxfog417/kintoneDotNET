Imports System.Reflection
Imports System.Linq.Expressions

Namespace API

    ''' <summary>
    ''' レコードを検索するためのクエリを作成するクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneQuery

        ''' <summary>
        ''' LINQのExpressionからクエリ式を生成する<br/>
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="expression">Booleanを返却するexpression</param>
        ''' <param name="nameConvertor">クエリ上の項目名を特定の項目名に変換したい場合、変換用ディクショナリを設定</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function toQuery(Of T As AbskintoneModel)(ByVal expression As Expression(Of Func(Of T, Boolean)), _
                                                                Optional ByVal nameConvertor As Dictionary(Of String, String) = Nothing) As String
            'Linq Expressions :http://msdn.microsoft.com/ja-jp/library/system.linq.expressions(v=vs.100).aspx
            'Linq ExpressionType :http://msdn.microsoft.com/ja-jp/library/bb361179(v=vs.100).aspx

            Dim query As New List(Of kintoneFieldGroup)

            Dim exp As Expression = expression.Body
            Dim expEnd As Boolean = False
            Dim groupNow As New kintoneFieldGroup

            While Not expEnd
                Dim field As kintoneQueryField = Nothing
                If TypeOf exp Is BinaryExpression Then
                    Dim conjunction As ExpressionType = exp.NodeType
                    '接続詞である場合、Leftを次回の評価式として渡す(Right側から順に評価されていく)
                    If (conjunction = ExpressionType.And Or conjunction = ExpressionType.AndAlso) Or _
                        (conjunction = ExpressionType.Or Or conjunction = ExpressionType.OrElse) Then
                        'AndAlso/OrElseはkintone側がそもそも対応していないので反映は行われない
                        field = evalExpression(CType(exp, BinaryExpression).Right)
                        exp = CType(exp, BinaryExpression).Left

                        Dim isAnd As Boolean = True
                        If (conjunction = ExpressionType.Or Or conjunction = ExpressionType.OrElse) Then
                            isAnd = False
                        End If

                        If groupNow.isAnd <> isAnd Then
                            query.Insert(0, groupNow)
                            groupNow = New kintoneFieldGroup
                            groupNow.isAnd = isAnd
                        End If

                    Else
                        field = evalExpression(CType(exp, BinaryExpression)) '接続詞でない場合、式を評価し終了する
                        expEnd = True
                    End If
                Else
                    field = evalExpression(exp)
                    expEnd = True
                End If

                If nameConvertor IsNot Nothing AndAlso nameConvertor.ContainsKey(field.FieldName) Then
                    field.FieldName = nameConvertor(field.FieldName)
                End If
                groupNow.FieldList.Insert(0, field)

            End While

            query.Insert(0, groupNow)

            Dim result As String = ""
            For Each group As kintoneFieldGroup In query
                Dim conj As String = If(group.isAnd, " and ", " or ")
                result += String.Join(conj, group.FieldList.Select(Function(x) x.ToString).ToArray)
            Next

            Return result

        End Function

        Private Shared Function evalExpression(ByVal exp As Expression) As kintoneQueryField
            Dim field As kintoneQueryField = Nothing
            Dim operand As ExpressionType = exp.NodeType
            Dim methodName As String = ""

            Dim left As Expression = Nothing
            Dim right As Expression = Nothing
            Dim name As String = ""
            Dim opr As String = ""
            Dim value As String = ""

            If TypeOf exp Is BinaryExpression Then
                left = CType(exp, BinaryExpression).Left
                right = CType(exp, BinaryExpression).Right
            ElseIf TypeOf exp Is UnaryExpression Then
                left = CType(exp, UnaryExpression).Operand
            Else
                left = exp
            End If

            If TypeOf left Is MethodCallExpression Then 'Method Callの場合Argumentsからleft/rightを設定
                Dim methodExp = CType(left, MethodCallExpression)
                methodName = methodExp.Method.Name

                If methodExp.Object Is Nothing Then
                    left = methodExp.Arguments(0)
                    right = methodExp.Arguments(1)
                Else
                    left = methodExp.Object
                    right = methodExp.Arguments(0)
                End If

                If Not TypeOf left Is MemberExpression Then 'プロパティが右辺にくる場合逆転
                    Dim temp As Expression = left
                    left = right
                    right = temp
                End If

            End If

            name = CType(left, MemberExpression).Member.Name
            value = extractValue(right)

            If operand = ExpressionType.Not Then
                opr = "not "
            End If
            Select Case methodName
                Case "LikeString" 'Like演算
                    opr += "like"
                Case "Contains" 'IN
                    opr += "in"
                Case Else 'methodがない場合、演算子から判定
                    Select Case operand
                        Case ExpressionType.Equal
                            opr = "="
                        Case ExpressionType.NotEqual
                            opr = "!="
                        Case ExpressionType.LessThan
                            opr = "<"
                        Case ExpressionType.GreaterThan
                            opr = ">"
                        Case ExpressionType.LessThanOrEqual
                            opr = "<="
                        Case ExpressionType.GreaterThanOrEqual
                            opr = ">="
                        Case Else
                            opr = "="
                    End Select
            End Select

            field = New kintoneQueryField(name, opr, value)

            Return field

        End Function

        Private Shared Function extractValue(ByVal exp As Expression) As String
            If TypeOf exp Is ConstantExpression Then
                Return valueToString(CType(exp, ConstantExpression).Value)
            ElseIf TypeOf exp Is NewArrayExpression Then
                Return valueToString(CType(exp, NewArrayExpression).Expressions)
            ElseIf TypeOf exp Is ListInitExpression Then
                '初期化子の配列を作成する
                Dim initArgs As List(Of Expression) = (From x As ElementInit In CType(exp, ListInitExpression).Initializers
                                                        Select x.Arguments(0)).ToList
                Return valueToString(initArgs)
            ElseIf TypeOf exp Is MethodCallExpression Then
                Dim method As MethodCallExpression = CType(exp, MethodCallExpression)
                Dim value As Object = method.Method.Invoke(method.Object, evalValues(method.Arguments).ToArray)
                Return valueToString(value)
            ElseIf TypeOf exp Is MemberExpression Then
                Return valueToString(evalValue(exp))
            ElseIf TypeOf exp Is UnaryExpression Then
                Return extractValue(CType(exp, UnaryExpression).Operand)
            Else
                Return ""
            End If

        End Function

        Private Shared Function evalValue(ByVal value As Expression) As Object
            Return Expression.Lambda(value).Compile().DynamicInvoke
        End Function
        Private Shared Function evalValues(ByVal list As IList(Of Expression)) As List(Of Object)
            Dim result As List(Of Object) = (From x As Expression In list
                                              Select Expression.Lambda(x).Compile().DynamicInvoke).ToList
            Return result
        End Function


        Private Shared Function valueToString(ByVal obj As Object) As String
            Dim result As String = ""
            If TypeOf obj Is IList Then
                Dim list As IList = CType(obj, IList)
                For Each item In list
                    If String.IsNullOrEmpty(result) Then
                        result += escapeValue(item)
                    Else
                        result += "," + escapeValue(item)
                    End If
                Next
                result = "(" + result + ")"
            Else
                result = escapeValue(obj)
            End If
            Return result
        End Function

        Private Shared Function escapeValue(ByVal obj As Object) As String
            If TypeOf obj Is Integer Or TypeOf obj Is Decimal Or TypeOf obj Is Double Then
                Return obj
            ElseIf TypeOf obj Is ConstantExpression Then
                Return escapeValue(CType(obj, ConstantExpression).Value)
            Else
                Return """" + obj.ToString + """"
            End If
        End Function

    End Class

    ''' <summary>
    ''' クエリにおいてAnd/Orのグループを管理するためのクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneFieldGroup
        Public Property isAnd As Boolean = True
        Public Property FieldList As New List(Of kintoneQueryField)
    End Class

    ''' <summary>
    ''' 条件式を表現するクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneQueryField
        Public Property FieldName As String
        Public Property Operand As String
        Public Property Value As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal fieldName As String, ByVal opr As String, ByVal value As Object)
            Me.FieldName = fieldName
            Me.Operand = opr
            Me.Value = value.ToString
        End Sub

        Public Overrides Function ToString() As String
            Return FieldName + " " + Operand + " " + Value
        End Function

    End Class

End Namespace
