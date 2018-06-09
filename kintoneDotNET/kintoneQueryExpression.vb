Imports System.Reflection
Imports System.Linq.Expressions
Imports kintoneDotNET.API.Types

Namespace API

    ''' <summary>
    ''' レコードを検索するためのクエリを作成するクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneQueryExpression

        ''' <summary>
        ''' LINQのExpressionからクエリ式を生成する<br/>
        ''' 変換用のディクショナリはモデルから取得
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="expression"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Eval(Of T As AbskintoneModel)(ByVal expression As Expression(Of Func(Of T, Boolean))) As String
            Dim model As T = Activator.CreateInstance(Of T)()
            Return Eval(expression, model.GetToItemNameDic)
        End Function


        ''' <summary>
        ''' LINQのExpressionからクエリ式を生成する<br/>
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="expression">Booleanを返却するexpression</param>
        ''' <param name="nameConvertor">クエリ上の項目名を特定の項目名に変換したい場合、変換用ディクショナリを設定</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Eval(Of T As AbskintoneModel)(ByVal expression As Expression(Of Func(Of T, Boolean)), _
                                                             ByVal nameConvertor As Dictionary(Of String, String)) As String
            'Linq Expressions :http://msdn.microsoft.com/ja-jp/library/system.linq.expressions(v=vs.100).aspx
            'Linq ExpressionType :http://msdn.microsoft.com/ja-jp/library/bb361179(v=vs.100).aspx

            Dim exp As Expression = expression.Body
            Dim query As List(Of kintoneFieldConnector) = searchExpression(exp)

            Dim result As String = ""
            For Each f As kintoneFieldConnector In query
                If nameConvertor IsNot Nothing AndAlso nameConvertor.ContainsKey(f.Field.FieldName) Then
                    f.Field.FieldName = nameConvertor(f.Field.FieldName)
                End If

                Dim conj As String = If(f.isAnd, " and ", " or ")
                If String.IsNullOrEmpty(result) Then
                    result += f.Field.ToString
                Else
                    result += conj + f.Field.ToString
                End If
            Next

            Return result

        End Function
        Private Shared Function searchExpression(ByVal exp As expression) As List(Of kintoneFieldConnector)
            Dim result As New List(Of kintoneFieldConnector)
            Dim expEnd As Boolean = False

            Dim field As kintoneFieldConnector = Nothing
            If TypeOf exp Is BinaryExpression Then
                Dim conjunction As ExpressionType = exp.NodeType
                '接続詞である場合、再度探索を行う
                If (conjunction = ExpressionType.And Or conjunction = ExpressionType.AndAlso) Or _
                    (conjunction = ExpressionType.Or Or conjunction = ExpressionType.OrElse) Then

                    Dim isAndParam As Boolean = True
                    If (conjunction = ExpressionType.Or Or conjunction = ExpressionType.OrElse) Then
                        isAndParam = False
                    End If

                    Dim expLeft As List(Of kintoneFieldConnector) = searchExpression(CType(exp, BinaryExpression).Left)
                    Dim expRight As List(Of kintoneFieldConnector) = searchExpression(CType(exp, BinaryExpression).Right)

                    If expLeft.Count > 0 Then result.AddRange(expLeft)

                    If expRight.Count > 0 Then
                        expRight.First.isAnd = isAndParam 'Left/Rightの境界を設定
                        result.AddRange(expRight)
                    End If
                Else
                    '接続詞でない場合、式を評価し終了する
                    field = New kintoneFieldConnector(evalExpression(CType(exp, BinaryExpression)))
                End If
            Else
                field = New kintoneFieldConnector(evalExpression(exp))
            End If

            If field IsNot Nothing Then result.Add(field)
            Return result

        End Function

        Private Shared Function evalExpression(ByVal exp As Expression) As kintoneQueryField
            Dim field As kintoneQueryField = Nothing
            Dim inverse As Boolean = False
            If exp.NodeType = ExpressionType.Not Then
                inverse = True
            End If
            Dim operand As ExpressionType = exp.NodeType
            Dim methodName As String = ""

            Dim left As Expression = Nothing
            Dim right As Expression = Nothing
            Dim prop As PropertyInfo = Nothing
            Dim name As String = ""
            Dim opr As String = ""
            Dim value As Object = ""

            If TypeOf exp Is BinaryExpression Then
                left = CType(exp, BinaryExpression).Left
                right = CType(exp, BinaryExpression).Right
            ElseIf TypeOf exp Is UnaryExpression Then
                left = CType(exp, UnaryExpression).Operand
                operand = left.NodeType
                If TypeOf left Is BinaryExpression Then
                    right = CType(left, BinaryExpression).Right
                    left = CType(left, BinaryExpression).Left
                End If
            Else
                left = exp
            End If

            If TypeOf left Is MethodCallExpression Then 'Method Callの場合Argumentsからleft/rightを設定
                Dim methodExp = CType(left, MethodCallExpression)
                methodName = methodExp.Method.Name

                If methodExp.Object IsNot Nothing Then 'インスタンスメソッドのコールの場合、Objectにインスタンスが格納される
                    left = methodExp.Object
                    right = methodExp.Arguments(0)
                Else
                    left = methodExp.Arguments(0)
                    right = methodExp.Arguments(1)
                End If

                Dim leftEval As Expression = extractMember(left)
                Dim rightMember As MemberExpression = extractMember(right)

                If Not (TypeOf leftEval Is MemberExpression AndAlso TypeOf CType(leftEval, MemberExpression).Member Is PropertyInfo) AndAlso
                    (rightMember IsNot Nothing AndAlso TypeOf rightMember.Member Is PropertyInfo) Then 'プロパティが右辺にくる場合逆転
                    Dim temp As Expression = left
                    left = rightMember
                    right = temp
                End If
            End If

            Dim leftMember As MemberExpression = extractMember(left)
            If leftMember Is Nothing Then Throw New InvalidExpressionException("左辺となるべきプロパティが見つかりません")

            name = leftMember.Member.Name
            value = extractValue(right)
            prop = leftMember.Member.DeclaringType.GetProperty(name)

            If inverse Then
                Select Case methodName
                    Case "LikeString" 'Like演算
                        opr = "not like"
                    Case "Contains" 'IN
                        opr = "not in"
                    Case Else 'methodがない場合、演算子から判定
                        Select Case operand
                            Case ExpressionType.Equal
                                opr = "!="
                            Case ExpressionType.NotEqual
                                opr = "="
                            Case ExpressionType.LessThan
                                opr = ">="
                            Case ExpressionType.GreaterThan
                                opr = "<="
                            Case ExpressionType.LessThanOrEqual
                                opr = ">"
                            Case ExpressionType.GreaterThanOrEqual
                                opr = "<"
                            Case Else
                                opr = "="
                        End Select
                End Select
            Else
                Select Case methodName
                    Case "LikeString" 'Like演算
                        opr = "like"
                    Case "Contains" 'IN
                        opr = "in"
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
            End If

            If prop.PropertyType = GetType(DateTime) Then 'DateTime型の場合比較値が特殊になるため、対応を行う
                Dim attr As kintoneItemAttribute = prop.GetCustomAttributes(GetType(kintoneItemAttribute), True).SingleOrDefault
                If attr IsNot Nothing Then
                    Dim dateFormat As String = If(Not String.IsNullOrEmpty(attr.FieldType), attr.FieldType, kintoneDatetime.YMDType)
                    value = kintoneDatetime.toKintoneFormat(value, dateFormat)
                Else
                    value = kintoneDatetime.toKintoneFormat(value)
                End If
            End If

            field = New kintoneQueryField(name, opr, valueToString(value))

            Return field

        End Function

        Private Shared Function extractMember(ByVal exp As Expression) As MemberExpression
            If TypeOf exp Is UnaryExpression Then
                Return extractMember(CType(exp, UnaryExpression).Operand)
            ElseIf TypeOf exp Is MemberExpression Then
                Return exp
            Else
                If TryCast(exp, MemberExpression) IsNot Nothing Then
                    Return CType(exp, MemberExpression)
                Else
                    Return Nothing
                End If
            End If
        End Function

        Private Shared Function extractValue(ByVal exp As Expression) As Object
            If TypeOf exp Is MethodCallExpression Then
                Dim method As MethodCallExpression = CType(exp, MethodCallExpression)
                Dim value As Object = method.Method.Invoke(method.Object, evalValues(method.Arguments).ToArray)
                Return value
            ElseIf TypeOf exp Is UnaryExpression Then
                Return extractValue(CType(exp, UnaryExpression).Operand)
            Else
                Return evalValue(exp)
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
    ''' フィールドの連結子を管理するクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneFieldConnector
        Public Property isAnd As Boolean = True
        Public Property Field As kintoneQueryField = Nothing
        Public Sub New()
        End Sub

        Public Sub New(ByVal field As kintoneQueryField, Optional ByVal isAnd As Boolean = True)
            Me.Field = field
            Me.isAnd = isAnd
        End Sub

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
