Imports System.Collections.ObjectModel
Imports System.Web.Script.Serialization
Imports System.Reflection
Imports kintoneDotNET.API.Types

Namespace API

    ''' <summary>
    ''' kintoneから返却されるJSONの読み込み、また送信する際のシリアライズを行うConvertor
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <remarks></remarks>
    Public Class kintoneContentConvertor(Of T As AbskintoneModel)
        Inherits JavaScriptConverter

        Private _serializer As New JavaScriptSerializer

        Public Overrides ReadOnly Property SupportedTypes As IEnumerable(Of Type)
            Get
                Return New ReadOnlyCollection(Of Type)(New List(Of Type) From {GetType(T)})
            End Get
        End Property

        ''' <summary>
        ''' kintoneへ送信する際のシリアライズを行う
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <param name="serializer"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function Serialize(obj As Object, serializer As JavaScriptSerializer) As IDictionary(Of String, Object)

            Dim result As New Dictionary(Of String, Object)

            'kintoneItemAttributeがセットされていて、更新対象である項目を取得してシリアライズ
            Dim objType As Type = obj.GetType
            Dim props As PropertyInfo() = objType.GetProperties()
            Dim targets = From p As PropertyInfo In props
                          Let attribute As kintoneItemAttribute = p.GetCustomAttributes(GetType(kintoneItemAttribute), True).SingleOrDefault
                          Where isSendTarget(obj, p, attribute)
                          Select p, attribute

            For Each tgt In targets
                Dim value As Object = tgt.p.GetValue(obj, Nothing)
                Dim pType As Type = getGenericsType(tgt.p.PropertyType)

                'リスト型の場合の処理
                If TypeOf value Is IList Then
                    Dim list As New List(Of Object)
                    For Each v In value
                        list.Add(makeKintoneItem(pType, tgt.attribute, v, serializer, True))
                    Next
                    result.Add(tgt.p.Name, New With {.value = list})
                Else
                    '通常のフィールド
                    result.Add(tgt.p.Name, makeKintoneItem(pType, tgt.attribute, value, serializer))
                End If
            Next

            Return result

        End Function

        ''' <summary>
        ''' 送信対象の項目であるかを判定する
        ''' </summary>
        ''' <param name="propInfo">項目のプロパティ情報</param>
        ''' <param name="attribute">属性情報</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function isSendTarget(ByVal obj As Object, ByVal propInfo As PropertyInfo, ByVal attribute As kintoneItemAttribute) As Boolean
            Dim isTarget As Boolean = False

            'isUploadであるものは対象(基本)
            If attribute IsNot Nothing AndAlso attribute.isUpload Then
                isTarget = True
            End If

            Return isTarget

        End Function

        ''' <summary>
        ''' Serialize時、項目の型に応じた変換処理を行う
        ''' </summary>
        ''' <param name="pType"></param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function makeKintoneItem(ByVal pType As Type, ByVal attr As kintoneItemAttribute, ByVal value As Object, ByVal serializer As JavaScriptSerializer, Optional ByVal isRaw As Boolean = False) As Object
            Dim result As Object = Nothing

            If pType.BaseType = GetType(kintoneSubTableItem) Then '内部テーブル項目
                Dim id As String = CType(value, kintoneSubTableItem).id
                Dim tableData As IDictionary(Of String, Object) = Serialize(value, serializer)
                If tableData.ContainsKey("id") Then
                    tableData.Remove("id")
                    If Not String.IsNullOrEmpty(id) Then
                        result = New With {.id = id, .value = tableData}
                    Else
                        result = New With {.value = tableData}
                    End If
                End If
            Else

                Select Case pType
                    Case GetType(DateTime)
                        Dim d As DateTime = value
                        If value Is Nothing OrElse d.Equals(DateTime.MinValue) Then
                            result = "" 'Requiredの場合kintoneDatetime.InitialValueをセットした方がいいかも
                        Else
                            result = kintoneDatetime.toKintoneFormat(d, attr.FieldType)
                        End If
                    Case GetType(kintoneFile)
                        Dim obj As kintoneFile = CType(value, kintoneFile)
                        result = New With {.fileKey = obj.fileKey}
                    Case Else
                        If (value Is Nothing OrElse String.IsNullOrEmpty(value.ToString)) And _
                            Not attr.InitialValue Is Nothing Then
                            result = attr.InitialValue
                        Else
                            result = value
                        End If
                End Select

            End If

            If isRaw Then
                Return result
            Else
                Return New With {.value = result} 'kintoneで値を送る場合、value:"xxxx"という形にする
            End If

        End Function

        ''' <summary>
        ''' kintoneから受け取ったJSONを、モデル型配列に変換する
        ''' </summary>
        ''' <param name="dictionary"></param>
        ''' <param name="type"></param>
        ''' <param name="serializer"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function Deserialize(dictionary As IDictionary(Of String, Object), type As Type, serializer As JavaScriptSerializer) As Object
            Dim objType As Type = type
            Dim m = Activator.CreateInstance(objType)

            'デフォルトプロパティ格納
            If TypeOf m Is AbskintoneModel Then
                For Each item As KeyValuePair(Of String, String) In m.GetToPropertyDic
                    If dictionary.ContainsKey(item.Key) Then 'デフォルト名称が使用されている場合
                        Dim defaultProp As PropertyInfo = objType.GetProperty(item.Value)
                        If defaultProp IsNot Nothing Then
                            Dim value As Object = readKintoneItem(defaultProp.PropertyType, dictionary(item.Key)("value"), serializer)
                            If defaultProp.GetSetMethod() IsNot Nothing Then 'Setterが存在する場合、値をセット
                                defaultProp.SetValue(m, value, Nothing)
                            End If
                        End If
                    End If
                Next
            End If

            'その他プロパティ処理
            Dim props As PropertyInfo() = objType.GetProperties()
            For Each p As PropertyInfo In props
                If dictionary.ContainsKey(p.Name) Then
                    Dim value As Object = readKintoneItem(p.PropertyType, dictionary(p.Name)("value"), serializer)
                    If p.GetSetMethod() IsNot Nothing Then 'Setterが存在する場合、値をセット
                        p.SetValue(m, value, Nothing)
                    End If
                End If
            Next

            Return m

        End Function

        ''' <summary>
        ''' モデルのプロパティ型に合わせてデータのセットを行う
        ''' </summary>
        ''' <param name="pType"></param>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function readKintoneItem(ByVal pType As Type, ByVal obj As Object, ByVal serializer As JavaScriptSerializer) As Object
            Dim result As Object = obj
            Dim t As Type = getGenericsType(pType)

            If t.BaseType = GetType(kintoneSubTableItem) Then '内部テーブル項目
                Dim list As ArrayList = CType(result, ArrayList)
                Dim genericsType As Type = GetType(List(Of )).MakeGenericType(t)
                Dim objArray = Activator.CreateInstance(genericsType)
                For Each item As Object In list
                    Dim innerTableRow = Deserialize(item("value"), t, serializer)
                    innerTableRow.id = item("id") '親クラスが kintoneSubTableItem であることは保証されているため、idプロパティは必ず存在する
                    objArray.Add(innerTableRow)
                Next
                result = objArray
            Else
                Select Case t
                    Case GetType(DateTime)
                        Dim d As DateTime = Nothing
                        result = kintoneDatetime.kintoneToDatetime(obj)
                    Case GetType(kintoneUser)
                        result = New kintoneUser(obj)
                    Case GetType(Decimal), GetType(Double), GetType(Integer)
                        If result Is Nothing OrElse String.IsNullOrWhiteSpace(result.ToString) Then
                            result = 0 '初期値を設定
                        End If
                End Select

                result = _serializer.ConvertToType(result, pType) '指定タイプでDeserialize
            End If

            Return result

        End Function

        Private Function getGenericsType(ByVal pType As Type) As Type
            If pType.IsGenericType Then
                Return pType.GetGenericArguments(0) '複数のジェネリクスは考慮しない
            Else
                Return pType
            End If
        End Function

    End Class

End Namespace
