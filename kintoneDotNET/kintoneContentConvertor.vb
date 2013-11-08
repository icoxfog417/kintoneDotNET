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
        ''' kintoneへ送信するのは、UploadTargetAttribute属性がついている項目のみ
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <param name="serializer"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overrides Function Serialize(obj As Object, serializer As JavaScriptSerializer) As IDictionary(Of String, Object)

            Dim result As New Dictionary(Of String, Object)

            'UploadTargetAttributeがセットされている項目を取得してシリアライズ
            Dim objType As Type = GetType(T)
            Dim props As PropertyInfo() = objType.GetProperties()
            Dim targets = From p As PropertyInfo In props
                          Let attribute As UploadTargetAttribute = p.GetCustomAttributes(GetType(UploadTargetAttribute), True).SingleOrDefault
                          Where Not attribute Is Nothing
                          Select p, attribute

            For Each tgt In targets
                Dim value As Object = tgt.p.GetValue(obj, Nothing)
                Dim pType As Type = tgt.p.PropertyType
                If pType.IsGenericType Then
                    pType = pType.GetGenericArguments(0)
                End If

                If TypeOf value Is IList Then
                    Dim list As New List(Of Object)
                    For Each v In value
                        list.Add(makeKintoneItem(pType, tgt.attribute, v, True))
                    Next
                    result.Add(tgt.p.Name, New With {.value = list})
                Else
                    result.Add(tgt.p.Name, makeKintoneItem(pType, tgt.attribute, value))
                End If
            Next

            Return result

        End Function

        ''' <summary>
        ''' Serialize時、項目の型に応じた変換処理を行う
        ''' </summary>
        ''' <param name="pType"></param>
        ''' <param name="value"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function makeKintoneItem(ByVal pType As Type, ByVal attr As UploadTargetAttribute, ByVal value As Object, Optional ByVal isRaw As Boolean = False) As Object
            Dim result As Object = Nothing
            Select Case pType
                Case GetType(DateTime)
                    Dim d As DateTime = value
                    If value Is Nothing Then
                        d = kintoneDatetime.InitialValue
                    End If
                    result = kintoneDatetime.toKintoneDate(d, attr.FieldType)
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
            Dim objType As Type = GetType(T)
            Dim m As T = Activator.CreateInstance(objType)

            '共通オブジェクト格納
            If dictionary.ContainsKey("レコード番号") Then m.record_id = dictionary("レコード番号")("value")
            If dictionary.ContainsKey("作成日時") Then m.create_time = kintoneDatetime.kintoneToDatetime(dictionary("作成日時"))
            If dictionary.ContainsKey("更新日時") Then m.update_time = kintoneDatetime.kintoneToDatetime(dictionary("更新日時"))
            If dictionary.ContainsKey("作成者") Then m.create_usr = New kintoneUser(dictionary("作成者"))
            If dictionary.ContainsKey("更新者") Then m.update_usr = New kintoneUser(dictionary("更新者"))
            If dictionary.ContainsKey("ステータス") Then m.status = dictionary("ステータス")("value")
            If dictionary.ContainsKey("作業者") Then m.work_usr = New kintoneUser(dictionary("作業者"))

            Dim props As PropertyInfo() = objType.GetProperties()
            For Each p As PropertyInfo In props
                If dictionary.ContainsKey(p.Name) Then
                    Dim value As Object = readKintoneItem(p.PropertyType, dictionary(p.Name))
                    p.SetValue(m, _serializer.ConvertToType(value, p.PropertyType), Nothing)
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
        Private Function readKintoneItem(ByVal pType As Type, ByVal obj As Object) As Object
            Dim result As Object = obj("value")
            Select Case pType
                Case GetType(DateTime)
                    Dim d As DateTime = Nothing
                    result = kintoneDatetime.kintoneToDatetime(obj)
                Case GetType(Decimal), GetType(Double), GetType(Integer)
                    If result Is Nothing OrElse String.IsNullOrWhiteSpace(result.ToString) Then
                        result = 0 '初期値を設定
                    End If
            End Select

            Return result

        End Function


    End Class

End Namespace
