Imports kintoneDotNET.API.Types
Imports System.Web.Script.Serialization
Imports System.Collections.ObjectModel
Imports System.Text.RegularExpressions

Namespace API.Convertor

    ''' <summary>
    ''' kintoneから返却されたエラー内容をデシリアライズするためのクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneErrorConvertor
        Inherits JavaScriptConverter

        Private _serializer As New JavaScriptSerializer

        Public Overrides ReadOnly Property SupportedTypes As IEnumerable(Of Type)
            Get
                Return New ReadOnlyCollection(Of Type)(New List(Of Type) From {GetType(kintoneError)})
            End Get
        End Property

        Public Overrides Function Deserialize(dictionary As IDictionary(Of String, Object), type As Type, serializer As JavaScriptSerializer) As Object
            '通常のシリアライズにまず任せる
            Dim kerror As kintoneError = _serializer.ConvertToType(dictionary, type)

            If dictionary.ContainsKey("errors") Then
                For Each e As Object In dictionary("errors")
                    Dim index As String = Regex.Replace(e.Key.ToString, "records\[(?<index>\d+)\](\w|\.)+", "${index}")
                    Dim eItem As kintoneErrorDetail = _serializer.ConvertToType(e.Value, GetType(kintoneErrorDetail))
                    If eItem IsNot Nothing AndAlso Integer.TryParse(index, eItem.index) Then
                        kerror.details.Add(eItem)
                    End If
                Next

            End If

            Return kerror

        End Function

        Public Overrides Function Serialize(obj As Object, serializer As JavaScriptSerializer) As IDictionary(Of String, Object)
            'シリアライズは行わない
            Throw New Exception("kintoneErrorのシリアライズは行われません")
        End Function

    End Class


End Namespace
