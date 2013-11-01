Imports System.Reflection
Imports kintoneDotNET.API.Types

Namespace API

    ''' <summary>
    ''' kintoneのレコードに対応するモデルの基となる、抽象クラス
    ''' </summary>
    ''' <remarks>
    ''' 読込は全てのプロパティに対して行われるが、kintoneへ送信するのはUploadTargetAttribute属性がついている項目のみとしている
    ''' このため、kintone側で更新したいプロパティについては&lt;UploadTarget()&gt;を付与する。
    ''' 
    ''' リスト型のデータ(チェックボックスリストや添付ファイルなど)については、List(Of )で宣言を行う必要あり
    ''' </remarks>
    Public MustInherit Class AbskintoneModel

        Public MustOverride ReadOnly Property app As String

        Public Property record_id As String = ""
        Public Property create_time As DateTime
        Public Property update_time As DateTime
        Public Property create_usr As kintoneUser = Nothing
        Public Property update_usr As kintoneUser = Nothing
        Public Property status As String = ""
        Public Property work_usr As kintoneUser = Nothing

        Private _error As New kintoneError

        ''' <summary>
        ''' 処理結果を取得する
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetError() As kintoneError
            Return _error
        End Function
        Public Function isError() As Boolean
            Return kintoneAPI.isError(_error)
        End Function

        ''' <summary>
        ''' 自身のデータを操作するAPIを提供する
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetAPI() As kintoneAPI
            Return New kintoneAPI(app)
        End Function

        ''' <summary>
        ''' kintone上のレコードのURL
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public ReadOnly Property record_show_url() As String
            Get
                Dim api As New kintoneAPI(app)
                Dim url As String = "https://" + kintoneAPI.Host + "/k/" + app + "/show?record=" + record_id
                Return url
            End Get
        End Property

        Public Overridable Function Create() As List(Of String)
            Dim result As Object = InvokeAPI("Create")
            If TypeOf result Is List(Of String) Then
                Return CType(result, List(Of String))
            Else
                Return Nothing
            End If
        End Function

        Public Function Update() As Boolean
            Dim result As Object = InvokeAPI("Update")
            If TypeOf result Is Boolean Then
                Return CType(result, Boolean)
            Else
                Return False
            End If
        End Function

        Public Function Delete() As Boolean
            Dim result As Object = InvokeAPI("Delete", {New List(Of String) From {record_id}})
            If TypeOf result Is Boolean Then
                Return CType(result, Boolean)
            Else
                Return False
            End If
        End Function

        ''' <summary>
        ''' APIクラスで宣言されているジェネリクスメソッドを、対応する型で呼び出すためのメソッド
        ''' </summary>
        ''' <param name="methodName">API.kintoneAPI上のメソッド名。kintoneAPI上でメソッド名の変更が行われた場合、呼び出し側で対応する必要あり</param>
        ''' <param name="arguments"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Function InvokeAPI(ByVal methodName As String, Optional ByVal arguments As Object() = Nothing) As Object

            Dim api As New kintoneAPI(app)

            Dim method As MethodInfo = GetType(kintoneAPI).GetMethod(methodName)
            Dim generic As MethodInfo = method.MakeGenericMethod(Me.GetType)
            Dim result As Object = Nothing
            If arguments Is Nothing Then
                result = generic.Invoke(api, {Me})
            Else
                result = generic.Invoke(api, arguments)
            End If
            _error = api.GetError

            Return result

        End Function

    End Class

End Namespace
