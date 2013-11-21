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

        Public Overridable Property record_id As String = ""
        Public Overridable Property created_time As DateTime = DateTime.MinValue
        Public Overridable Property updated_time As DateTime = DateTime.MinValue
        Public Overridable Property create_usr As New kintoneUser()
        Public Overridable Property update_usr As New kintoneUser()
        Public Overridable Property status As String = ""
        Public Overridable Property work_usr As New kintoneUser()

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

        Public Shared Function Find(Of T As AbskintoneModel)(ByVal query As String) As List(Of T)
            Dim model As T = Activator.CreateInstance(Of T)()
            Dim api As New kintoneAPI(model.app)
            Return api.Find(Of T)(query)

        End Function

        Public Shared Function FindAll(Of T As AbskintoneModel)(ByVal query As String) As List(Of T)
            Dim model As T = Activator.CreateInstance(Of T)()
            Dim api As New kintoneAPI(model.app)
            Return api.FindAll(Of T)(query)

        End Function

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

        Public Function Delete(ByVal ids As List(Of String)) As Boolean
            Dim result As Object = InvokeAPI("Delete", {ids})
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
        Private Function InvokeAPI(ByVal methodName As String, Optional ByVal arguments As Object() = Nothing) As Object

            Dim api As New kintoneAPI(app)

            Dim method As MethodInfo = GetType(kintoneAPI).GetMethod(methodName)
            Dim generic As MethodInfo = method.MakeGenericMethod(Me.GetType)
            Dim result As Object = Nothing
            If arguments Is Nothing Then
                result = generic.Invoke(api, {Me})
            Else
                result = generic.Invoke(api, arguments)
            End If

            Return result

        End Function

        ''' <summary>
        ''' kintone上デフォルトで日本語である項目("レコード番号","作成日時" など)をプロパティ名に変換するためのDictionaryを取得する
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetDefaultToPropertyDic() As Dictionary(Of String, String)
            Return GetNameConvertDic(True)
        End Function

        ''' <summary>
        ''' プロパティ名をkintone上のデフォルト名称に変換するためのDictionaryを取得する
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function GetPropertyToDefaultDic() As Dictionary(Of String, String)
            Return GetNameConvertDic(False)
        End Function

        Private Shared Function GetNameConvertDic(Optional ByVal isDefaultToProperty As Boolean = True) As Dictionary(Of String, String)
            Dim dic As New Dictionary(Of String, String)
            dic.Add("レコード番号", "record_id")
            dic.Add("作成日時", "created_time")
            dic.Add("更新日時", "updated_time")
            dic.Add("作成者", "create_usr")
            dic.Add("更新者", "update_usr")
            dic.Add("ステータス", "status")
            dic.Add("作業者", "work_usr")

            If isDefaultToProperty Then
                Return dic
            Else
                Dim opposit As New Dictionary(Of String, String)
                For Each item As KeyValuePair(Of String, String) In dic
                    opposit.Add(item.Value, item.Key) '逆にする
                Next
                Return opposit
            End If

        End Function

    End Class

End Namespace
