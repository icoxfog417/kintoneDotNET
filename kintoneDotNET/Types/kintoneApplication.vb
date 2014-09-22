
Namespace API.Types

    ''' <summary>
    ''' kintoneのアプリケーション情報
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneApplication

        Public Property appId As String = ""
        Public Property code As String = ""
        Public Property name As String = ""
        Public Property description As String = ""
        Public Property spaceId As String = ""
        Public Property threadId As String = ""
        Public Property createdAt As DateTime = Nothing
        Public Property creator As New kintoneUser()
        Public Property modifiedAt As DateTime = Nothing
        Public Property modifier As New kintoneUser()

    End Class

End Namespace
