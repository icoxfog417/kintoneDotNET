
Namespace API.Types

    ''' <summary>
    ''' kintone上のユーザーフィールド型に対応するためのクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneUser
        Implements IkintoneType

        Public Property code As String
        Public Property name As String

        Public Sub New()
        End Sub

        Public Sub New(ByVal obj As Object)

            If obj.Count > 0 Then
                Me.code = obj("code")
                Me.name = obj("name")
            End If

        End Sub

    End Class

End Namespace

