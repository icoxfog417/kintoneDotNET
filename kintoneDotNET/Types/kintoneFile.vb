
Namespace API.Types

    ''' <summary>
    ''' kintone上の添付ファイルフィールド型に対応するためのクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneFile
        Implements IkintoneType

        Public Property contentType As String
        <kintoneItemAttribute()>
        Public Property fileKey As String
        Public Property name As String
        Public Property size As String

        Public Sub New()
        End Sub
        Public Sub New(ByVal fileKey As String)
            Me.fileKey = fileKey
        End Sub

        ''' <summary>
        ''' kintoneからfileKeyで指定されたファイルのダウンロードを行う
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function GetFile() As IO.MemoryStream
            Dim result As IO.MemoryStream = Nothing
            If Not String.IsNullOrEmpty(fileKey) Then
                result = kintoneAPI.DownloadFile(fileKey)
            End If
            Return result
        End Function

        ''' <summary>
        ''' ファイルのダウンロードを行い、HttpResponseで返却する
        ''' Webサイトなどでファイル出力を行うためのメソッド
        ''' </summary>
        ''' <remarks></remarks>
        Public Sub DownloadFile()
            Dim response As HttpResponse = HttpContext.Current.Response
            response.Clear()
            response.ContentType = contentType
            response.AddHeader("Content-Disposition", String.Format("attachment;filename={0}", name))
            response.BinaryWrite(GetFile.GetBuffer)
            response.End()
        End Sub

    End Class

End Namespace

