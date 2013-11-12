Imports System.IO

Namespace API.Types

    ''' <summary>
    ''' HttpPostedFileをエミュレートするクラス。ファイルパスから作成可能。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class PostedFile
        Inherits HttpPostedFileBase

        Private _fileName As String = ""
        Public Overrides ReadOnly Property FileName As String
            Get
                Return _fileName
            End Get
        End Property

        Private _contentType As String = ""
        Public Overrides ReadOnly Property ContentType As String
            Get
                Return _contentType
            End Get
        End Property

        Private _contentLength As Integer = 0
        Public Overrides ReadOnly Property ContentLength As Integer
            Get
                Return _contentLength
            End Get
        End Property

        Private _stream As IO.Stream = Nothing
        Public Overrides ReadOnly Property InputStream As IO.Stream
            Get
                Return _stream
            End Get
        End Property

        Public Sub New()
        End Sub

        ''' <summary>
        ''' 受け取ったパスから、FileStreamで読み込みを行いインスタンスを初期化する
        ''' </summary>
        ''' <param name="filePath"></param>
        ''' <remarks></remarks>
        Public Sub New(ByVal filePath As String)
            Dim file As New FileStream(filePath, FileMode.Open, FileAccess.Read)
            _fileName = Path.GetFileName(filePath)
            'TODO ContentType判定。 Upload時にいるわけではないので、必須ではない
            _contentLength = file.Length
            _stream = file

        End Sub

        Public Overrides Sub SaveAs(filename As String)
            Using target As New FileStream(filename, FileMode.Create)
                Dim fByte(_stream.Length - 1) As Byte
                _stream.Read(fByte, 0, fByte.Length)
                _stream.Close()
                target.Write(fByte, 0, fByte.Length)
            End Using
        End Sub

        Public Overrides Function GetHashCode() As Integer
            Return _stream.GetHashCode
        End Function

    End Class

End Namespace
