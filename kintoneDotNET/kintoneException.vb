Imports kintoneDotNET.API.Types

Namespace API

    ''' <summary>
    ''' kintone上で発生したエラーを補足するための例外
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneException
        Inherits Exception

        Private _detail As kintoneError = Nothing
        Public Property Detail() As kintoneError
            Get
                Return _detail
            End Get
            Set(ByVal value As kintoneError)
                _detail = value
            End Set
        End Property

        Private _message As String = ""
        Public Overrides ReadOnly Property Message As String
            Get
                If _detail IsNot Nothing Then
                    Return _detail.message
                Else
                    Return _message
                End If
            End Get
        End Property

        Public Sub New()
        End Sub

        Public Sub New(ByVal kerror As kintoneDotNET.API.Types.kintoneError)
            _detail = kerror
        End Sub

        Public Sub New(ByVal message As String)
            _message = message
        End Sub

        Public Sub New(ByVal message As String, ByVal ex As Exception)
            MyBase.New(message, ex)
            _message = message
        End Sub

    End Class

End Namespace
