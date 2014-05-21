Imports kintoneDotNET.API.Types

Namespace API

    ''' <summary>
    ''' kintone上で発生したエラーを補足するための例外
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneException
        Inherits Exception

        Private _error As kintoneError = Nothing
        Public Property [error]() As kintoneError
            Get
                Return _error
            End Get
            Set(ByVal value As kintoneError)
                _error = value
            End Set
        End Property

        Private _message As String = ""
        Public Overrides ReadOnly Property Message As String
            Get
                If _error IsNot Nothing Then
                    Return _error.Summary
                Else
                    Return _message
                End If
            End Get
        End Property

        Public ReadOnly Property Detail As String
            Get
                If _error IsNot Nothing Then
                    Return _error.ToString
                Else
                    Return _message
                End If
            End Get
        End Property

        Public Sub New()
        End Sub

        Public Sub New(ByVal kerror As kintoneDotNET.API.Types.kintoneError)
            _error = kerror
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
