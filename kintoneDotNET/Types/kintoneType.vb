Imports Microsoft.VisualBasic
Imports System.Web.Script.Serialization
Imports System.Collections.ObjectModel
Imports System.Reflection
Imports System.Runtime.CompilerServices

Namespace API.Types

    ''' <summary>
    ''' kintoneのフィールド型に対応するクラスを示すための、マーカーインタフェース
    ''' </summary>
    ''' <remarks></remarks>
    Public Interface IkintoneType
    End Interface

    ''' <summary>
    ''' JSON読み取り用のクラス
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <remarks></remarks>
    Public Class kintoneRecords(Of T)
        Public Property app As String
        Public Property records As New List(Of T)
    End Class

    ''' <summary>
    ''' JSON送信用のクラス
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <remarks></remarks>
    Public Class kintoneRecord(Of T As AbskintoneModel)
        Public Property id As String
        Public Property record As T
        Public Sub New()
        End Sub
        Public Sub New(ByVal model As T)
            id = model.record_id
            record = model
        End Sub
    End Class

    ''' <summary>
    ''' JSONを送受信する際の、各項目に対応するクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneItem
        Property type As String
        Property value As Object
    End Class

    ''' <summary>
    ''' Create/Deleteなどでidの一覧を受け取るためのクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneIds
        Public Property app As String
        Property ids As List(Of String)
    End Class

    ''' <summary>
    ''' kintone上でのエラーを取得するためのクラス
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneError
        Public Property message As String
        Public Property id As String
        Public Property code As String
    End Class

End Namespace

