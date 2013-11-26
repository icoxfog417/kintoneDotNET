Imports Microsoft.VisualBasic
Imports System.Web.Script.Serialization
Imports System.Collections.ObjectModel
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Collections.Specialized

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
        Public Property details As New List(Of kintoneErrorDetail)

        Public Overrides Function ToString() As String
            Dim result As String = ""
            result += "id:" + id + " code:" + code + vbCrLf
            result += message + vbCrLf

            If details.Count > 0 Then
                result += "---------- detail ----------" + vbCrLf
                For Each d As kintoneErrorDetail In details.OrderBy(Function(x) x.index)
                    result += d.ToString + vbCrLf
                Next
            End If

            Return result
        End Function

    End Class

    ''' <summary>
    ''' エラーの詳細を格納するためのクラス<br/>
    ''' APIの定義書に未記載の内容を元にしているため、APIバージョンアップに伴い編集が必要な可能性あり
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneErrorDetail
        ''' <summary>
        ''' エラーの発生したレコード位置
        ''' </summary>
        Public Property index As Integer

        ''' <summary>
        ''' 各フィールドで発生したエラーメッセージ
        ''' </summary>
        Public Property messages As List(Of String)
        Public Sub New()
        End Sub
        Public Sub New(ByVal index As String, ByVal messages As List(Of String))
            If Integer.TryParse(index, Me.index) Then
            End If
            Me.messages = messages
        End Sub

        Public Overrides Function ToString() As String
            Dim result As String = ""
            '見出しをそろえるためにpadする。3は適当な値(更新限界が数百件なので、これで十分なはず)
            result += "[" + index.ToString.PadLeft(3) + "]:" + String.Join(" / ", messages)

            Return result
        End Function

    End Class

End Namespace

