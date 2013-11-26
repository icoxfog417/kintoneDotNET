Imports System.Configuration

Public MustInherit Class AbskintoneTest

    ''' <summary>
    ''' テスト対象のアプリケーションIDをConfigファイルから読込
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property TargetAppId As String
        Get
            Return ConfigurationManager.AppSettings("testAppId")
        End Get
    End Property

    ''' <summary>
    ''' 更新/読込テストを行うのに使用するレコードの抽出条件を設定
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public ReadOnly Property QueryForUpdateAndRead As String
        Get
            Return "methodinfo=""ExecuteUpdateAndRead""" '文字列の比較について""でのくくりが必要
        End Get
    End Property

    ''' <summary>
    ''' Read/Write用レコードを抽出するユーティリティ関数
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function getRecordForUpdateAndRead() As kintoneTestModel

        Dim result As List(Of kintoneTestModel) = kintoneTestModel.Find(Of kintoneTestModel)(QueryForUpdateAndRead)

        If result IsNot Nothing AndAlso result.Count = 1 Then
            Return result(0)
        Else
            Return Nothing
        End If

    End Function

    ''' <summary>
    ''' リスト値の比較を行う(比較はEqualで行うため、必要に応じ実装が必要)
    ''' </summary>
    ''' <typeparam name="T"></typeparam>
    ''' <param name="left"></param>
    ''' <param name="right"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function ListEqual(Of T)(ByVal left As List(Of T), ByVal right As List(Of T)) As Boolean
        Dim result As Boolean = True
        If left.Count <> right.Count Then
            result = False
        End If

        If result Then
            For i As Integer = 0 To left.Count - 1
                If Not left(i).Equals(right(i)) Then
                    result = False
                End If
            Next
        End If

        Return result

    End Function

    ''' <summary>
    ''' Read/Write用レコードを初期化するユーティリティ関数
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Protected Function getInitializedRecord() As kintoneTestModel

        Dim record As kintoneTestModel = getRecordForUpdateAndRead()
        Dim item As New kintoneTestModel
        item.record_id = record.record_id
        item.methodinfo = "ExecuteUpdateAndRead"
        item.Update()

        Return item

    End Function

End Class
