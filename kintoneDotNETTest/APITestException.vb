Imports System.Configuration
Imports System.Net
Imports kintoneDotNET.API
Imports kintoneDotNET.API.Types
Imports System.Reflection

''' <summary>
''' 例外発生時のテストを検証
''' </summary>
''' <remarks></remarks>
<TestClass()>
Public Class APITestException
    Inherits AbskintoneTest

    ''' <summary>
    ''' 接続確立前のエラー
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub ConnectionException()

        '意図的に接続エラーを発生させるため、手動でメソッドをコールする
        Dim headerMethod As MethodInfo = GetType(kintoneAPI).GetMethod("makeHeader", BindingFlags.Static Or BindingFlags.NonPublic)
        Dim responseMethod As MethodInfo = GetType(kintoneAPI).GetMethod("getResponse", BindingFlags.Static Or BindingFlags.NonPublic)

        Dim request As HttpWebRequest = headerMethod.Invoke(Nothing, {"records", "GET", ""})
        Dim kerror As kintoneError = Nothing
        Dim responseStr As String = ""

        'ヘッダの宛先ホストを空白にする(接続エラー)
        request.Host = "kintonetest.dummy"

        Dim param As Object() = {request, kerror}
        Using response As HttpWebResponse = responseMethod.Invoke(Nothing, param)
        End Using

        kerror = param(1)
        Console.WriteLine(kerror.message)
        Assert.IsFalse(String.IsNullOrEmpty(kerror.message))

    End Sub

    ''' <summary>
    ''' フィールド更新時エラー
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub FieldUpdateException()

        Dim item As kintoneTestModel = getInitializedRecord()
        Dim result As kintoneIndex = Nothing
        item.validationFld = "" '入力必須項目を空白にする
        item.datetimeField = DateTime.MinValue.AddDays(1) 'kintone上エラーになる日付を設定

        Try
            result = item.Update()
        Catch ex As kintoneException
            Console.WriteLine(ex.Message)
            Assert.IsTrue(Not String.IsNullOrEmpty(ex.Message))
            Assert.IsTrue(ex.error.details.Count > 0)
            Console.WriteLine(ex.error)
        End Try

        Assert.IsTrue(result Is Nothing)

    End Sub

    ''' <summary>
    ''' レコード更新時エラー
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub MultipleRecordException()

        Const METHOD_NAME As String = "MultipleRecordException"

        '事前に削除
        Dim remained As List(Of kintoneTestModel) = kintoneTestModel.Find(Of kintoneTestModel)(Function(x) x.methodinfo = METHOD_NAME).OrderBy(Function(x) x.textarea).ToList
        kintoneTestModel.Delete(Of kintoneTestModel)(remained.Select(Function(x) x.record_id).ToList)

        Dim item1 As New kintoneTestModel()
        item1.methodinfo = METHOD_NAME
        item1.validationFld = "" '入力必須項目を空白にする
        item1.dateField = DateTime.MinValue.AddDays(1) 'kintone上エラーになる日付を設定

        Dim item2 As New kintoneTestModel()
        item2.methodinfo = METHOD_NAME
        item2.datetimeField = DateTime.MinValue.AddDays(1) 'kintone上エラーになる日付を設定

        Dim ids As List(Of kintoneTestModel) = Nothing
        Try
            ids = kintoneTestModel.Create(New List(Of kintoneTestModel) From {item1, item2})
        Catch ex As kintoneException
            Console.WriteLine(ex.Message)
            Assert.IsTrue(Not String.IsNullOrEmpty(ex.Message))
            Assert.IsTrue(ex.error.details.Count > 0)
            Console.WriteLine(ex.error)
        End Try

        Assert.IsTrue(ids Is Nothing)

    End Sub

    ''' <summary>
    ''' リビジョンによる楽観ロックのエラー<br/>
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub RevisionLockException()
        Const METHOD_NAME As String = "RevisionLockException"

        '事前に削除
        Dim remained As List(Of kintoneTestModel) = kintoneTestModel.Find(Of kintoneTestModel)(Function(x) x.methodinfo = METHOD_NAME).ToList
        kintoneTestModel.Delete(Of kintoneTestModel)(remained.Select(Function(x) x.record_id).ToList)

        Dim item As New kintoneTestModel
        item.methodinfo = METHOD_NAME

        '登録を行う
        Dim index As kintoneIndex = item.Save
        Assert.IsFalse(String.IsNullOrEmpty(item.record_id))
        Assert.IsTrue(1, item.revision) '初回のリビジョンは1からスタート

        '初回の更新を行う
        item.radio = "radio1"
        index = item.Update()
        Assert.IsTrue(2, item.revision)

        'リビジョンがバッティングする更新を行う
        item.radio = "radio2"
        item.revision = 1

        Try
            index = item.Update()
        Catch ex As Exception
            Console.WriteLine(ex.Message)
            index = Nothing
        End Try

        Assert.IsTrue(index Is Nothing)

        'リビジョンを無視して更新する
        item.IgnoreRevision = True
        index = item.Update()
        Assert.IsTrue(index IsNot Nothing)

        '使用したレコードを削除
        item.Delete()

    End Sub

End Class
