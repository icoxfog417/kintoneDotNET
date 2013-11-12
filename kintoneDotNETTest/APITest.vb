Imports System.Text
Imports Microsoft.VisualStudio.TestTools.UnitTesting
Imports System.Configuration
Imports kintoneDotNET.API
Imports kintoneDotNET.API.Types
Imports System.IO

''' <summary>
''' kintoneAPIの単体テストを実施する
''' </summary>
''' <remarks></remarks>
<TestClass()>
Public Class APITest

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
    ''' Readのテスト(単純に読み込みで例外が発生しないことを確認)
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub Read()

        Dim kerror As kintoneError = Nothing
        Dim result As List(Of kintoneTestModel) = kintoneAPI.Find(Of kintoneTestModel)(String.Empty, kerror)

        Assert.IsTrue(kerror Is Nothing)

        For Each item As kintoneTestModel In result
            Console.WriteLine(item)
        Next

    End Sub

    ''' <summary>
    ''' FindAllのテスト 全件取得可能かチェックする
    ''' テストアプリには制限を超えるようなレコード数は登録しないため、Limitを調整しテスト
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub FindAllTest()

        Dim kerror As kintoneError = Nothing
        Dim result As List(Of kintoneTestModel) = kintoneAPI.Find(Of kintoneTestModel)(String.Empty, kerror)

        Assert.IsTrue(kerror Is Nothing)

        Dim before As Integer = kintoneAPI.ReadLimit
        kintoneAPI.ReadLimit = 1
        Dim resultAll As List(Of kintoneTestModel) = kintoneAPI.FindAll(Of kintoneTestModel)(String.Empty, kerror)

        Assert.AreEqual(result.Count, resultAll.Count)
        For Each item In result
            Dim sameItem As kintoneTestModel = resultAll.Find(Function(x) x.record_id = item.record_id)
            Assert.IsTrue(sameItem IsNot Nothing)
        Next

        kintoneAPI.ReadLimit = before

    End Sub


    ''' <summary>
    ''' Create/Deleteのテスト
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub ExcecuteCreateDelete()
        Const METHOD_NAME As String = "ExcecuteCreateDelete"

        Dim item As New kintoneTestModel
        item.methodinfo = METHOD_NAME

        Dim ids As List(Of String) = item.Create

        Assert.IsTrue(ids.Count > 0)

        item.GetAPI.Delete(Of kintoneTestModel)(ids)

        Dim kerror As kintoneError = Nothing
        Dim result As List(Of kintoneTestModel) = kintoneAPI.Find(Of kintoneTestModel)("methodinfo=""" + METHOD_NAME + """", kerror)

        Assert.IsTrue(kerror Is Nothing)


    End Sub

    ''' <summary>
    ''' 文字列フィールドのRead/Writeテスト
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub ExecuteUpdateReadString()
        Const updString As String = "文字列"
        Const updTexts As String = "テスト実行" + vbCrLf + "テスト実行"

        Dim item As kintoneTestModel = getInitializedRecord()
        Assert.IsFalse(String.IsNullOrEmpty(item.record_id))

        item.stringField = updString
        item.textarea = updTexts

        Assert.IsTrue(item.Update())
        Assert.IsFalse(item.isError)

        item = getRecordForUpdateAndRead()
        Assert.IsTrue(item.stringField = updString)
        Assert.IsTrue(item.textarea = updTexts)

    End Sub

    ''' <summary>
    ''' 日付/時刻/日付時刻型フィールドのRead/Writeテスト
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub ExecuteUpdateReadDate()

        Dim updDate As New DateTime(1000, 2, 3)
        Dim updTime As New DateTime(1000, 1, 2, 11, 22, 33)
        Dim updDateTime As New DateTime(9999, 12, 31, 23, 53, 59)

        Dim item As kintoneTestModel = getInitializedRecord()
        Assert.IsFalse(String.IsNullOrEmpty(item.record_id))

        item.dateField = updDate
        item.time = updTime
        item.datetimeField = updDateTime

        Assert.IsTrue(item.Update())
        Assert.IsFalse(item.isError)

        item = getRecordForUpdateAndRead()
        Assert.AreEqual(updDate.ToString("yyyyMMdd"), item.dateField.ToString("yyyyMMdd"))
        'TODO 要確認：更新時 秒の情報が落ちる？(秒の値が常に00になっている)
        Assert.AreEqual(updTime.ToString("HHmm"), item.time.ToString("HHmm"))
        Assert.AreEqual(updDateTime.ToString("yyyyMMdd HHmm"), item.datetimeField.ToString("yyyyMMdd HHmm"))


    End Sub

    ''' <summary>
    ''' 複数選択項目のRead/Writeテスト
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub ExecuteUpdateReadMultiSelect()
        Dim updCheck As New List(Of String) From {"check2"}
        Dim updSelect As New List(Of String) From {"select1", "select3"}

        Dim item As kintoneTestModel = getInitializedRecord()
        Assert.IsFalse(String.IsNullOrEmpty(item.record_id))

        item.checkbox = updCheck
        item.multiselect = updSelect

        Assert.IsTrue(item.Update())
        Assert.IsFalse(item.isError)

        item = getRecordForUpdateAndRead()
        Assert.IsTrue(ListEqual(Of String)(updCheck, item.checkbox))
        Assert.IsTrue(ListEqual(Of String)(updSelect, item.multiselect))

    End Sub

    ''' <summary>
    ''' kintoneへファイルをアップロード/ダウンロードする
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub UploadDownloadFile()
        Const DOWNLOAD_FILE As String = "downloadedFile.PNG"
        Dim item As kintoneTestModel = getInitializedRecord()

        'ダウンロードファイルを事前に削除
        System.IO.File.Delete(FileOutPath(DOWNLOAD_FILE))

        'ファイルをアップロード
        Dim file As PostedFile = New PostedFile(FileOutPath("_uploadFile.PNG"))
        item.attachfile.UploadFile(file)
        Assert.IsTrue(item.attachfile.Count > 0)
        item.Update()

        'ファイルをダウンロードし、書き出し
        item = getRecordForUpdateAndRead()
        Dim downloaded As MemoryStream = item.attachfile(0).GetFile()
        Using target As New FileStream(FileOutPath(DOWNLOAD_FILE), FileMode.Create)
            target.Write(downloaded.ToArray, 0, downloaded.ToArray.Length)
        End Using

        Assert.IsTrue(System.IO.File.Exists(FileOutPath(DOWNLOAD_FILE)))

    End Sub

    ''' <summary>
    ''' 内部テーブルのRead/Writeテスト
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub UploadDownloadSubTable()

        Dim addLog As New ChangeLog(New DateTime(1000, 1, 1), "The Beggining Day ")
        Dim nextLog As New ChangeLog(DateTime.MaxValue, "The End of Century")
        Dim updLog As New ChangeLog(New DateTime(1999, 7, 31), "The Ending Day")

        '現在のレコードを取得
        Dim item As kintoneTestModel = getInitializedRecord()
        Assert.IsTrue(item.changeLogs.Count = 0)

        '内部テーブルレコードを追加
        item.changeLogs.Add(addLog)
        Assert.IsTrue(item.Update())
        Assert.IsFalse(item.isError)

        item = getRecordForUpdateAndRead()
        Assert.AreEqual(addLog.changeYMD.ToString("yyyyMMdd"), item.changeLogs(0).changeYMD.ToString("yyyyMMdd"))
        Assert.AreEqual(addLog.historyDesc, item.changeLogs(0).historyDesc)

        'もう一件追加+内容を編集
        updLog.id = item.changeLogs(0).id
        item.changeLogs(0) = updLog
        item.changeLogs.Add(nextLog)

        Assert.IsTrue(item.Update())
        Assert.IsFalse(item.isError)

        item = getRecordForUpdateAndRead()
        Assert.AreEqual(item.changeLogs.Count, 2)

        Dim log As ChangeLog = item.changeLogs.Find(Function(x) x.id = updLog.id)
        Assert.AreEqual(updLog.changeYMD.ToString("yyyyMMdd"), log.changeYMD.ToString("yyyyMMdd"))
        Assert.AreEqual(updLog.historyDesc, log.historyDesc)

    End Sub

    ''' <summary>
    ''' Read/Write用レコードを抽出するユーティリティ関数
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function getRecordForUpdateAndRead() As kintoneTestModel

        Dim kerror As kintoneError = Nothing
        Dim result As List(Of kintoneTestModel) = kintoneAPI.Find(Of kintoneTestModel)(QueryForUpdateAndRead, kerror)

        If kerror Is Nothing AndAlso result.Count = 1 Then
            Return result(0)
        Else
            Console.WriteLine(kerror.message)
            Throw New Exception("UpdateReadテスト用レコードの読取に失敗しました")
        End If

    End Function

    Private Function ListEqual(Of T)(ByVal left As List(Of T), ByVal right As List(Of T)) As Boolean
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
    Private Function getInitializedRecord() As kintoneTestModel

        Dim record As kintoneTestModel = getRecordForUpdateAndRead()
        Dim item As New kintoneTestModel
        item.record_id = record.record_id
        item.methodinfo = "ExecuteUpdateAndRead"
        item.Update()
        If item.isError Then
            Console.WriteLine(item.GetError.message)
            Throw New Exception("レコードの初期化更新に失敗しました")
        End If

        Return item

    End Function

    ''' <summary>
    ''' テストファイル保管用フォルダへのパスを取得する
    ''' </summary>
    ''' <param name="fileName"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function FileOutPath(ByVal fileName As String)
        '{ProjectRoot}/bin/Debug(orRelease)/xxx.dll で実行されるため、Parent/Parentでルートに戻る
        Return New DirectoryInfo(Environment.CurrentDirectory).Parent.Parent.FullName.ToString + "/App_Data/" + fileName
    End Function


End Class