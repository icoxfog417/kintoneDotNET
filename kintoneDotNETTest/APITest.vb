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
    Inherits AbskintoneTest

    ''' <summary>
    ''' Readのテスト(単純に読み込みで例外が発生しないことを確認)
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub Read()

        Dim result As List(Of kintoneTestModel) = kintoneTestModel.Find(Of kintoneTestModel)(String.Empty)

        For Each item As kintoneTestModel In result
            Console.WriteLine(item)
        Next

    End Sub

    <TestMethod()>
    Public Sub ReadExpression()
        Dim querys As New Dictionary(Of String, String)
        Dim q As String = ""
        'Simple Expression
        q = kintoneQuery.Make(Of kintoneTestModel)(Function(x) x.status > "a")
        Console.WriteLine("querySimple1:" + q)
        Assert.AreEqual("ステータス > ""a""", q)

        q = kintoneQuery.Make(Of kintoneTestModel)(Function(x) Not x.status > "a")
        Console.WriteLine("querySimple2:" + q)
        Assert.AreEqual("ステータス <= ""a""", q)

        'Equal
        q = kintoneQuery.Make(Of kintoneTestModel)(Function(x) x.status.Equals(1) Or x.record_id <> 1 And Not x.numberField = 1, Nothing)
        Console.WriteLine("queryEqual:" + q)
        Assert.AreEqual("status = 1 or record_id != 1 and numberField != 1", q)

        'like
        q = kintoneQuery.Make(Of kintoneTestModel)(Function(x) x.status Like "A" And Not x.link Like "http", Nothing)
        Console.WriteLine("queryLike:" + q)
        Assert.AreEqual("status like ""A"" and link not like ""http""", q)

        'IN
        Dim ids As New List(Of String)() From {"1", "2", "3"}
        q = kintoneQuery.Make(Of kintoneTestModel)(Function(x) {"1", "a"}.Contains(x.radio) And Not ids.Contains(x.record_id) And New List(Of Integer)() From {1, 2}.Contains(x.numberField), Nothing)
        Console.WriteLine("queryArray&List:" + q)
        Assert.AreEqual("radio in (""1"",""a"") and record_id not in (""1"",""2"",""3"") and numberField in (1,2)", q)

        'method Equal
        q = kintoneQuery.Make(Of kintoneTestModel)(Function(x) x.status = String.Empty And x.created_time < DateTime.MaxValue, Nothing)
        Console.WriteLine("queryMethodCall:" + q)
        Assert.AreEqual("status = """" and created_time < ""9999-12-31T23:59:59+09:00""", q)

        'MemberEqual
        Dim m As New kintoneTestModel
        m.status = "X"
        q = kintoneQuery.Make(Of kintoneTestModel)(Function(x) x.status = m.status, Nothing)
        Console.WriteLine("queryMemberEval:" + q)
        Assert.AreEqual("status = ""X""", q)

    End Sub


    ''' <summary>
    ''' FindAllのテスト 全件取得可能かチェックする
    ''' テストアプリには制限を超えるようなレコード数は登録しないため、Limitを調整しテスト
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub FindAllTest()

        Dim result As List(Of kintoneTestModel) = kintoneTestModel.Find(Of kintoneTestModel)(String.Empty)

        Dim before As Integer = kintoneAPI.ReadLimit
        kintoneAPI.ReadLimit = 1
        Dim resultAll As List(Of kintoneTestModel) = kintoneTestModel.FindAll(Of kintoneTestModel)(String.Empty)

        Assert.AreEqual(result.Count, resultAll.Count)
        For Each item In result
            Dim sameItem As kintoneTestModel = resultAll.Find(Function(x) x.record_id = item.record_id)
            Assert.IsTrue(sameItem IsNot Nothing)
        Next

        kintoneAPI.ReadLimit = before

    End Sub


    ''' <summary>
    ''' レコードの登録/更新/削除
    ''' ※複数のTestMethodに分かれてCreate/Deleteを行うとidが混線して?予期しないエラーになる場合があるので、一つにまとめる
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub ExecuteCreateDeleteUnit()
        Const METHOD_NAME As String = "ExecuteCreateDeleteUnit"

        '事前に削除
        Dim remained As List(Of kintoneTestModel) = kintoneTestModel.Find(Of kintoneTestModel)(Function(x) x.methodinfo = METHOD_NAME).ToList
        kintoneTestModel.Delete(Of kintoneTestModel)(remained.Select(Function(x) x.record_id).ToList)

        '単一のケース -------------------------------------------------------------
        Dim item As New kintoneTestModel
        item.methodinfo = METHOD_NAME

        '登録
        Dim id As String = item.Create
        Assert.IsFalse(String.IsNullOrEmpty(id))
        Dim created As kintoneTestModel = kintoneTestModel.FindById(Of kintoneTestModel)(id)
        Assert.IsFalse(created Is Nothing)

        '削除
        Assert.IsTrue(item.Delete())
        Dim deleted As kintoneTestModel = kintoneTestModel.FindById(Of kintoneTestModel)(item.record_id)
        Assert.IsTrue(deleted Is Nothing)

        'Save(登録)
        id = item.Save
        Assert.IsFalse(String.IsNullOrEmpty(id))
        Dim savedc As kintoneTestModel = kintoneTestModel.FindById(Of kintoneTestModel)(id)
        Assert.IsFalse(savedc Is Nothing)

        '更新
        item.textarea = "update from unit"
        Assert.IsTrue(item.Update())
        Dim updated As kintoneTestModel = kintoneTestModel.FindById(Of kintoneTestModel)(item.record_id)
        Assert.AreEqual(item.textarea, updated.textarea)

        'Save(更新)
        item.textarea = "save from unit"
        Assert.IsTrue(String.IsNullOrEmpty(item.Save))
        Dim saved As kintoneTestModel = kintoneTestModel.FindById(Of kintoneTestModel)(item.record_id)
        Assert.AreEqual(item.textarea, saved.textarea)

        '削除して終了
        Assert.IsTrue(item.Delete())

     End Sub

    ''' <summary>
    ''' モデルのキーを使用した更新のテスト。<br/>
    ''' record_idがない場合、key項目を使用しkintoneで検索を行い、一致していればそのオブジェクトのidを使用して更新を行う
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub ExecuteUnitByKey()
        Const METHOD_NAME As String = "ExecuteUnitByKey"

        '事前に削除
        Dim remained As List(Of kintoneTestModel) = kintoneTestModel.Find(Of kintoneTestModel)(Function(x) x.methodinfo = METHOD_NAME).ToList
        kintoneTestModel.Delete(Of kintoneTestModel)(remained.Select(Function(x) x.record_id).ToList)

        Dim item As New kintoneTestModel
        item.methodinfo = METHOD_NAME

        '更新対象レコードを事前登録
        Dim id As String = item.Create
        Assert.IsFalse(String.IsNullOrEmpty(item.record_id))

        'キー推定による更新
        item.record_id = String.Empty 'レコードidをクリア
        item.numberField = 10.111
        Assert.IsTrue(item.Update())
        Dim updated As kintoneTestModel = kintoneTestModel.FindById(Of kintoneTestModel)(id)
        Assert.AreEqual(item.numberField, updated.numberField)

        'キー推定による削除
        item.record_id = String.Empty 'レコードidをクリア
        Assert.IsTrue(item.Delete())
        Dim deleted As kintoneTestModel = kintoneTestModel.FindById(Of kintoneTestModel)(id)
        Assert.IsTrue(deleted Is Nothing)

    End Sub

    <TestMethod()>
    Public Sub ExecuteCreateDeleteMulti()
        Const METHOD_NAME As String = "ExecuteCreateDeleteMulti"
        Const CNT_TEST_DATA As Integer = 2
        Dim before As Integer = kintoneAPI.ExecuteLimit

        '事前に削除
        Dim remained As List(Of kintoneTestModel) = kintoneTestModel.Find(Of kintoneTestModel)(Function(x) x.methodinfo = METHOD_NAME).ToList
        kintoneTestModel.Delete(Of kintoneTestModel)(remained.Select(Function(x) x.record_id).ToList)

        '複合のケース -------------------------------------------------------------
        Dim list As New List(Of kintoneTestModel)
        kintoneAPI.ExecuteLimit = 1 '並列実行を検証するため、Limitを下げる

        For i As Integer = 0 To CNT_TEST_DATA
            Dim m As New kintoneTestModel
            m.methodinfo = METHOD_NAME
            m.textarea = "bulk insert " + i.ToString
            list.Add(m)
        Next

        '登録
        Dim ids As List(Of kintoneTestModel) = kintoneTestModel.Create(list)
        Assert.AreEqual(list.Count, ids.Count)
        Dim inserted As List(Of kintoneTestModel) = kintoneTestModel.Find(Of kintoneTestModel)(Function(x) x.methodinfo = METHOD_NAME).OrderBy(Function(x) x.textarea).ToList
        For i As Integer = 0 To CNT_TEST_DATA
            Assert.AreEqual(list(i).textarea, inserted(i).textarea)
            list(i).textarea = "bulk saved " + i.ToString
        Next

        '削除
        Assert.IsTrue(kintoneTestModel.Delete(Of kintoneTestModel)(inserted.Select(Function(x) x.record_id).ToList))

        'Save(登録)
        ids = kintoneTestModel.Save(list)
        Assert.AreEqual(inserted.Count, ids.Count)
        Dim savedc As List(Of kintoneTestModel) = kintoneTestModel.Find(Of kintoneTestModel)(Function(x) x.methodinfo = METHOD_NAME).OrderBy(Function(x) x.textarea).ToList
        For i As Integer = 0 To CNT_TEST_DATA
            Assert.AreEqual(list(i).textarea, savedc(i).textarea)
            list(i).textarea = "bulk updated " + i.ToString
            list(i).record_id = savedc(i).record_id
        Next

        '更新
        Assert.IsTrue(kintoneTestModel.Update(list))
        Dim updated As List(Of kintoneTestModel) = kintoneTestModel.Find(Of kintoneTestModel)(Function(x) x.methodinfo = METHOD_NAME).OrderBy(Function(x) x.textarea).ToList
        For i As Integer = 0 To CNT_TEST_DATA
            Assert.AreEqual(list(i).textarea, updated(i).textarea)
            list(i).textarea = "bulk updated by save " + i.ToString
        Next

        'Save(更新)
        ids = kintoneTestModel.Save(list)
        Assert.IsTrue(ids Is Nothing)
        Dim saved As List(Of kintoneTestModel) = kintoneTestModel.Find(Of kintoneTestModel)(Function(x) x.methodinfo = METHOD_NAME).OrderBy(Function(x) x.textarea).ToList
        For i As Integer = 0 To CNT_TEST_DATA
            Assert.AreEqual(list(i).textarea, saved(i).textarea)
        Next

        '削除
        Assert.IsTrue(kintoneTestModel.Delete(Of kintoneTestModel)(saved.Select(Function(x) x.record_id).ToList))

        kintoneAPI.ExecuteLimit = before

    End Sub

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <remarks></remarks>
    <TestMethod()>
    Public Sub ExecuteSaveByOverLimitURL()

        Const METHOD_NAME As String = "ExecuteSaveByOverLimitURL"

        '事前に削除
        Dim remained As List(Of kintoneTestModel) = kintoneTestModel.Find(Of kintoneTestModel)(Function(x) x.methodinfo = METHOD_NAME).ToList
        kintoneTestModel.Delete(Of kintoneTestModel)(remained)

        'URLが限界をオーバーするオブジェクトを作成
        Dim overs As New List(Of kintoneTestModel)
        For i As Integer = 1 To 100
            Dim m As New kintoneTestModel(METHOD_NAME)
            m.my_key = i.ToString.PadLeft(64, "X")
            overs.Add(m)
        Next

        Dim saved As List(Of kintoneTestModel) = kintoneTestModel.Save(Of kintoneTestModel)(overs)
        Assert.IsTrue(saved IsNot Nothing)

        For Each item As kintoneTestModel In saved
            Assert.IsFalse(String.IsNullOrEmpty(item.record_id))
            item.record_id = String.Empty
        Next

        '削除
        kintoneTestModel.Delete(Of kintoneTestModel)(overs)

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

        item = getRecordForUpdateAndRead()
        Assert.AreEqual(addLog.changeYMD.ToString("yyyyMMdd"), item.changeLogs(0).changeYMD.ToString("yyyyMMdd"))
        Assert.AreEqual(addLog.historyDesc, item.changeLogs(0).historyDesc)

        'もう一件追加+内容を編集
        updLog.id = item.changeLogs(0).id
        item.changeLogs(0) = updLog
        item.changeLogs.Add(nextLog)

        Assert.IsTrue(item.Update())

        item = getRecordForUpdateAndRead()
        Assert.AreEqual(item.changeLogs.Count, 2)

        Dim log As ChangeLog = item.changeLogs.Find(Function(x) x.id = updLog.id)
        Assert.AreEqual(updLog.changeYMD.ToString("yyyyMMdd"), log.changeYMD.ToString("yyyyMMdd"))
        Assert.AreEqual(updLog.historyDesc, log.historyDesc)

    End Sub

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