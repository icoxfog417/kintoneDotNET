Imports Microsoft.VisualBasic
Imports System.Net
Imports System.IO
Imports System.Reflection
Imports System.Runtime.Serialization
Imports System.Web.Script.Serialization
Imports System.Collections.ObjectModel
Imports kintoneDotNET.API.Types

Namespace API

    ''' <summary>
    ''' kintoneのREST APIをコールするためのクラス<br/>
    ''' REST APIについては<a href="https://developers.cybozu.com/ja/kintone-api/common-appapi.html">サイボウズ公式のドキュメント</a>参照
    ''' </summary>
    ''' <remarks>
    ''' </remarks>
    Public Class kintoneAPI

        Protected Const KINTONE_HOST As String = "{0}.cybozu.com"
        Protected Const KINTONE_PORT As String = "443"
        Protected Const KINTONE_API_FORMAT As String = "https://{0}/k/v1/{1}.json" 'TODO:APIのバージョンが上がれば変更する必要あり

        ''' <summary>
        ''' kintone APIの読み込み上限値
        ''' </summary>
        ''' <remarks></remarks>
        Public Const KINTONE_REC_LIMIT As Integer = 100 'TODO:レコード取得の上限値。変更されれば上げる必要あり

        ''' <summary>
        ''' kintone APIの更新上限値
        ''' </summary>
        ''' <remarks></remarks>
        Public Const KINTONE_EXE_LIMIT As Integer = 100 'TODO:レコード更新の上限値。変更されれば上げる必要あり(現時点で件数は同じだが、読込と別にしておく)

        ''' <summary>
        ''' 本APIで扱うレコードの最大値(余りに大きい件数の処理を途中で止めるための措置)
        ''' </summary>
        ''' <remarks></remarks>
        Public Const RECORD_LIMIT As Integer = 60000 'API内で処理する最大件数

        ''' <summary>
        ''' 送受信時のエンコード
        ''' </summary>
        ''' <value></value>
        Public Shared ReadOnly Property ApiEncoding As Encoding
            Get
                Return Encoding.GetEncoding("UTF-8") 'kintoneのエンコードはUTF-8
            End Get
        End Property

        Private Shared _domain As String = ""
        ''' <summary>
        ''' アプリケーションのドメイン。単一で使用することはまずないので、Protected化
        ''' </summary>
        ''' <value></value>
        Private Shared ReadOnly Property Domain As String
            Get
                If String.IsNullOrEmpty(_domain) Then _domain = ConfigurationManager.AppSettings("ktDomain")
                Return _domain
            End Get
        End Property

        Private Shared _host As String = ""
        ''' <summary>
        ''' kintoneのアクセス先。"xxx.cybozu.com"というようなアドレスで表現される(xxxはDomain)
        ''' </summary>
        ''' <value></value>
        Public Shared ReadOnly Property Host As String
            Get
                If String.IsNullOrEmpty(_host) Then _host = String.Format(KINTONE_HOST, Domain)
                Return _host
            End Get
        End Property

        Private _appId As String = ""
        ''' <summary>
        ''' アクセス先アプリケーションのID。コンストラクタで指定
        ''' </summary>
        ''' <value></value>
        Public ReadOnly Property AppId() As String
            Get
                Return _appId
            End Get
        End Property

        Private Shared _readLimit As Integer = 0
        ''' <summary>
        ''' レコードの読み取り上限を設定する。設定がない場合、APIの上限値が設定される<br/>
        ''' 上限値については、<a href="https://developers.cybozu.com/ja/kintone-api/apprec-readapi.html">レコード取得</a>の「クエリで条件を指定する」を参照
        ''' </summary>
        ''' <value></value>
        Public Shared Property ReadLimit As Integer
            Get
                If _readLimit < 1 Then
                    Return KINTONE_REC_LIMIT '設定がない場合、既定上限を返却
                Else
                    Return _readLimit
                End If
            End Get
            Set(value As Integer)
                _readLimit = value
            End Set
        End Property

        Private Shared _executeLimit As Integer = 0
        ''' <summary>
        ''' レコード更新件数の上限を設定する。設定がない場合、APIの上限値が設定される<br/>
        ''' </summary>
        ''' <value></value>
        Public Shared Property ExecuteLimit As Integer
            Get
                If _executeLimit < 1 Then
                    Return KINTONE_EXE_LIMIT '設定がない場合、既定上限を返却
                Else
                    Return _executeLimit
                End If
            End Get
            Set(value As Integer)
                _executeLimit = value
            End Set
        End Property

        ''' <summary>
        ''' Basic認証のためのキーを作成する。形式については<a href="http://developers.cybozu.com/ja/kintone-api/common-appapi.html">公式ドキュメント</a>を参照
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared ReadOnly Property AccessKey As String
            Get
                Dim id As String = ConfigurationManager.AppSettings("ktAccessId")
                Dim password As String = ConfigurationManager.AppSettings("ktAccessPassword")
                Dim key As String = Convert.ToBase64String(ApiEncoding.GetBytes(id + ":" + password))
                Return key
            End Get
        End Property

        ''' <summary>
        ''' kintoneのログインを行うためのキーを作成する
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared ReadOnly Property LoginKey As String
            Get
                Dim id As String = ConfigurationManager.AppSettings("ktLoginId")
                Dim password As String = ConfigurationManager.AppSettings("ktLoginPassword")
                Dim key As String = Convert.ToBase64String(ApiEncoding.GetBytes(id + ":" + password))
                Return key
            End Get
        End Property

        ''' <summary>
        ''' Proxyを経由する場合、プロキシの設定を行う
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared ReadOnly Property Proxy As WebProxy
            Get
                Dim myProxy As New WebProxy()
                Dim proxyAddress As String = ConfigurationManager.AppSettings("proxy")

                If Not String.IsNullOrEmpty(proxyAddress) Then
                    myProxy.Address = New Uri(proxyAddress)
                    Dim proxyUser As String = ConfigurationManager.AppSettings("proxyUser")
                    '認証が必要なプロキシの場合、認証情報を設定
                    If Not String.IsNullOrEmpty(proxyUser) Then
                        Dim credential As New NetworkCredential(proxyUser, ConfigurationManager.AppSettings("proxyPassword"))
                        myProxy.Credentials = credential
                    End If
                End If
                Return myProxy
            End Get
        End Property

        ''' <summary>
        ''' アプリケーションIDを指定し、APIの生成を行う
        ''' </summary>
        ''' <param name="id">アプリケーションID</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal id As String)
            Me._appId = id
        End Sub

        ''' <summary>
        ''' kintoneにアクセスするためのHTTPヘッダを作成する
        ''' </summary>
        ''' <param name="command">record/records、fileなどのREST APIの名称</param>
        ''' <param name="method">POST/GETなどのメソッドタイプ</param>
        ''' <param name="query">クエリ引数</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Shared Function makeHeader(ByVal command As String, ByVal method As String, Optional ByVal query As String = "") As HttpWebRequest
            Dim uri As String = String.Format(KINTONE_API_FORMAT, Host, command)
            If Not String.IsNullOrEmpty(query) Then
                uri += "?" + query
            End If
            Dim request As HttpWebRequest = DirectCast(Net.WebRequest.Create(uri), HttpWebRequest)
            request.Headers.Add("X-Cybozu-Authorization", LoginKey)
            If Not String.IsNullOrEmpty(AccessKey) Then
                request.Headers.Add("Authorization", "Basic " + AccessKey)
            End If
            request.Host = Host + ":" + KINTONE_PORT
            request.Method = method

            'ContentTypeの設定。デフォルト設定という感じにし、複雑な分岐・設定が必要な場合個別のメソッド側で実装する
            If Not method.ToUpper = "GET" Then 'GET(読込)の場合、ContentTypeを指定しない(指定すると400 BadRequestエラー)
                If command.ToUpper = "FILE" Then 'ファイル送信
                    request.ContentType = "multipart/form-data; boundary={0}" 'boundaryは自分で設定
                Else
                    request.ContentType = "application/json"
                End If
            End If

            request.Proxy = Proxy

            Return request

        End Function


        ''' <summary>
        ''' レコードの検索を行う<br/>
        ''' ※このメソッドは、kintoneのレコード数上限までしか取得を行いません。全件取得する場合はFindAllを使用してください
        ''' </summary>
        ''' <param name="query">
        ''' queryの形式については、<a href="https://developers.cybozu.com/ja/kintone-api/apprec-readapi.html">公式ドキュメント</a>を参照<br/>
        ''' item="xxxx"のように、文字列値は""で囲う必要があるため注意
        ''' </param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Find(Of T As AbskintoneModel)(Optional ByVal query As String = "") As List(Of T)
            Dim model As T = Activator.CreateInstance(Of T)()
            Dim q As String = "app=" + model.app + If(Not String.IsNullOrEmpty(query), "&query=" + HttpUtility.UrlEncode(query), String.Empty)

            Dim kerror As kintoneError = Nothing
            Dim result As List(Of T) = FindBase(Of T)(q, kerror)
            If kerror IsNot Nothing Then Throw New kintoneException(kerror)

            Return result

        End Function

        ''' <summary>
        ''' API上限値を超えたレコードの検索を行う<br/>
        ''' ※並列でリクエストを投げ取得を行うため、order等の指定は考慮されない。データ取得後LINQ等で並び替えを行う必要あり
        ''' </summary>
        ''' <param name="query"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function FindAll(Of T As AbskintoneModel)(Optional ByVal query As String = "") As List(Of T)
            Dim model As T = Activator.CreateInstance(Of T)()
            Dim q As String = "app=" + model.app + If(Not String.IsNullOrEmpty(query), "&query=" + HttpUtility.UrlEncode(query), "&query=") 'limitを指定する必要があるので、条件指定がなくてもqueryパラメータは付与

            Dim stepCount As Integer = 1
            Dim offset As Integer = 0
            Dim recordStillExist As Boolean = True
            Dim threadCount As Integer = 4 'TODO:並列実行数は様子を見ながら調整
            Dim result As New List(Of T)
            Dim kex As kintoneException = Nothing

            While recordStillExist
                Dim tasks As New List(Of Task(Of List(Of T)))
                For i As Integer = 1 To threadCount
                    tasks.Add(makeTaskForFind(Of T)(q, offset))
                    offset += ReadLimit
                Next
                Dim taskRuns As Task(Of List(Of T))() = tasks.ToArray

                Try
                    '並列処理実行
                    Array.ForEach(taskRuns, Sub(tx) tx.Start())

                    '実行待ち合わせ
                    Task.WaitAll(taskRuns)

                    '処理結果マージ
                    For Each tx In taskRuns
                        If tx.IsCompleted Then
                            If tx.Result.Count > 0 Then result.AddRange(tx.Result)
                            If tx.Result.Count < ReadLimit Then 'リミットのサイズより取得結果が少なくなれば、もう取得する必要がないため停止
                                recordStillExist = False
                            End If
                        End If
                    Next

                Catch ae As AggregateException
                    Dim ex As AggregateException = ae.Flatten
                    kex = New kintoneException(ex.Message, ex)
                    Dim kerror As kintoneError = ex.InnerExceptions.Where(Function(x) TypeOf x Is kintoneException).Select(Function(x) CType(x, kintoneException).error).FirstOrDefault
                    kex.error = kerror
                Catch ex As Exception
                    kex = New kintoneException(ex.Message, ex)
                End Try

                If stepCount * threadCount * KINTONE_REC_LIMIT > RECORD_LIMIT Then
                    kex = New kintoneException("レコード件数が多すぎます(" + RECORD_LIMIT + " 以上)")
                    recordStillExist = False
                End If

                If kex IsNot Nothing Then
                    recordStillExist = False '例外が発生した場合終了する
                End If

                stepCount += 1

            End While

            If kex IsNot Nothing Then Throw kex

            Return result

        End Function

        ''' <summary>
        ''' 検索を行うタスクを生成するための内部処理
        ''' </summary>
        Private Function makeTaskForFind(Of T As AbskintoneModel)(baseQuery As String, offset As Integer) As Task(Of List(Of T))

            Dim task As New Task(Of List(Of T))(Function()
                                                    Dim query = baseQuery + HttpUtility.UrlEncode(" limit " + ReadLimit.ToString + " offset " + offset.ToString)
                                                    Dim localError As kintoneError = Nothing
                                                    Dim list As List(Of T) = FindBase(Of T)(query, localError)
                                                    If Not localError Is Nothing Then
                                                        Throw New kintoneException(localError)
                                                    End If
                                                    Return list
                                                End Function
                                                )
            Return task
        End Function

        ''' <summary>
        ''' 検索処理の実体
        ''' </summary>
        Private Function FindBase(Of T As AbskintoneModel)(ByVal query As String, Optional ByRef kerror As kintoneError = Nothing) As List(Of T)

            Dim serialized As kintoneRecords(Of T) = Nothing
            Dim result As New List(Of T)

            Dim request As HttpWebRequest = makeHeader("records", "GET", query)
            Dim responseStr As String = ""

            Using response As HttpWebResponse = getResponse(request, kerror)
                If Not response Is Nothing Then
                    Dim reader As New StreamReader(response.GetResponseStream)
                    responseStr = reader.ReadToEnd
                End If
            End Using

            If Not String.IsNullOrEmpty(responseStr) Then
                Dim js As New JavaScriptSerializer
                Dim kc As New kintoneContentConvertor(Of T)
                js.RegisterConverters(New List(Of JavaScriptConverter) From {kc})

                serialized = js.Deserialize(Of kintoneRecords(Of T))(responseStr)

                For Each record As T In serialized.records
                    result.Add(record)
                Next

            End If

            Return result

        End Function

        ''' <summary>
        ''' レコード登録を行う(単一)
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Create(Of T As AbskintoneModel)(ByVal obj As T) As String
            Return taskCreate(Of T)(New List(Of T) From {obj}).FirstOrDefault
        End Function

        ''' <summary>
        ''' レコード登録を行う(複数件)
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="objs"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Create(Of T As AbskintoneModel)(ByVal objs As List(Of T)) As List(Of String)
            Dim result As List(Of List(Of String)) = executeParallel(Of T, List(Of String))(objs, AddressOf taskCreate(Of T))
            Dim flatten As New List(Of String)
            For Each li As List(Of String) In result
                flatten.AddRange(li)
            Next
            Return flatten
        End Function

        ''' <summary>
        ''' レコード登録処理の実体
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="objs"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function taskCreate(Of T As AbskintoneModel)(ByVal objs As List(Of T)) As List(Of String)
            'JsonオブジェクトをRequestに書き出し
            'http://stackoverflow.com/questions/9145667/how-to-post-json-to-the-server

            Dim request As HttpWebRequest = makeHeader("records", "POST")
            Dim sendData As New kintoneRecords(Of T)
            sendData.app = AppId
            sendData.records = objs

            'Request Body にJSONデータを書き込み
            Dim js As New JavaScriptSerializer
            Dim kc As New kintoneContentConvertor(Of T)
            js.RegisterConverters(New List(Of JavaScriptConverter) From {kc})

            Dim jsonData As String = js.Serialize(sendData)
            writeToRequestBody(request, jsonData)

            Dim list As New List(Of String)
            Dim kerror As kintoneError = Nothing
            Using response As HttpWebResponse = getResponse(request, kerror)
                If Not response Is Nothing Then
                    Dim reader As New StreamReader(response.GetResponseStream)
                    Dim serialized As kintoneIds = js.Deserialize(Of kintoneIds)(reader.ReadToEnd)
                    list = serialized.ids
                End If
            End Using
            If kerror IsNot Nothing Then Throw New kintoneException(kerror)

            Return list

        End Function

        ''' <summary>
        ''' レコードの更新を行う(単一)
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Update(Of T As AbskintoneModel)(ByVal obj As T) As Boolean
            Return taskUpdate(Of T)(New List(Of T) From {obj})
        End Function

        ''' <summary>
        ''' レコードの更新を行う(複数件)
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="objs"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Update(Of T As AbskintoneModel)(ByVal objs As List(Of T)) As Boolean
            Dim result As List(Of Boolean) = executeParallel(Of T, Boolean)(objs, AddressOf taskUpdate(Of T))
            Dim falseIndex As Integer = result.IndexOf(False)
            Return If(falseIndex < 0, True, False)

        End Function

        ''' <summary>
        ''' レコード更新処理の実体
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="objs"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function taskUpdate(Of T As AbskintoneModel)(ByVal objs As List(Of T)) As Boolean
            Dim request As HttpWebRequest = makeHeader("records", "PUT")
            Dim sendData As New kintoneRecords(Of kintoneRecord(Of T))
            sendData.app = AppId
            sendData.records = objs.Select(Function(x) New kintoneRecord(Of T)(x)).ToList

            'Request Body にJSONデータを書き込み
            Dim js As New JavaScriptSerializer
            Dim kc As New kintoneContentConvertor(Of T)
            js.RegisterConverters(New List(Of JavaScriptConverter) From {kc})

            Dim jsonData As String = js.Serialize(sendData)
            writeToRequestBody(request, jsonData)

            Dim kerror As kintoneError = Nothing
            Using response As HttpWebResponse = getResponse(request, kerror)
            End Using

            If kerror IsNot Nothing Then Throw New kintoneException(kerror)
            Return True

        End Function

        ''' <summary>
        ''' レコードの保存処理(単一)<br/>
        ''' モデル上 isKey = True と設定された項目をキーとし、一致するキーがある場合はUpdate、なければCreateを行う
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Save(Of T As AbskintoneModel)(ByVal obj As T) As String
            Dim ids As List(Of String) = Save(Of T)(New List(Of T) From {obj})
            If ids IsNot Nothing AndAlso ids.Count > 0 Then
                Return ids.First
            Else
                Return String.Empty
            End If
        End Function

        ''' <summary>
        ''' レコードの保存処理(複数件)
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="objs"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Save(Of T As AbskintoneModel)(ByVal objs As List(Of T)) As List(Of String)
            Dim key = From p As PropertyInfo In GetType(T).GetProperties
            Let attribute As kintoneItemAttribute = p.GetCustomAttributes(GetType(kintoneItemAttribute), True).SingleOrDefault
            Where attribute IsNot Nothing AndAlso attribute.isKey
            Select p, attribute

            If key Is Nothing OrElse key.Count <> 1 Then
                If key Is Nothing OrElse key.Count = 0 Then Throw New kintoneException("モデルにキーが設定されていません")
                If key.Count > 1 Then Throw New kintoneException("モデルにキーが重複して設定されています")
            End If

            '更新対象のオブジェクトの検索
            Dim keyName As String = key.First.p.Name
            Dim dic As Dictionary(Of String, String) = AbskintoneModel.GetPropertyToDefaultDic() 'プロパティ名を変換
            Dim query As String = If(dic.ContainsKey(keyName), dic(keyName), keyName) + " in "
            Dim params As New List(Of String)
            For Each obj As T In objs
                params.Add("""" + key.First.p.GetValue(obj, Nothing).ToString + """")
            Next
            query += "(" + String.Join(",", params) + ")"

            Dim list As List(Of T) = Nothing

            If objs.Count > ReadLimit Then
                list = FindAll(Of T)(query)
            Else
                list = Find(Of T)(query)
            End If

            'Update/Create分を振り分け
            Dim creates As New List(Of T)
            Dim updates As New List(Of T)

            For Each tgt As T In objs
                Dim tgtKey As Object = key.First.p.GetValue(tgt, Nothing)
                Dim sameKey As List(Of T) = (From x As T In list Where tgtKey = key.First.p.GetValue(x, Nothing) Select x).ToList
                If sameKey.Count = 0 Then
                    creates.Add(tgt)
                ElseIf sameKey.Count = 1 Then
                    tgt.record_id = sameKey.First.record_id '更新のためにidをセット
                    updates.Add(tgt)
                Else
                    Throw New kintoneException("kintone上にキーとして指定された項目(" + keyName + ")で値が重複するデータが存在します")
                End If
            Next

            'Update/Createの実行
            Dim result As Boolean = True
            Dim ids As List(Of String) = Nothing
            If updates.Count > 0 Then
                result = Update(Of T)(updates)
            End If
            If creates.Count > 0 Then
                ids = Create(Of T)(creates)
            End If

            Return ids

        End Function

        ''' <summary>
        ''' レコードの削除を行う(単一)
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="id"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Delete(Of T As AbskintoneModel)(ByVal id As String) As Boolean
            Return taskDelete(Of T)(New List(Of String) From {id})
        End Function

        ''' <summary>
        ''' レコードの削除を行う(複数件)
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="ids"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Delete(Of T As AbskintoneModel)(ByVal ids As List(Of String)) As Boolean
            Dim result As List(Of Boolean) = executeParallel(Of String, Boolean)(ids, AddressOf taskDelete(Of T))
            Dim falseIndex As Integer = result.IndexOf(False)
            Return If(falseIndex < 0, True, False)

        End Function

        ''' <summary>
        ''' 削除処理の実体
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="ids"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function taskDelete(Of T As AbskintoneModel)(ByVal ids As List(Of String)) As Boolean

            Dim request As HttpWebRequest = makeHeader("records", "DELETE")
            Dim sendData As New kintoneIds()
            sendData.app = Me.AppId
            sendData.ids = ids

            'Request Body にJSONデータを書き込み
            Dim js As New JavaScriptSerializer
            Dim kc As New kintoneContentConvertor(Of T)
            js.RegisterConverters(New List(Of JavaScriptConverter) From {kc})

            Dim jsonData As String = js.Serialize(sendData)
            writeToRequestBody(request, jsonData)

            Dim kerror As kintoneError = Nothing
            Using response As HttpWebResponse = getResponse(request, kerror)
            End Using

            If kerror IsNot Nothing Then Throw New kintoneException(kerror)
            Return True

        End Function

        Private Sub writeToRequestBody(ByRef request As HttpWebRequest, ByVal body As String)

            Using reqStream As IO.Stream = request.GetRequestStream
                Using bodyWriter As New StreamWriter(reqStream)
                    bodyWriter.Write(body)
                    bodyWriter.Flush()
                    bodyWriter.Close()
                End Using
            End Using

        End Sub

        ''' <summary>
        ''' ファイルのアップロードを行う(単一)<br/>
        ''' </summary>
        ''' <param name="file"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function UploadFile(ByVal file As HttpPostedFile) As String
            Return UploadFile(file.ToHttpPostedFileBase)
        End Function

        ''' <summary>
        ''' ファイルのアップロードを行う(単一) <br/>
        ''' HttpPostedFileは非常に特殊な型でWeb上でファイルアップロードを行っている以外の場合は使用しにくいため、上位クラスのHttpPostedFileBaseを引数にとるメソッドを用意
        ''' </summary>
        ''' <param name="file"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' <example>
        ''' <para>
        ''' PostedFile型を使用することで通常のFileStreamを処理できます
        ''' <code>
        '''   Dim file As PostedFile = New PostedFile("C:\temp\xxxx.PNG"))
        '''   kintoneAPI.UploadFile(file)
        ''' </code>
        ''' </para>
        ''' </example>
        ''' </remarks>
        Public Shared Function UploadFile(ByVal file As HttpPostedFileBase) As String
            Return UploadFile(New ReadOnlyCollection(Of HttpPostedFileBase)(New List(Of HttpPostedFileBase) From {file}))
        End Function

        ''' <summary>
        ''' ファイルのアップロードを行う(複数件)<br/>
        ''' kintone上、複数ファイルをアップロードしてもキーは単一になる。このためアップロードに使用したキーとファイルダウンロードのキーは異なるので注意
        ''' </summary>
        ''' <param name="files"></param>
        ''' <returns></returns>
        ''' <remarks> 
        ''' </remarks>
        Public Shared Function UploadFile(ByVal files As ReadOnlyCollection(Of HttpPostedFileBase)) As String

            Const FORM_NAME As String = "file" 'fileという名前でないとエラーになる
            Dim fileKey As String = ""
            Dim boundary As String = Guid.NewGuid().ToString.Replace("-", "")
            Dim reqBegin As String = "--" + boundary
            Dim reqEnd As String = vbCrLf + reqBegin + "--" + vbCrLf
            Dim postData As New MemoryStream

            If files.Count = 0 OrElse files.Count = files.Where(Function(f) String.IsNullOrEmpty(f.FileName)).Count Then
                Return String.Empty
            End If

            Dim request As HttpWebRequest = makeHeader("file", "POST")
            request.ContentType = String.Format(request.ContentType, boundary) 'boundaryをセット

            Dim js As New JavaScriptSerializer
            Dim result As Boolean = True

            For Each f As HttpPostedFileBase In files
                If String.IsNullOrEmpty(f.FileName) Then
                    Continue For
                End If
                Dim sb As New StringBuilder
                sb.Append(reqBegin + vbCrLf)
                sb.Append(String.Format("Content-Disposition: form-data; name=""{0}""; filename=""{1}""", FORM_NAME, f.FileName) + vbCrLf)
                sb.Append("Content-Type: application/octet-stream" + vbCrLf)
                sb.Append("Content-Transfer-Encoding: binary" + vbCrLf)
                sb.Append(vbCrLf)

                Dim header As Byte() = ApiEncoding.GetBytes(sb.ToString)

                postData.Write(header, 0, header.Length)
                Dim fileStream As New MemoryStream
                f.InputStream.Position = 0
                f.InputStream.CopyTo(fileStream)
                postData.Write(fileStream.ToArray, 0, fileStream.ToArray.Length)


            Next
            postData.Write(ApiEncoding.GetBytes(reqEnd), 0, reqEnd.Length)

            Dim postDataArray As Byte() = postData.ToArray
            request.GetRequestStream.Write(postDataArray, 0, postDataArray.Length)

            Dim kerror As kintoneError = Nothing
            Using response = getResponse(request, kerror)
                If Not response Is Nothing Then
                    Dim reader As New StreamReader(response.GetResponseStream)
                    Dim serialized As kintoneFile = js.Deserialize(Of kintoneFile)(reader.ReadToEnd)
                    fileKey = serialized.fileKey
                ElseIf kerror Is Nothing Then 'エラーは発生していないが、何らかの理由でresponseがNothingの場合
                    kerror = New kintoneError
                    kerror.message = "予期しないエラーが発生しました"
                End If
            End Using

            If kerror IsNot Nothing Then Throw New kintoneException(kerror)

            Return fileKey

        End Function

        ''' <summary>
        ''' ファイルのアップロードを行う(複数件)
        ''' </summary>
        ''' <param name="files"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function UploadFile(ByVal files As ReadOnlyCollection(Of HttpPostedFile)) As String
            Return UploadFile(files.ToHttpPostedFileBaseList)
        End Function

        ''' <summary>
        ''' ファイルのダウンロードを行う
        ''' </summary>
        ''' <param name="fileKey"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DownloadFile(ByVal fileKey As String) As MemoryStream

            Dim result As New MemoryStream
            Dim request As HttpWebRequest = makeHeader("file", "GET", "fileKey=" + fileKey)
            Dim kerror As kintoneError = Nothing
            Using response As HttpWebResponse = getResponse(request, kerror)
                If Not response Is Nothing Then
                    response.GetResponseStream.CopyTo(result)
                Else
                    result = Nothing
                End If
            End Using

            If kerror IsNot Nothing Then Throw New kintoneException(kerror)

            Return result

        End Function

        ''' <summary>
        ''' Response取得のための共通関数<br/>
        ''' 返却値のHttpWebResponseはUsing/End Usingセクションで扱い、使用後破棄されるようにしないと
        ''' <a href="http://stackoverflow.com/questions/12513078/system-net-webrequest-timeout-error">Timeoutが発生することがある</a>ため注意<br/>
        ''' Responseの破棄を確実にするため、例外をthrowせずエラー処理は呼び出し側に委譲する<br/>
        ''' </summary>
        ''' <param name="request"></param>
        ''' <param name="kerror"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Private Shared Function getResponse(ByVal request As HttpWebRequest, ByRef kerror As kintoneError) As HttpWebResponse

            Dim response As HttpWebResponse = Nothing
            Dim result As Boolean = True

            Try
                response = request.GetResponse
            Catch ex As WebException
                result = False
                response = Nothing
                kerror = New kintoneError
                kerror.message = ex.Message + vbCrLf + ex.StackTrace

                If ex.Response IsNot Nothing Then
                    Try
                        Dim responseStr As String = ""
                        Using res As HttpWebResponse = ex.Response
                            Using reader As New StreamReader(res.GetResponseStream)
                                responseStr = reader.ReadToEnd
                            End Using
                        End Using

                        Dim js As New JavaScriptSerializer
                        js.RegisterConverters({New kintoneErrorConvertor})
                        Dim serialized As kintoneError = js.Deserialize(Of kintoneError)(responseStr)
                        kerror = serialized

                    Catch resEx As Exception
                        '今のところ何もなし
                    End Try
                End If

            Finally
                'リクエストを閉じる(要否不明だが・・)
                request.ServicePoint.CloseConnectionGroup(request.ConnectionGroupName)
            End Try

            If result Then
                Return response
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' 並列処理のための共通関数
        ''' </summary>
        ''' <typeparam name="T">モデルの型</typeparam>
        ''' <typeparam name="R">各処理での返り値の型</typeparam>
        ''' <param name="objs">モデルの配列</param>
        ''' <param name="executor">スレッドで使用する関数</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function executeParallel(Of T, R)(ByVal objs As List(Of T), ByVal executor As Func(Of List(Of T), R))
            Dim tasks As New List(Of Task(Of R))
            Dim result As New List(Of R)

            If objs.Count > RECORD_LIMIT Then Throw New kintoneException("レコード件数が多すぎます")

            Dim splited = Enumerable.Range(0, (objs.Count + (ExecuteLimit - 1)) \ ExecuteLimit) _
                          .Select(Function(i) objs.Skip(i * ExecuteLimit).Take(ExecuteLimit).ToList) _
                          .ToList

            For Each target In splited
                Dim tgt As New Task(Of R)(Function()
                                              Return executor(target)
                                          End Function)
                tasks.Add(tgt)
            Next

            '並列処理実行
            Dim taskRuns As Task(Of R)() = tasks.ToArray
            Dim kex As kintoneException = Nothing

            Try
                Array.ForEach(taskRuns, Sub(tx) tx.Start())
                Task.WaitAll(taskRuns)

                '処理結果マージ
                For Each tx In taskRuns
                    If tx.IsCompleted Then
                        result.Add(tx.Result)
                    End If
                Next

            Catch ae As AggregateException
                Dim ex As AggregateException = ae.Flatten
                kex = New kintoneException(ex.Message, ex)
                Dim kerror As kintoneError = ex.InnerExceptions.Where(Function(x) TypeOf x Is kintoneException).Select(Function(x) CType(x, kintoneException).error).FirstOrDefault
                kex.error = kerror
            Catch ex As Exception
                kex = New kintoneException(ex.Message, ex)
            End Try

            If kex IsNot Nothing Then Throw kex

            Return result

        End Function


    End Class


End Namespace
