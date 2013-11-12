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
    ''' kintoneのREST APIをコールするためのクラス
    ''' https://developers.cybozu.com/ja/kintone-api/appapi-requestjsapi.html
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneAPI

        Protected Const KINTONE_HOST As String = "{0}.cybozu.com"
        Protected Const KINTONE_PORT As String = "443"
        Protected Const KINTONE_API_FORMAT As String = "https://{0}/k/v1/{1}.json"
        Public Const KINTONE_READ_LIMIT As Integer = 100 '100件がMAX

        Private _error As New kintoneError
        Public ReadOnly Property GetError As kintoneError
            Get
                Return _error
            End Get
        End Property

        Public Shared ReadOnly Property ApiEncoding As Encoding
            Get
                Return Encoding.GetEncoding("UTF-8") 'kintoneのエンコードはUTF-8
            End Get
        End Property

        Public Shared ReadOnly Property Domain As String
            Get
                Dim d As String = ConfigurationManager.AppSettings("ktDomain")
                Return d
            End Get
        End Property

        Public Shared ReadOnly Property Host As String
            Get
                Return String.Format(KINTONE_HOST, Domain)
            End Get
        End Property

        Private _appId As String
        Public ReadOnly Property AppId() As String
            Get
                Return _appId
            End Get
        End Property

        Private Shared _readLimit As Integer = 0
        Public Shared Property ReadLimit As Integer
            Get
                If _readLimit < 1 Then
                    Return KINTONE_READ_LIMIT '設定がない場合、既定上限を返却
                Else
                    Return _readLimit
                End If
            End Get
            Set(value As Integer)
                _readLimit = value
            End Set
        End Property


        ''' <summary>
        ''' Basic認証のためのキーを作成する。形式などについては公式ドキュメントを参照
        ''' http://developers.cybozu.com/ja/kintone-api/common-appapi.html
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property AccessKey As String
            Get
                Dim id As String = ConfigurationManager.AppSettings("ktAccessId")
                Dim password As String = ConfigurationManager.AppSettings("ktAccessPassword")
                Dim key As String = Convert.ToBase64String(ApiEncoding.GetBytes(id + ":" + password))
                Return key
            End Get
        End Property

        ''' <summary>
        ''' kintone上のログインを行うためのキーを作成する
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared ReadOnly Property LoginKey As String
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
        Public Shared ReadOnly Property Proxy As WebProxy
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
        ''' kintoneのアプリケーションIDを引数とする。
        ''' 引数として受け取ったアプリケーションに対して、レコードの登録や更新などの操作を行う。アプリケーションIDは、URLから判断可能。
        ''' https://xxx.cybozu.com/k/{アプリケーションID}/
        ''' </summary>
        ''' <param name="appId">kintoneのアプリケーションID</param>
        ''' <remarks></remarks>
        Public Sub New(ByVal appId As String)
            Me._appId = appId
        End Sub

        ''' <summary>
        ''' kintoneにアクセスするためのHttpヘッダを作成する
        ''' </summary>
        ''' <param name="command">record/records、fileなどのREST APIの名称</param>
        ''' <param name="method">POST/GETなど</param>
        ''' <param name="query">クエリ引数</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' </remarks>
        Protected Shared Function makeHttpHeader(ByVal command As String, ByVal method As String, Optional ByVal query As String = "") As HttpWebRequest
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
        ''' 指定された型でデータの抽出を行う
        ''' ※このメソッドは、kintoneのレコード数上限までしか取得を行いません。全件取得する場合はFindAllを使用してください
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="query">
        ''' queryの形式については、kintone上の公式ドキュメントを参照
        ''' https://developers.cybozu.com/ja/kintone-api/apprec-readapi.html
        ''' item="xxxx"のように、値はダブルクウォーテーションで囲う必要があるため要注意
        ''' </param>
        ''' <param name="kerror"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Find(Of T As AbskintoneModel)(ByVal query As String, Optional ByRef kerror As kintoneError = Nothing) As List(Of T)
            Dim model As T = Activator.CreateInstance(Of T)()
            Dim q As String = "app=" + model.app + If(Not String.IsNullOrEmpty(query), "&query=" + HttpUtility.UrlEncode(query), String.Empty)

            Return FindBase(Of T)(q, kerror)

        End Function

        ''' <summary>
        ''' レコード上限を超えたレコードの抽出を行う
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="query"></param>
        ''' <param name="kerror"></param>
        ''' <returns></returns>
        ''' <remarks>orderは考慮されないため、データ取得後LINQ等で並び替えを行ってください</remarks>
        Public Shared Function FindAll(Of T As AbskintoneModel)(ByVal query As String, Optional ByRef kerror As kintoneError = Nothing) As List(Of T)
            Dim model As T = Activator.CreateInstance(Of T)()
            Dim q As String = "app=" + model.app + If(Not String.IsNullOrEmpty(query), "&query=" + HttpUtility.UrlEncode(query), "&query=") 'limitを指定する必要があるので、条件指定がなくてもqueryパラメータは付与

            Dim stepCount As Integer = 1
            Dim offset As Integer = 0
            Dim recordStillExist As Boolean = True
            Dim threadCount As Integer = 4 'TODO:並列実行数は様子を見ながら調整
            Dim result As New List(Of T)

            While recordStillExist
                Dim tasks As New List(Of Task(Of List(Of T)))
                For i As Integer = 1 To threadCount
                    tasks.Add(createTask(Of T)(q, offset))
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
                            If Not tx.Result Is Nothing AndAlso tx.Result.Count > 0 Then
                                result.AddRange(tx.Result)
                            End If
                            If tx.Result.Count < ReadLimit Then 'リミットのサイズより取得結果が少なくなれば、もう取得する必要がないため停止
                                recordStillExist = False
                            End If
                        End If
                    Next

                Catch ex As Exception
                    kerror = New kintoneError
                    kerror.message = ex.Message
                    recordStillExist = False '例外が発生した場合終了する
                End Try

                If stepCount > 150 Then '無限ループを回避するための保険 threadCount * 150 * ReadLimit (大体 4 * 150 * 100 = 60000件)　で無理ならあきらめたほうがいい
                    recordStillExist = False
                End If

                stepCount += 1

            End While

            Return result

        End Function

        Private Shared Function createTask(Of T As AbskintoneModel)(baseQuery As String, offset As Integer) As Task(Of List(Of T))

            Dim task As New Task(Of List(Of T))(Function()
                                                    Dim query = baseQuery + HttpUtility.UrlEncode(" limit " + ReadLimit.ToString + " offset " + offset.ToString)
                                                    Dim localError As kintoneError = Nothing
                                                    Dim list As List(Of T) = FindBase(Of T)(query, localError)
                                                    If Not localError Is Nothing Then
                                                        Throw New Exception(localError.message) 'TODO:kintoneErrorを格納できる独自例外があったほうがいいかも
                                                    End If
                                                    Return list
                                                End Function
                                                )
            Return task
        End Function

        Private Shared Function FindBase(Of T As AbskintoneModel)(ByVal query As String, Optional ByRef kerror As kintoneError = Nothing) As List(Of T)

            Dim serialized As kintoneRecords(Of T) = Nothing
            Dim result As New List(Of T)

            Dim request As HttpWebRequest = makeHttpHeader("records", "GET", query)
            Using response As HttpWebResponse = getResponse(request, kerror)
                If Not response Is Nothing Then
                    Dim js As New JavaScriptSerializer
                    Dim kc As New kintoneContentConvertor(Of T)
                    js.RegisterConverters(New List(Of JavaScriptConverter) From {kc})
                    Dim reader As New StreamReader(response.GetResponseStream)
                    serialized = js.Deserialize(Of kintoneRecords(Of T))(reader.ReadToEnd)

                    For Each record As T In serialized.records
                        result.Add(record)
                    Next

                End If
            End Using

            Return result

        End Function

        ''' <summary>
        ''' レコード登録を行う
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Create(Of T As AbskintoneModel)(ByVal obj As T) As List(Of String)
            'JsonオブジェクトをRequestに書き出し
            'http://stackoverflow.com/questions/9145667/how-to-post-json-to-the-server

            Dim request As HttpWebRequest = makeHttpHeader("records", "POST")
            Dim sendData As New kintoneRecords(Of T)
            sendData.app = AppId
            sendData.records.Add(obj)

            'Request Body にJSONデータを書き込み
            Dim js As New JavaScriptSerializer
            Dim kc As New kintoneContentConvertor(Of T)
            js.RegisterConverters(New List(Of JavaScriptConverter) From {kc})

            Dim jsonData As String = js.Serialize(sendData)

            Using bodyWriter As New StreamWriter(request.GetRequestStream)
                bodyWriter.Write(jsonData)
                bodyWriter.Flush()
                bodyWriter.Close()
            End Using

            Dim list As New List(Of String)
            Using response As HttpWebResponse = getResponse(request, _error)
                If Not response Is Nothing Then
                    Dim reader As New StreamReader(response.GetResponseStream)
                    Dim serialized As kintoneIds = js.Deserialize(Of kintoneIds)(reader.ReadToEnd)
                    list = serialized.ids
                End If
            End Using

            Return list

        End Function

        ''' <summary>
        ''' レコードの更新を行う
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Update(Of T As AbskintoneModel)(ByVal obj As T) As Boolean
            Dim request As HttpWebRequest = makeHttpHeader("records", "PUT")
            Dim sendData As New kintoneRecords(Of kintoneRecord(Of T))
            sendData.app = AppId
            sendData.records.Add(New kintoneRecord(Of T)(obj))

            'Request Body にJSONデータを書き込み
            Dim js As New JavaScriptSerializer
            Dim kc As New kintoneContentConvertor(Of T)
            js.RegisterConverters(New List(Of JavaScriptConverter) From {kc})

            Dim jsonData As String = js.Serialize(sendData)

            Using bodyWriter As New StreamWriter(request.GetRequestStream)
                bodyWriter.Write(jsonData)
                bodyWriter.Flush()
                bodyWriter.Close()
            End Using

            Using response As HttpWebResponse = getResponse(request, _error)
            End Using

            Return Not isError()

        End Function

        ''' <summary>
        ''' レコードの削除を行う
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="ids"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Delete(Of T As AbskintoneModel)(ByVal ids As List(Of String)) As Boolean

            Dim request As HttpWebRequest = makeHttpHeader("records", "DELETE")
            Dim sendData As New kintoneIds()
            sendData.app = Me.AppId
            sendData.ids = ids

            'Request Body にJSONデータを書き込み
            Dim js As New JavaScriptSerializer
            Dim kc As New kintoneContentConvertor(Of T)
            js.RegisterConverters(New List(Of JavaScriptConverter) From {kc})

            Dim jsonData As String = js.Serialize(sendData)

            Using bodyWriter As New StreamWriter(request.GetRequestStream)
                bodyWriter.Write(jsonData)
                bodyWriter.Flush()
                bodyWriter.Close()
            End Using

            Using response As HttpWebResponse = getResponse(request, _error)
            End Using

            Return Not isError()

        End Function


        ''' <summary>
        ''' ファイルのアップロードを行う
        ''' </summary>
        ''' <param name="files"></param>
        ''' <param name="kerror"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' 複数ファイルをアップロードしても、キーは単一になる模様。アップロードに使用したキーとファイルダウンロードのキーは異なるので注意
        ''' </remarks>
        Public Shared Function UploadFile(ByVal files As ReadOnlyCollection(Of HttpPostedFileBase), Optional ByRef kerror As kintoneError = Nothing) As String

            Const FORM_NAME As String = "file" 'fileという名前でないとエラーになる
            Dim fileKey As String = ""
            Dim boundary As String = Guid.NewGuid().ToString.Replace("-", "")
            Dim reqBegin As String = "--" + boundary
            Dim reqEnd As String = vbCrLf + reqBegin + "--" + vbCrLf
            Dim postData As New MemoryStream

            If files.Count = 0 OrElse files.Count = files.Where(Function(f) String.IsNullOrEmpty(f.FileName)).Count Then
                Return String.Empty
            End If

            Dim request As HttpWebRequest = makeHttpHeader("file", "POST")
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

            Using response = getResponse(request, kerror)
                If Not response Is Nothing Then
                    Dim reader As New StreamReader(response.GetResponseStream)
                    Dim serialized As kintoneFile = js.Deserialize(Of kintoneFile)(reader.ReadToEnd)
                    fileKey = serialized.fileKey
                End If
            End Using

            Return fileKey

        End Function

        Public Shared Function UploadFile(ByVal file As HttpPostedFile, Optional ByRef kerror As kintoneError = Nothing) As String
            Return UploadFile(file.ToHttpPostedFileBase, kerror)
        End Function


        Public Shared Function UploadFile(ByVal files As ReadOnlyCollection(Of HttpPostedFile), Optional ByRef kerror As kintoneError = Nothing) As String
            Return UploadFile(files.ToHttpPostedFileBaseList, kerror)
        End Function

        Public Shared Function UploadFile(ByVal file As HttpPostedFileBase, Optional ByRef kerror As kintoneError = Nothing) As String
            Return UploadFile(New ReadOnlyCollection(Of HttpPostedFileBase)(New List(Of HttpPostedFileBase) From {file}), kerror)
        End Function

        ''' <summary>
        ''' ファイルのダウンロードを行う
        ''' </summary>
        ''' <param name="fileKey"></param>
        ''' <param name="kerror"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function DownloadFile(ByVal fileKey As String, Optional ByRef kerror As kintoneError = Nothing) As MemoryStream

            Dim result As New MemoryStream
            Dim request As HttpWebRequest = makeHttpHeader("file", "GET", "fileKey=" + fileKey)
            Using response As HttpWebResponse = getResponse(request, kerror)
                If Not response Is Nothing Then
                    response.GetResponseStream.CopyTo(result)
                Else
                    result = Nothing
                End If
            End Using

            Return result

        End Function

        Public Function isError() As Boolean
            Return isError(_error)
        End Function
        Public Shared Function isError(ByVal kerror As kintoneError) As Boolean
            If kerror Is Nothing OrElse String.IsNullOrEmpty(kerror.message) Then
                Return False
            Else
                Return True
            End If
        End Function

        ''' <summary>
        ''' レスポンス取得のための共通関数
        ''' </summary>
        ''' <param name="request"></param>
        ''' <param name="kerror"></param>
        ''' <returns></returns>
        ''' <remarks>
        ''' 返却値のHttpWebResponseはUsing/End Usingセクションで扱い、使用後破棄されるようにしないとTimeoutが発生することがあるため注意
        ''' http://stackoverflow.com/questions/12513078/system-net-webrequest-timeout-error
        ''' </remarks>
        Private Shared Function getResponse(ByVal request As HttpWebRequest, ByRef kerror As kintoneError) As HttpWebResponse

            Dim response As HttpWebResponse = Nothing
            Dim js As New JavaScriptSerializer
            Dim result As Boolean = True

            Try
                response = request.GetResponse
            Catch ex As WebException
                result = False
                response = ex.Response
                kerror = New kintoneError
                kerror.message = ex.Message + vbCrLf + ex.StackTrace
            End Try

            If result Then
                Return response
            Else
                If Not response Is Nothing Then
                    Dim reader As New StreamReader(response.GetResponseStream)
                    Dim serialized As kintoneError = js.Deserialize(Of kintoneError)(reader.ReadToEnd)
                    If Not serialized Is Nothing Then
                        kerror = serialized
                    End If
                End If

                Return Nothing
            End If
        End Function

    End Class


End Namespace
