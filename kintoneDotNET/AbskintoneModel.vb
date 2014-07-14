Imports System.Reflection
Imports kintoneDotNET.API.Types
Imports System.Linq.Expressions

Namespace API

    ''' <summary>
    ''' kintoneのレコードに対応するモデルの基となる、抽象クラス
    ''' </summary>
    ''' <remarks>
    ''' 読込は全てのプロパティに対して行われるが、kintoneへ送信するのはkintoneItemAttribute属性が付与されており、isUpload=Trueのもの(デフォルトTrue)<br/>
    ''' このため、kintone側で更新したいプロパティについては&lt;kintoneItem()&gt;を付与する。<br/>
    ''' リスト型のデータ(チェックボックスリストや添付ファイルなど)については、List(Of )で宣言を行う必要あり
    ''' </remarks>
    Public MustInherit Class AbskintoneModel

        ''' <summary>
        ''' kintoneのアプリケーションID
        ''' </summary>
        Public MustOverride ReadOnly Property app As String

        ''' <summary>
        ''' [共通]レコード番号
        ''' </summary>
        Public Overridable Property record_id As String = ""

        ''' <summary>
        ''' [共通]登録時刻
        ''' </summary>
        <kintoneItem(FieldType:=kintoneDatetime.DateTimeType, isUpload:=False)>
        Public Overridable Property created_time As DateTime = DateTime.MinValue

        ''' <summary>
        ''' [共通]更新時刻
        ''' </summary>
        <kintoneItem(FieldType:=kintoneDatetime.DateTimeType, isUpload:=False)>
        Public Overridable Property updated_time As DateTime = DateTime.MinValue

        ''' <summary>
        ''' [共通]登録者
        ''' </summary>
        Public Overridable Property create_usr As New kintoneUser()

        ''' <summary>
        ''' [共通]更新者
        ''' </summary>
        Public Overridable Property update_usr As New kintoneUser()

        ''' <summary>
        ''' [共通]ステータス (※プロセスを使っている場合値を取得可能)
        ''' </summary>
        Public Overridable Property status As String = ""

        ''' <summary>
        ''' [共通]作業者
        ''' </summary>
        Public Overridable Property work_usr As New kintoneUser()

        ''' <summary>
        ''' [共通]リビジョン番号<br/>
        ''' 初期値は-1(この場合、送信してもkintone側で検証は行われない)
        ''' </summary>
        Public Overridable Property revision As Integer = -1

        ''' <summary>更新時、リビジョンを無視して更新する(デフォルトFalse)</summary>
        Public Property IgnoreRevision As Boolean = False

        '$idは条件指定では使えないため、リクエストを送る場合は"レコード番号"を利用する必要がある
        Private _convertDictionary As New List(Of NameConvertor) From {
            NameConvertor.Create("$id", "record_id", NameConvertor.Direction.Read),
            NameConvertor.Create("レコード番号", "record_id", NameConvertor.Direction.Send),
            NameConvertor.Create("$revision", "revision", NameConvertor.Direction.Read),
            NameConvertor.Create("作成日時", "created_time"),
            NameConvertor.Create("更新日時", "updated_time"),
            NameConvertor.Create("作成者", "create_usr"),
            NameConvertor.Create("更新者", "update_usr"),
            NameConvertor.Create("ステータス", "status"),
            NameConvertor.Create("作業者", "work_usr")
        }

        ''' <summary>
        ''' デフォルトの日本語項目名称を変換するためのディクショナリを取得<br/>
        ''' レコード番号などのデフォルト項目の項目名を変更している場合、このプロパティをオーバーライドしてください
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Protected Overridable Property ConvertDictionary As List(Of NameConvertor)
            Get
                Return _convertDictionary
            End Get
            Set(value As List(Of NameConvertor))
                _convertDictionary = value
            End Set
        End Property


        ''' <summary>
        ''' kintone上でのレコードURL
        ''' </summary>
        Public ReadOnly Property record_show_url() As String
            Get
                Dim api As New kintoneAPI(app)
                Dim url As String = "https://" + kintoneAPI.Host + "/k/" + app + "/show?record=" + record_id
                Return url
            End Get
        End Property

        ''' <summary>
        ''' レコードの検索を行う(expression指定)
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="expression">Booleanを返却する関数式</param>
        ''' <param name="isConvert">デフォルトの項目変換をかけるか否か</param>
        ''' <returns></returns>
        ''' <remarks>
        ''' kintone上ではレコード番号などのデフォルト項目をrecord_id等でなく「レコード番号」と日本語そのままでもっているため、この名前で検索を行わないとエラーになります<br/>
        ''' isConvert=True(デフォルト値)としておけば、record_id->レコード番号といったデフォルト項目の変換を自動で行ってくれます。<br/>
        ''' <example>
        ''' <para>
        ''' <code>
        '''   'AbskintoneModelを継承して作成したBookModelを使用し、検索を行う
        '''   Dim list AS List(Of BookModel) = BookModel.Find(Of BookModel)(Function(x) x.title Like "Mathematics" And x.price &lt; 3000 )
        '''   
        '''   'Attributeを設定しておけば、日付型の条件指定もDateTime型オブジェクトから直接行えます
        '''   Dim upds AS List(Of BookModel) = BookModel.Find(Of BookModel)(Function(x) x.updated_time >= DateTime.Now)
        ''' </code>
        ''' </para>
        ''' </example>
        ''' </remarks>
        Public Shared Function Find(Of T As AbskintoneModel)(ByVal expression As Expression(Of Func(Of T, Boolean)), _
                                                                Optional ByVal isConvert As Boolean = True) As List(Of T)
            Dim model As T = Activator.CreateInstance(Of T)()
            Dim query As String = kintoneQueryExpression.Eval(Of T)(expression, If(isConvert, model.GetToItemNameDic(), Nothing))
            Return Find(Of T)(query)
        End Function

        ''' <summary>
        ''' レコードの検索を行う(文字列クエリ指定)<br/>
        ''' ※クエリを作成する際、title="hoge"と文字列型の場合""で比較値を囲う必要がある点に注意してください<br/>
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="query">任意のクエリ文字列</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Find(Of T As AbskintoneModel)(ByVal query As String) As List(Of T)
            Return GetAPI(Of T).Find(Of T)(query)
        End Function

        ''' <summary>
        ''' 直接条件を指定せず、クエリオブジェクトから検索を行う
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Find(Of T As AbskintoneModel)() As kintoneQuery(Of T)
            Dim q As New kintoneQuery(Of T)(False)
            Return q
        End Function

        ''' <summary>
        ''' id指定によるレコードの検索
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="id"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function FindById(Of T As AbskintoneModel)(ByVal id As String) As T
            Dim result As List(Of T) = Find(Of T)(Function(x) x.record_id = id)
            If result IsNot Nothing AndAlso result.Count > 0 Then
                Return result.First
            Else
                Return Nothing
            End If
        End Function

        ''' <summary>
        ''' idの複数指定によるレコードの検索
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="ids"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function FindByIds(Of T As AbskintoneModel)(ByVal ids As List(Of String)) As List(Of T)
            Dim result As List(Of T) = Nothing
            If ids Is Nothing Then Return Nothing

            If ids.Count > kintoneAPI.ReadLimit Then
                result = FindAll(Of T)(Function(x) ids.Contains(x.record_id))
            Else
                result = Find(Of T)(Function(x) ids.Contains(x.record_id))
            End If
            Return result
        End Function

        ''' <summary>
        ''' レコードの検索を行う(全件)(expression指定)<br/>
        ''' kintone APIの上限値を超える件数のレコードを取得します
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="expression">Booleanを返却する関数式</param>
        ''' <param name="isConvert">デフォルトの項目変換をかけるか否か</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function FindAll(Of T As AbskintoneModel)(ByVal expression As Expression(Of Func(Of T, Boolean)), _
                                                                Optional ByVal isConvert As Boolean = True) As List(Of T)
            Dim model As T = Activator.CreateInstance(Of T)()
            Dim query As String = kintoneQueryExpression.Eval(Of T)(expression, If(isConvert, model.GetToItemNameDic, Nothing))
            Return FindAll(Of T)(query)
        End Function

        ''' <summary>
        ''' レコードの検索を行う(全件)(文字列クエリ指定)<br/>
        ''' kintone APIの上限値を超える件数のレコードを取得します
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="query"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function FindAll(Of T As AbskintoneModel)(ByVal query As String) As List(Of T)
            Return GetAPI(Of T).FindAll(Of T)(query)
        End Function

        ''' <summary>
        ''' 直接条件を指定せず、クエリオブジェクトから検索を行う
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function FindAll(Of T As AbskintoneModel)() As kintoneQuery(Of T)
            Dim q As New kintoneQuery(Of T)(True)
            Return q
        End Function

        ''' <summary>
        ''' レコードの登録を行う<br/>
        ''' 登録を行った後、登録を行ったレコードをkintone上から取得し返却します
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="objs">登録対象オブジェクト</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Create(Of T As AbskintoneModel)(ByVal objs As List(Of T)) As List(Of T)
            Dim indexes As kintoneIndexes = GetAPI(Of T)().BulkCreate(Of T)(objs)
            Return FindByIds(Of T)(indexes.ids)
        End Function

        ''' <summary>
        ''' レコードの更新を行う
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="objs"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Update(Of T As AbskintoneModel)(ByVal objs As List(Of T)) As kintoneIndexes
            Return GetAPI(Of T).BulkUpdate(Of T)(objs)
        End Function

        ''' <summary>
        ''' レコードの保存を行う<br/>
        ''' モデル上 isKey = True と設定された項目をキーとし、一致するキーがある場合はUpdate、なければCreateを行う
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="objs"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Save(Of T As AbskintoneModel)(ByVal objs As List(Of T)) As List(Of T)
            Dim indexes As kintoneIndexes = GetAPI(Of T)().BulkSave(Of T)(objs)
            Return FindByIds(Of T)(indexes.ids)

        End Function

        ''' <summary>
        ''' レコードの削除を行う
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="ids"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function Delete(Of T As AbskintoneModel)(ByVal ids As List(Of String)) As Boolean
            Return GetAPI(Of T).BulkDelete(Of T)(ids)
        End Function

        Public Shared Function Delete(Of T As AbskintoneModel)(ByVal objs As List(Of T)) As Boolean
            Dim model As T = Activator.CreateInstance(Of T)()
            objs = model.UpdateHook(objs) 'idをセット

            Dim ids As List(Of String) = (From x As T In objs Where Not String.IsNullOrEmpty(x.record_id) Select x.record_id).ToList
            Dim result As Boolean = GetAPI(Of T).BulkDelete(Of T)(ids)

            If result Then '成功した場合、objsに設定されていたid/リビジョンをクリアする(削除されたため)
                objs.ForEach(Sub(x)
                                 x.record_id = String.Empty
                                 x.revision = -1
                             End Sub)
            End If

            Return result

        End Function

        ''' <summary>
        ''' レコードの登録(単一)を行う<br/>
        ''' 自身をkintone上に登録します
        ''' </summary>
        Public Function Create() As kintoneIndex
            Dim result As Object = execute("Create", Me)
            Return setkintoneIndex(result)
        End Function

        ''' <summary>
        ''' レコードの更新(単一)を行う<br/>
        ''' 自身のレコードを更新します
        ''' </summary>
        Public Function Update() As kintoneIndex
            Dim result As Object = execute("Update", Me)
            Return setkintoneIndex(result)
        End Function

        ''' <summary>
        ''' レコードの保存を行う<br/>
        ''' isKey = True で設定されたレコードがある場合更新、なければ登録を行います。
        ''' </summary>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Function Save() As kintoneIndex
            Dim result As Object = execute("Save", Me)
            Return setkintoneIndex(result)

        End Function

        ''' <summary>
        ''' オブジェクトをkintoneIndexにキャストし、値を自身にコピーする
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function setkintoneIndex(ByVal obj As Object) As kintoneIndex
            Dim typedResult As New kintoneIndex
            If obj IsNot Nothing AndAlso TypeOf obj Is kintoneIndex Then
                typedResult = CType(obj, kintoneIndex)
                Me.record_id = typedResult.id
                Me.revision = typedResult.revision
            End If
            Return typedResult
        End Function

        ''' <summary>
        ''' レコードの削除(単一)を行う<br/>
        ''' 自身のIDに一致するレコードを削除します
        ''' </summary>
        Public Function Delete() As Boolean

            If String.IsNullOrEmpty(Me.record_id) Then
                'レコードidの設定がない場合、keyからidの取得を試みる
                execute("SetIndexToModel", Me)
            End If

            If String.IsNullOrEmpty(Me.record_id) Then Return True '既に削除されている場合、Trueを返却

            Dim result As Boolean = CBool(execute("Delete", Me.record_id))
            If result Then '削除に成功したら、id/リビジョンをクリアする
                Me.record_id = String.Empty
                Me.revision = -1
            End If

            Return result

        End Function

        ''' <summary>
        ''' 実体である自身のタイプでkintoneAPIのジェネリクスメソッドをコールする
        ''' </summary>
        ''' <param name="methodName"></param>
        ''' <param name="params"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Private Function execute(ByVal methodName As String, ParamArray params As Object()) As Object
            Dim method As MethodInfo = (From m As MethodInfo In GetType(kintoneAPI).GetMethods Where m.Name = methodName).FirstOrDefault
            Dim generic As MethodInfo = method.MakeGenericMethod(Me.GetType)
            Dim result As Object = Nothing

            Try
                result = generic.Invoke(GetAPI, params)
            Catch ex As System.Reflection.TargetInvocationException
                'リフレクションによる呼び出しの場合本来の例外が内部に隠蔽されるため、取り出し
                Throw ex.InnerException
            End Try

            Return result

        End Function


        ''' <summary>
        ''' Create処理実行前に行われる処理<br/>
        ''' 事前に行っておくべき処理(値設定/対象の追加・削除)があればここに実装する
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="objs">Create対象オブジェクト</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overridable Function CreateHook(Of T As AbskintoneModel)(ByVal objs As List(Of T)) As List(Of T)
            Return objs
        End Function

        ''' <summary>
        ''' Update処理実行前に行われる処理<br/>
        ''' 事前に行っておくべき処理(値設定/対象の追加・削除)があればここに実装する
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="objs">>Update対象オブジェクト</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overridable Function UpdateHook(Of T As AbskintoneModel)(ByVal objs As List(Of T)) As List(Of T)

            Dim ids As List(Of String) = (From x As T In objs Select x.record_id).ToList
            Dim idline As String = String.Join("", ids)

            If String.IsNullOrEmpty(idline) Then 'レコードidの設定がない場合、keyからidの取得を試みる
                GetAPI(Of T).SetIndexToModels(Of T)(objs)
            End If

            Return objs

        End Function

        ''' <summary>
        ''' Delete処理実行前に行われる処理<br/>
        ''' 事前に行っておくべき処理(値設定/対象の追加・削除)があればここに実装する
        ''' </summary>
        ''' <typeparam name="T"></typeparam>
        ''' <param name="ids">Delete対象id</param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Overridable Function DeleteHook(Of T As AbskintoneModel)(ByVal ids As List(Of String)) As List(Of String)
            Return ids
        End Function


        ''' <summary>
        ''' 自身を操作するAPIを取得する
        ''' </summary>
        Private Function GetAPI() As kintoneAPI
            Dim api As New kintoneAPI(app)
            Return api
        End Function

        ''' <summary>
        ''' 自身を操作するAPIを取得する(Shared Method用)
        ''' </summary>
        Private Shared Function GetAPI(Of T As AbskintoneModel)() As kintoneAPI
            Dim model As T = Activator.CreateInstance(Of T)()
            Dim api As New kintoneAPI(model.app)
            Return api
        End Function

        ''' <summary>
        ''' kintone上の項目名称("レコード番号","作成日時" など)をプロパティ名(record_id,updated_time etc)に変換するためのDictionaryを取得する
        ''' </summary>
        Public Function GetToPropertyDic() As Dictionary(Of String, String)
            Return GetNameConvertDic(NameConvertor.Direction.Read)
        End Function

        ''' <summary>
        ''' プロパティ名をkintone上の項目名称に変換するためのDictionaryを取得する
        ''' </summary>
        Public Function GetToItemNameDic() As Dictionary(Of String, String)
            Return GetNameConvertDic(NameConvertor.Direction.Send)
        End Function

        ''' <summary>
        ''' 変換用Dictionaryを取得するための内部処理
        ''' </summary>
        ''' <param name="direction"></param>
        Private Function GetNameConvertDic(ByVal direction As NameConvertor.Direction) As Dictionary(Of String, String)
            Dim nconvertors = _convertDictionary.FindAll(Function(cn) cn.ConvertDirection = direction Or cn.ConvertDirection = NameConvertor.Direction.Both)

            If direction = NameConvertor.Direction.Read Then
                Return nconvertors.ToDictionary(Function(cn) cn.ItemName, Function(cn) cn.PropertyName)
            Else
                Return nconvertors.ToDictionary(Function(cn) cn.PropertyName, Function(cn) cn.ItemName)
            End If

        End Function

    End Class

    ''' <summary>
    ''' kintone上のアイテム名とコード上のプロパティ名を対応させるための変換ルール
    ''' </summary>
    ''' <remarks></remarks>
    Public Class NameConvertor
        Public Enum Direction
            Both
            Read
            Send
        End Enum

        Public Property ItemName As String = ""
        Public Property PropertyName As String = ""
        Public Property ConvertDirection As Direction = Direction.Both

        Public Sub New(ByVal itemName As String, ByVal propertyName As String, Optional ByVal direction As Direction = Direction.Both)
            Me.ItemName = itemName
            Me.PropertyName = propertyName
            Me.ConvertDirection = direction
        End Sub

        Public Shared Function Create(ByVal itemName As String, ByVal propertyName As String, Optional ByVal direction As Direction = Direction.Both) As NameConvertor
            Return New NameConvertor(itemName, propertyName, direction)
        End Function

    End Class

End Namespace
