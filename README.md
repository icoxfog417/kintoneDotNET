kintoneDotNET (v1.0.4)
=============

kintoneDotNETは、[kintone API](https://developers.cybozu.com/ja/kintone-api/common-appapi.html) を.NET Framework上で扱うためのライブラリです。  
[Download from Nuget](https://www.nuget.org/packages/kintoneDotNET)  
.NET Frameworkは4.0以上で動作します。

## Feature
詳細な仕様はこちらをご参照ください。[How to use](https://github.com/icoxfog417/kintoneDotNET/wiki/How-to-use-kintoneDotNET)・[API Document](http://icoxfog417.github.io/kintoneDotNET/Index.html)  

特徴的な機能は、以下になります。 

### LINQライクなFind
kintone からデータを抽出する際、LINQのような構文で条件を指定することができます。  

```
List<BookModel> books = BookModel.Find<BookModel>(x => x.title Like "harry potter");
List<BookModel> latest = BookModel.Find<BookModel>(x => x.updated_time > DateTime.Now);
```

```
Dim books As List(Of BookModel) = BookModel.Find(Of BookModel)(Function(x) x.title Like "harry potter")
Dim latest As List(Of BookModel) = BookModel.Find(Of BookModel)(Function(x) x.updated_time > DateTime.Now)
```

Order、offset/limitを指定する場合、下記のようにメソッドチェーンで簡易に書けます。

```
List<BookModel> books = BookModel.Find<BookModel>().Where(x => x.price > 1000).Descending("price");
List<BookModel> popular = BookModel.Find<BookModel>().Ascending("rank").Offset(0).Limit(10);
```

```
Dim books As List(Of BookModel) = BookModel.Find(Of BookModel)().Where(Function(x) x.price > 1000).Descending("price")
Dim popular As List(Of BookModel) = BookModel.Find(Of BookModel)().Ascending("rank").Offset(0).Limit(10)
```

### Save
事前にModelクラスにKeyを設定しておくことで、Save処理を行えます(キーに一致するレコードがあればUpdate/なければInsert)。  

```
BookModel book = new BookModel();
book.title = "How to use kintone";
BookModel.Save();
```

```
Dim book As New BookModel()
book.title = "How to use kintone"
BookModel.Save()
```

なお、リビジョンの指定/取得にも対応済みです。

```
Dim index As kintoneIndex = book.Save
Console.Write("record_id=" + index.id + ",revision=" + index.revision.ToString())
```

```
kintoneIndex index = book.Save;
Console.Write("record_id=" + index.id + ",revision=" + index.revision.ToString());
```

### Bulk Process
kintone apiの上限設定値を超えるレコードの読み取り/更新が可能です。  
※内部的には、複数回APIを呼び出すことで処理しています。あまり多すぎてもkintone側へ負荷をかけてしまうので、一応60000件をリミットとして設定しています。

```
List<BookModel> books = BookModel.FindAll<BookModel>(x => x.created_time > DateTime.Now.AddMonths(-1));
books.ForEach(x => x.isOnSale = true);
BookModel.Save(books);
```

```
Dim books As List(Of BookModel) = BookModel.FindAll(Of BookModel)(Function(x) x.created_time > DateTime.Now.AddMonths(-1))
books.ForEach(Function(x) x.isOnSale = True)
BookModel.Save(books)
```

### Deal with Specific Field
kintone上の特殊なフィールドについても対応を行っています。

* 内部テーブル
* ルックアップ
* ファイル

※なお、C#のコードはVB.NETから変換しているので、誤りがあるかもしれません。もし発見した場合はご連絡ください。

## About Test Code
単体テストコードを実行するには、kintoneのアカウント取得とテストに使用するアプリの登録が必要になります。  
kintoneのアカウントは、[開発者ライセンスの申し込み](https://cybozudev.zendesk.com/hc/ja)で取得することができます。  
その後の設定方法については、[こちら](https://github.com/icoxfog417/kintoneDotNET/wiki/How-to-Run-Test-Code)をご参照ください。

