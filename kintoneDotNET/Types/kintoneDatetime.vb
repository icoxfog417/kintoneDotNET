
Namespace API.Types

    ''' <summary>
    ''' kintone上の日付データ型と合わせるための変換を行うクラス
    ''' このクラスは型宣言するためでなく、変換を行うために存在する(日付型は通常通りDatetime型で宣言する)
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneDatetime

        ''' <summary>
        ''' kintone上の日付の初期値を取得する
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks>
        ''' kintoen上のDatetimeは"1,000年以上9,999年以下の日付でなければなりません。"
        ''' UTC時刻に換算すると1000を下回ってしまう場合に備え、時刻は23:59:59とする
        ''' </remarks>
        Public Shared ReadOnly Property InitialValue As DateTime
            Get
                Return New DateTime(1000, 1, 1, 23, 59, 59)
            End Get
        End Property

        ''' <summary>
        ''' kintone上の時刻をDateTimeオブジェクトに変換する
        ''' </summary>
        ''' <param name="obj"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function kintoneToDatetime(ByVal obj As Object) As DateTime
            Dim d As DateTime = Nothing

            'kintone上はUTC時間なので、それを考慮
            '※なお、UTCは日付/時刻には適用されない
            'https://developers.cybozu.com/ja/kintone-api/common-appapi.html#i-5

            If DateTime.TryParseExact(obj("value"), "yyyy-MM-ddTHH:mm:ssZ", System.Globalization.DateTimeFormatInfo.InvariantInfo, _
                                      System.Globalization.DateTimeStyles.AssumeUniversal, d) Then
                Return d
            ElseIf DateTime.TryParseExact(obj("value"), "yyyy-MM-dd", System.Globalization.DateTimeFormatInfo.InvariantInfo, _
                                      System.Globalization.DateTimeStyles.None, d) Then
                '日付型
                Return d
            ElseIf DateTime.TryParseExact(obj("value"), "HH:mm", System.Globalization.DateTimeFormatInfo.InvariantInfo, _
                                      System.Globalization.DateTimeStyles.None, d) Then
                '時刻型
                Return d
            Else
                Return Nothing
            End If

        End Function

        ''' <summary>
        ''' DateTime型変数をkintone上の時刻フィールドに変換する
        ''' </summary>
        ''' <param name="vdate"></param>
        ''' <param name="format"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function toKintoneDate(ByVal vdate As DateTime, Optional ByVal format As String = "") As String

            Dim result As String = ""
            Dim value As DateTime = vdate
            If value.Equals(DateTime.MinValue) Then
                value = InitialValue
            End If

            Select Case format
                Case "DATETIME"
                    result = value.ToString("yyyy-MM-ddTHH:mm:ss")
                    Dim utc As DateTime = value.ToUniversalTime()
                    Dim span As TimeSpan = value - utc 'UTC時刻どの差分を計算し末尾に付与する
                    If span >= TimeSpan.Zero Then result += "+" Else result += "-"
                    result += span.ToString("hh\:mm")
                Case "TIME"
                    result = value.ToString("HH:mm:ss")
                Case Else 'デフォルト日付として扱う
                    'kintone上はハイフン区切りのStringになる
                    result = value.ToString("yyyy-MM-dd")
            End Select

            Return result

        End Function

        ''' <summary>
        ''' テキスト文字列からDateTime型に変換するユーティリティ関数
        ''' </summary>
        ''' <param name="text"></param>
        ''' <param name="format"></param>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Shared Function toDatetime(ByVal text As String, ByVal format As String) As DateTime
            Dim d As DateTime = DateTime.Now
            If DateTime.TryParseExact(text, format, _
                      System.Globalization.DateTimeFormatInfo.InvariantInfo, _
                    System.Globalization.DateTimeStyles.None, d) Then
                Return d
            Else
                Return d
            End If
        End Function

    End Class

End Namespace

