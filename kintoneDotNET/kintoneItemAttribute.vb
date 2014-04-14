
Namespace API

    ''' <summary>
    ''' kintoneへの更新対象項目を示すアトリビュート
    ''' </summary>
    ''' <remarks></remarks>
    <AttributeUsage(AttributeTargets.Property)>
    Public Class kintoneItemAttribute
        Inherits Attribute
        Public Sub New()
        End Sub

        ''' <summary>
        ''' kintone上のタイプを明示的に指定する。
        ''' 1.Datetime型の場合に、日付時刻(DATETIME)/日付(YMD)/時刻(TIME)で3タイプあるため明示的に指定する
        ''' 2.Integer型の場合で、リビジョンを指定するフィールドであることを明示
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property FieldType As String = ""

        ''' <summary>
        ''' 値がNothingの場合の初期値(空白文字列も対象となる。String.isNullOrEmptyの場合設定値が適用される)
        ''' ラジオボタンなど、値の設定が必須となる場合に使用する
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property InitialValue As Object = Nothing

        ''' <summary>
        ''' 更新対象か否かのフラグ。デフォルトTrue
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property isUpload As Boolean = True

        ''' <summary>
        ''' 重複を禁止するキーである項目に付与<br/>
        ''' Saveを行う際、データの存在判定に使用される
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property isKey As Boolean = False


    End Class

End Namespace
