
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
        ''' 例：Datetime型の場合、日付時刻(DATETIME)/日付(YMD)/時刻(TIME)で3タイプあるため明示的に指定するなど
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


    End Class

End Namespace
