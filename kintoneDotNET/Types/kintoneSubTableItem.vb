
Namespace API.Types

    ''' <summary>
    ''' 内部テーブル内の項目を表現するクラス。継承して使用
    ''' </summary>
    ''' <remarks></remarks>
    Public MustInherit Class kintoneSubTableItem
        Implements IkintoneType

        <UploadTarget()>
        Public Property id As String = ""

    End Class

End Namespace
