Namespace API

    ''' <summary>
    ''' kintoneアカウント情報を格納するためのクラス。
    ''' コンストラクタが読み込むapp.configの設定に対して優先使用されます。
    ''' </summary>
    ''' <remarks></remarks>
    Public Class kintoneAccount

        ''' <summary>
        ''' アプリケーションのドメイン
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Domain As String = ""

        ''' <summary>
        ''' Basic認証のためのID
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property AccessId As String = ""

        ''' <summary>
        ''' Basic認証のためのパスワード
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property AccessPassword As String = ""

        ''' <summary>
        ''' kintoneのログインを行うためのID
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property LoginId As String = ""

        ''' <summary>
        ''' kintoneのログインを行うためのパスワード
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property LoginPassword As String = ""

        ''' <summary>
        ''' kintoneのアプリごとに生成するAPIトークン。
        ''' 設定されている場合、ログインID/パスワードより優先される。
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ApiToken As String = ""

        ''' <summary>
        ''' Proxyを経由する場合のアドレス
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property Proxy As String = ""

        ''' <summary>
        ''' 認証が必要なプロキシの場合のユーザー名
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ProxyUser As String = ""

        ''' <summary>
        ''' 認証が必要なプロキシの場合のパスワード
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property ProxyPassword As String = ""

        ''' <summary>
        ''' kintoneアカウント情報を格納するためのクラス。
        ''' コンストラクタが読み込むapp.configの設定に対して優先使用されます。
        ''' </summary>
        ''' <param name="domain">アプリケーションのドメイン</param>
        ''' <param name="accessId">Basic認証のためのID</param>
        ''' <param name="accessPassword">Basic認証のためのパスワード</param>
        ''' <param name="loginId">kintoneのログインを行うためのID</param>
        ''' <param name="loginPassword">kintoneのログインを行うためのパスワード</param>
        ''' <param name="apiToken">kintoneのアプリごとに生成するAPIトークン。
        ''' 設定されている場合、ログインID/パスワードより優先される。</param>
        ''' <param name="proxy">Proxyを経由する場合のアドレス</param>
        ''' <param name="proxyUser">認証が必要なプロキシの場合のユーザー名</param>
        ''' <param name="proxyPassword">認証が必要なプロキシの場合のパスワード</param>
        ''' <remarks></remarks>
        Public Sub New(Optional ByVal domain As String = "",
                       Optional ByVal accessId As String = "",
                       Optional ByVal accessPassword As String = "",
                       Optional ByVal loginId As String = "",
                       Optional ByVal loginPassword As String = "",
                       Optional ByVal apiToken As String = "",
                       Optional ByVal proxy As String = "",
                       Optional ByVal proxyUser As String = "",
                       Optional ByVal proxyPassword As String = "")

            For Each config As String In ConfigurationManager.AppSettings.AllKeys
                Dim value As String = ConfigurationManager.AppSettings(config)
                Select Case config
                    Case "ktDomain"
                        _Domain = If(Not String.IsNullOrEmpty(domain), domain, value)
                    Case "ktAccessId"
                        _AccessId = If(Not String.IsNullOrEmpty(accessId), accessId, value)
                    Case "ktAccessPassword"
                        _AccessPassword = If(Not String.IsNullOrEmpty(accessPassword), accessPassword, value)
                    Case "ktLoginId"
                        _LoginId = If(Not String.IsNullOrEmpty(loginId), loginId, value)
                    Case "ktLoginPassword"
                        _LoginPassword = If(Not String.IsNullOrEmpty(loginPassword), loginPassword, value)
                    Case "ktApiToken"
                        _ApiToken = If(Not String.IsNullOrEmpty(apiToken), apiToken, value)
                    Case "proxy"
                        _Proxy = If(Not String.IsNullOrEmpty(proxy), proxy, value)
                    Case "proxyUser"
                        _ProxyUser = If(Not String.IsNullOrEmpty(proxyUser), proxyUser, value)
                    Case "proxyPassword"
                        _ProxyPassword = If(Not String.IsNullOrEmpty(proxyPassword), proxyPassword, value)
                End Select
            Next

        End Sub

    End Class

End Namespace
