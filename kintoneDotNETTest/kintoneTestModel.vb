Imports kintoneDotNET.API
Imports kintoneDotNET.API.Types
Imports System.Configuration

Public Class kintoneTestModel
    Inherits AbskintoneModel

    <kintoneItem(isUpload:=False, isKey:=True)>
    Public Overrides Property record_id As String = ""

    <kintoneItem>
    Public Property methodinfo As String = ""
    <kintoneItem>
    Public Property stringField As String = ""
    <kintoneItem>
    Public Property numberField As Decimal = 0
    <kintoneItem(InitialValue:="radio1")>
    Public Property radio As String = ""
    <kintoneItem>
    Public Property checkbox As New List(Of String)
    <kintoneItem(InitialValue:="drop1")>
    Public Property dropdown As String = ""
    <kintoneItem>
    Public Property textarea As String = ""
    <kintoneItem>
    Public Property editor As String = ""
    <kintoneItem>
    Public Property multiselect As New List(Of String)
    <kintoneItem>
    Public Property dateField As DateTime = Nothing
    <kintoneItem(FieldType:=kintoneDatetime.TimeType)>
    Public Property time As DateTime = Nothing
    <kintoneItem(FieldType:=kintoneDatetime.DateTimeType)>
    Public Property datetimeField As DateTime = Nothing
    <kintoneItem>
    Public Property link As String = ""
    <kintoneItem>
    Public Property attachfile As New List(Of kintoneFile)

    <kintoneItem>
    Public Property validationFld As String = "ABCDE"

    <kintoneItem>
    Public Property changeLogs As New List(Of ChangeLog)

    Private _app As String = ConfigurationManager.AppSettings("testAppId")
    Public Overrides ReadOnly Property app As String
        Get
            Return _app
        End Get
    End Property

    Public Sub New()
    End Sub
    Public Sub New(ByVal methodInfo As String)
        Me.methodinfo = methodInfo
    End Sub

    Public Overrides Function ToString() As String
        Dim result As String = methodinfo + ":" + stringField
        Return result
    End Function

End Class

Public Class ChangeLog
    Inherits kintoneSubTableItem

    <kintoneItem>
    Public Property changeYMD As DateTime = Nothing

    <kintoneItem>
    Public Property historyDesc As String = ""

    Public Sub New()
    End Sub

    Public Sub New(changeYMD As DateTime, historyDesc As String)
        Me.changeYMD = changeYMD
        Me.historyDesc = historyDesc
    End Sub

End Class
