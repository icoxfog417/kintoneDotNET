Imports kintoneDotNET.API
Imports kintoneDotNET.API.Types
Imports System.Configuration

Public Class kintoneTestModel
    Inherits AbskintoneModel

    <UploadTarget>
    Public Property methodinfo As String = ""
    <UploadTarget>
    Public Property stringField As String = ""
    <UploadTarget>
    Public Property numberField As Decimal = 0
    <UploadTarget(InitialValue:="radio1")>
    Public Property radio As String = ""
    <UploadTarget>
    Public Property checkbox As New List(Of String)
    <UploadTarget(InitialValue:="drop1")>
    Public Property dropdown As String = ""
    <UploadTarget>
    Public Property textarea As String = ""
    <UploadTarget>
    Public Property editor As String = ""
    <UploadTarget>
    Public Property multiselect As New List(Of String)
    <UploadTarget>
    Public Property dateField As DateTime = Nothing
    <UploadTarget(FieldType:="TIME")>
    Public Property time As DateTime = Nothing
    <UploadTarget(FieldType:="DATETIME")>
    Public Property datetimeField As DateTime = Nothing
    <UploadTarget>
    Public Property link As String = ""
    <UploadTarget>
    Public Property attachfile As New List(Of kintoneFile)

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
