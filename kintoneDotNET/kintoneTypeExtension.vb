Imports System.Runtime.CompilerServices
Imports System.Collections.ObjectModel
Imports kintoneDotNET.API.Types
Imports System.Linq.Expressions

Namespace API

    ''' <summary>
    ''' データ型の性質上List型となる項目について、メソッド拡張を行う
    ''' </summary>
    ''' <remarks></remarks>
    Public Module kintoneTypeExtension

        <Extension()>
        Public Sub UploadFile(ByVal item As List(Of kintoneFile), ByVal file As HttpPostedFile)
            UploadFile(item, file.ToHttpPostedFileBase)
        End Sub

        <Extension()>
        Public Sub UploadFile(ByVal item As List(Of kintoneFile), ByVal files As ReadOnlyCollection(Of HttpPostedFile))
            UploadFile(item, files.ToHttpPostedFileBaseList)
        End Sub

        <Extension()>
        Public Sub UploadFile(ByVal item As List(Of kintoneFile), ByVal file As HttpPostedFileBase)
            UploadFile(item, New ReadOnlyCollection(Of HttpPostedFileBase)(New List(Of HttpPostedFileBase) From {file}))
        End Sub

        <Extension()>
        Public Sub UploadFile(ByVal item As List(Of kintoneFile), ByVal files As ReadOnlyCollection(Of HttpPostedFileBase))
            Dim fKey As String = kintoneAPI.UploadFile(files)
            If Not String.IsNullOrEmpty(fKey) Then
                item.Add(New kintoneFile(fKey))
            End If
        End Sub

        <Extension()>
        Public Function ToHttpPostedFileBaseList(ByVal files As ReadOnlyCollection(Of HttpPostedFile)) As ReadOnlyCollection(Of HttpPostedFileBase)
            Dim baseList As New ReadOnlyCollection(Of HttpPostedFileBase)( _
                (From f As HttpPostedFile In files
                 Select f.ToHttpPostedFileBase).ToList)
            Return baseList
        End Function

        <Extension()>
        Public Function ToHttpPostedFileBase(ByVal file As HttpPostedFile) As HttpPostedFileBase
            Return New HttpPostedFileWrapper(file)
        End Function

    End Module

End Namespace
