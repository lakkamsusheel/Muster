Imports System.IO
Public Class LetterGenerator
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.LetterGenerator
    '   Letters Generator Class
    '
    ' Copyright (C) 2004 CIBER, Inc.
    ' All rights reserved.
    '
    ' Release   Initials    Date      Description
    '  1.0        ??      ??/??/??    Original class definition.
    '  1.1        JVC2    02/08/2005  Integrated calls to new ProfileData
    '  1.2        AN      02/10/2005  Integrated new AppFlags Object
    '-------------------------------------------------------------------------------

    Protected DOC_PATH As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_SystemGenerated).ProfileValue & "\"
    Protected TmpltPath As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_Templates).ProfileValue & "\"
    Protected SketchPath As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_Sketches).ProfileValue & "\"

    Protected frmCaller As Form

    Protected Sub New()
    End Sub

    Protected Sub New(ByVal frmTemp As Form)
        frmCaller = frmTemp
    End Sub
    'To check the given file Exist or Not
    Friend Shared Function FileExists(ByVal FilePath As String) As Boolean
        Dim file As File
        Try
            If file.Exists(FilePath) Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    ' To Kill the Word Application
    Friend Shared Sub KillWordApp(ByRef WordApp As Word.Application)
        Try
            If Not WordApp Is Nothing Then
                WordApp.Quit(False)
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Friend Shared Function CopyTemplatesFromServerToLocal(ByVal ServerPath As String, ByVal path As String)

        Dim dir As New DirectoryInfo(ServerPath)
        Dim TempDir As Directory
        Dim files() As FileInfo
        Dim i As Integer = 0

        If Not TempDir.Exists(path) Then
            TempDir.CreateDirectory(path)
        End If
        ' Copy the Templates From Server to local Path
        files = dir.GetFiles("*.Doc")
        For i = 0 To files.Length - 1
            If FileExists(path + "\" + files(i).Name) Then
                File.Delete(path + "\" + files(i).Name)
                files(i).CopyTo(path + "\" + files(i).Name)
            Else
                files(i).CopyTo(path + "\" + files(i).Name)
            End If
        Next
        dir = Nothing

    End Function

    Public Shared Sub ViewDocument(ByVal DocumentLocation As String)
        Dim SrcDoc As Word.Document
        Dim WordApp As Word.Application

        Try

            WordApp = UIUtilsGen.GetWordApp

            If Not WordApp Is Nothing Then
                WordApp.Visible = True

                If File.Exists(DocumentLocation) Then

                    SrcDoc = WordApp.Documents.Open(DocumentLocation)
                Else
                    MsgBox("This particular document cannot be found. You may need to regenerate this document and enter in the proper date of orginal generation", MsgBoxStyle.OKOnly, _
                          String.Format("Doc File: {0}", DocumentLocation))
                End If
            End If

            WordApp = Nothing

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception(ex.Message))
            MyErr.ShowDialog()
        Finally
            'Delay()
            UIUtilsGen.Delay(, 1)
            SrcDoc = Nothing
        End Try

    End Sub
End Class
