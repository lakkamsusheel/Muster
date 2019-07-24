Imports System.Drawing
Imports System.Text
Imports System
Imports System.Drawing.Design

'-------------------------------------------------------------------------------
' MUSTER.BusinessLogic.pReportParams
'   Provides the operations required to Generate Letters.
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date        Description
'   1.0         MR      03/28/05    Original Class Implementation.
'   1.1         MNR     07/25/05    Added function to handle inspection letters
'   2.0   Thomas Franey 03/09/09    Added line of code to replace VB's vbcrlf with ANSI carriage return to remove blocks from docs
'
'
' Function          Description
' CreateLetter()   Generic Function to Generate Letters based on Input Collection.
'
'-------------------------------------------------------------------------------
Imports System.IO
Imports Microsoft.Office.Interop

Namespace MUSTER.BusinessLogic
    <Serializable()> _
    Public Class pLetterGen

#Region "WordTemplate Class"
        <Serializable()> _
        Public Class WordTemplate

            Private data(23) As Object

            <NonSerialized()> Public path As String

            Sub New(ByVal Template As Word.ListTemplate, ByVal thisPath As String)

                path = thisPath

                Try

                    data(0) = Template.ListLevels.Item(1).LinkedStyle()
                    data(1) = Template.ListLevels.Item(1).NumberFormat
                    data(2) = Template.ListLevels.Item(1).NumberPosition
                    data(3) = Template.ListLevels.Item(1).NumberStyle
                    data(4) = Template.ListLevels.Item(1).StartAt
                    data(5) = Template.ListLevels.Item(1).TrailingCharacter
                    data(6) = Template.ListLevels.Item(1).TabPosition
                    data(7) = Template.ListLevels.Item(1).TextPosition
                Catch
                End Try

                Try
                    data(8) = Template.ListLevels.Item(2).LinkedStyle()
                    data(9) = Template.ListLevels.Item(2).NumberFormat
                    data(10) = Template.ListLevels.Item(2).NumberPosition
                    data(11) = Template.ListLevels.Item(2).NumberStyle
                    data(12) = Template.ListLevels.Item(2).StartAt
                    data(13) = Template.ListLevels.Item(2).TrailingCharacter
                    data(14) = Template.ListLevels.Item(2).TabPosition
                    data(15) = Template.ListLevels.Item(2).TextPosition
                Catch
                End Try


                Try
                    data(16) = Template.ListLevels.Item(3).LinkedStyle()
                    data(17) = Template.ListLevels.Item(3).NumberFormat
                    data(18) = Template.ListLevels.Item(3).NumberPosition
                    data(19) = Template.ListLevels.Item(3).NumberStyle
                    data(20) = Template.ListLevels.Item(3).StartAt
                    data(21) = Template.ListLevels.Item(3).TrailingCharacter
                    data(22) = Template.ListLevels.Item(3).TabPosition
                    data(23) = Template.ListLevels.Item(3).TextPosition
                Catch
                End Try


            End Sub


            Public Sub ExtractTemplateDatafromRecord(ByRef Template As Word.ListTemplate, ByVal hasMultipleFacilities As Boolean)

                Try

                    Template.ListLevels.Item(1).NumberStyle = IIf(hasMultipleFacilities, Word.WdListNumberStyle.wdListNumberStyleNone, Word.WdListNumberStyle.wdListNumberStyleLowercaseLetter)
                    Template.ListLevels.Item(1).LinkedStyle = data(0)
                    Template.ListLevels.Item(1).NumberFormat = IIf(hasMultipleFacilities, data(1), data(1) + ".")
                    Template.ListLevels.Item(1).NumberPosition = data(2)
                    Template.ListLevels.Item(1).StartAt = data(4)
                    Template.ListLevels.Item(1).TrailingCharacter = data(5)
                    Template.ListLevels.Item(1).TabPosition = data(6)
                    Template.ListLevels.Item(1).TextPosition = data(7)
                Catch
                End Try


                Try

                    Template.ListLevels.Item(2).NumberStyle = IIf(hasMultipleFacilities, Word.WdListNumberStyle.wdListNumberStyleLowercaseLetter, data(19))
                    Template.ListLevels.Item(2).LinkedStyle() = IIf(hasMultipleFacilities, data(8), data(16))
                    If hasMultipleFacilities Then
                        Template.ListLevels.Item(2).NumberFormat = data(9) + "."
                    Else
                        Template.ListLevels.Item(2).NumberFormat = data(17)
                    End If

                    Template.ListLevels.Item(2).NumberPosition = IIf(hasMultipleFacilities, data(10), data(18))
                    Template.ListLevels.Item(2).StartAt = IIf(hasMultipleFacilities, data(12), data(20))
                    Template.ListLevels.Item(2).TrailingCharacter = IIf(hasMultipleFacilities, data(13), data(21))
                    Template.ListLevels.Item(2).TabPosition = IIf(hasMultipleFacilities, data(14), data(22))
                    Template.ListLevels.Item(2).TextPosition = IIf(hasMultipleFacilities, data(15), data(23))
                Catch
                End Try

                Try
                    Template.ListLevels.Item(3).NumberStyle = data(19)
                    Template.ListLevels.Item(3).LinkedStyle() = data(16)
                    Template.ListLevels.Item(3).NumberFormat = data(17)
                    Template.ListLevels.Item(3).NumberPosition = data(18)
                    Template.ListLevels.Item(3).StartAt = data(20)
                    Template.ListLevels.Item(3).TrailingCharacter = data(21)
                    Template.ListLevels.Item(3).TabPosition = data(22)
                    Template.ListLevels.Item(3).TextPosition = data(23)
                Catch
                End Try

            End Sub



            Public Shared Sub ExtractTemplateFromFile(ByRef temp As WordTemplate, ByVal path As String)
                Dim objStream As New FileStream(String.Format("{0}\ListTemplate.dat", path), FileMode.Open)
                Dim objFormatter As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
                Dim newTemp As Word.ListTemplate

                ' Serialization exception will be raised because

                ' objHandler is not serializable - note

                ' we did not intend to serialize objHandler

                temp = objFormatter.Deserialize(objStream)


                objFormatter = Nothing
                objStream.Close()
                objStream = Nothing



            End Sub




        End Class
#End Region


#Region "Public Events"
        Public Event CheckListProgress(ByVal percent As Single)
        Public Event CloseCheckListProgress()
#End Region
#Region "Private Member Variables"
        Private strModuleID As String
        Private strLetter_To_Print As String
        'Private WordApp As Word.Application
        Private Const FILENAME As Object = "normal.dot"
        Private Const NEWTEMPLATE As Object = False
        Private Const DOCTYPE As Object = 0
        Private Const ISVISIBLE As Object = True
        Private Const [READONLY] As Object = True
        Private missing As Object = System.Reflection.Missing.Value
        Private DestDoc As Word.Document
        Private SrcDoc As Word.Document

        Private Structure InspPrepGuidelineVars
            Dim P_LDhasElecALLD As Boolean
            Dim T_LDhasGroundWaterOrVaporMonitoring As Boolean
            Dim P_LDhasGroundWaterOrVaporMonitoring As Boolean
            Dim T_LDhasAutomaticTankGauging As Boolean
            Dim T_LDhasStatisticalInventory As Boolean
            Dim P_LDhasStatisticalInventory As Boolean
            Dim T_LDhasInventoryControl As Boolean
            Dim T_hasInstDtgt12_21_88_neverInspected As Boolean
            Dim T_hasInstDtgt12_21_88_neverInspected_MatOfConstructionhasFigerGlassReinforcedPlastic As Boolean
            Dim P_LDhasLineTightnessTesting As Boolean
            Dim P_hasInstDtgt12_21_88_neverInspected As Boolean
            Dim T_LDhasManualTankGauging As Boolean
            Dim T_LDhasContinuous_Electronic_Visual_Manual_Monitoring As Boolean
            Dim P_LDhasContinuous_Electronic_Visual_Manual_Monitoring As Boolean
            Dim T_SteelTank As Boolean
            Dim T_ImpressedCurrentCP As Boolean
            Dim P_ImpressedCurrentCP As Boolean
            Dim P_Steel_MetallicPiping As Boolean
            Dim P_LDhasMechanicalAutomaticLine As Boolean
            Dim RecordsHeader As Boolean
            ' Components
            Dim T_EPGonly As Boolean
            Dim T_UsedOilonly As Boolean
            Dim P_pressurizedPipeSystem As Boolean
        End Structure
#End Region
#Region "Constructors"
        Public Sub New()
            MyBase.New()
        End Sub
#End Region
#Region "Exposed Operations"
#Region "Collection Operations"

#Region "generic (basic) letter operation"
        Public Function CreateLetter(ByVal strModuleID As String, ByVal strLetter_To_Print As String, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal strfile As String = "", Optional ByVal strSignature As String = "", Optional ByVal break As Boolean = False) ', Optional ByVal strFiles As String = "")
            Try

                Dim i As Integer = 0
                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty
                Dim strToday As String = String.Empty
                Dim strInfoNeededValue As String = String.Empty


                'Instantiate the Word Object
                LetterTemplateInit(WordApp, TemplatePath, DestinationPath)

                With WordApp

                    .Visible = False

                    With DestDoc

                        If Not colParams.Item("<ERAC Contact>, <ERAC>") Is Nothing Then
                            .Content.Find.Execute(FindText:="<ERAC Contact>, <ERAC>", ReplaceWith:="<ERAC>", Replace:=Word.WdReplace.wdReplaceAll)
                        End If


                        ' Find and Replace the TAGs with Values.
                        For i = 0 To colParams.Count - 1

                            strKey = colParams.Keys(i).ToString
                            If Not colParams.Get(strKey) Is Nothing Then

                                strValue = colParams.Item(strKey)
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                If strKey = "<Reasons>" Then
                                    If .Tables.Count > 0 Then
                                        .Tables.Item(1).Cell(1, 1).Range.Text = strValue
                                    End If
                                ElseIf strKey <> "<Cert Mail Number>" And strValue.Length < 255 Then
                                    .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                                ElseIf strKey = "<ERAC Contact>, <ERAC>" Then
                                    'ignore
                                ElseIf strKey = "<InfoNeeded>" Then

                                    strInfoNeededValue = strValue

                                Else
                                    Dim go As Boolean = True

                                    While go
                                        go = False

                                        With WordApp.Selection.Find
                                            .Text = strKey
                                            .Replacement.Text = ""
                                            .Forward = True
                                            .Wrap = Word.WdFindWrap.wdFindContinue
                                            .Execute()
                                        End With

                                        If Not WordApp.Selection.Text Is Nothing AndAlso WordApp.Selection.Text.ToUpper.Trim = strKey.ToUpper.Trim Then
                                            go = True
                                            WordApp.Selection.Text = IIf(strValue = String.Empty, "", strValue)
                                        End If

                                    End While
                                End If
                            Else
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceOne)

                            End If

                        Next

                        'Info needed Determination
                        .Content.Find.Execute(FindText:="<InfoNeeded>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)

                        If .Tables.Count > 0 AndAlso Not strInfoNeededValue Is Nothing AndAlso strInfoNeededValue.Length > 0 Then
                            .Tables.Item(1).Cell(1, 1).Range.Text = strInfoNeededValue
                        End If



                        If Not break Then
                            Dim strPhoto As String = String.Empty
                            If strfile <> String.Empty Then
                                strPhoto = UCase(strfile)
                            End If
                            If strfile <> String.Empty And strPhoto.IndexOf(UCase("\\Opcgw\MUSTER\Images\Licensees\Nophoto.gif")) < 0 Then
                                '"Z:\Images\Licensees\1002.gif"
                                'WordApp.ActiveDocument.Tables(1).Rows(1).Cells(1).Tables(1)

                                With .Tables.Item(1).Cell(1, 1).Tables.Item(1).Cell(1, 1)
                                    Dim oPic As Word.InlineShape
                                    oPic = .Range.InlineShapes.AddPicture(FileName:=strfile _
                                        , LinkToFile:=False, SaveWithDocument:=True)
                                    oPic.Height = 55
                                    oPic.Width = 45
                                End With

                                If colParams.HasKeys AndAlso Not colParams.Item("<LicenseeID>") Is Nothing AndAlso colParams.Item("<LicenseeID>").ToUpper.IndexOf("RX") > -1 Then
                                    With .Tables.Item(1).Cell(1, 1).Tables.Item(1).Cell(2, 1)
                                        Dim oPic As Word.InlineShape
                                        oPic = .Range.InlineShapes.AddPicture(FileName:=String.Format("{0}\{1}", strfile.Substring(0, strfile.LastIndexOf("\")), "RestrictedPics.bmp") _
                                            , LinkToFile:=False, SaveWithDocument:=True)
                                        oPic.Height = CInt(.Height * 0.4)
                                        oPic.Width = CInt(.Width * 0.95)
                                    End With
                                End If


                                'oPic = .InlineShapes.AddPicture(FILENAME:=strfile _
                                '        , LinkToFile:=False, SaveWithDocument:=True)
                                'oPic.Height = 70
                                'oPic.Width = 70
                            End If

                            If strSignature <> String.Empty And strSignature.IndexOf("\\Opcgw\MUSTER\Images\Licensees\NoSignature.gif") < 0 Then

                                With .Tables.Item(1).Cell(1, 1).Tables.Item(1).Cell(3, 2)
                                    Dim oPic As Word.InlineShape
                                    oPic = .Range.InlineShapes.AddPicture(FileName:=strSignature _
                                            , LinkToFile:=False, SaveWithDocument:=True)
                                    oPic.Height = 25
                                    oPic.Width = 70
                                End With

                            End If

                        End If

                    End With
                End With

                LetterTemplateSave(WordApp, TemplatePath, DestinationPath)

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End Function

        Public Function CreateLabels(ByVal strModuleID As String, ByVal strLetter_To_Print As String, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal strAddress As String = "", Optional ByVal nRow As Integer = 1, Optional ByVal nColumn As Integer = 1) ', Optional ByVal strFiles As String = "")
            Try

                Dim i As Integer = 0
                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty

                'Instantiate the Word Object
                'Me.LetterTemplateInit(WordApp, TemplatePath, DestinationPath)
                Me.TemporaryDocInit(WordApp, TemplatePath)

                With WordApp

                    .Visible = False
                    With DestDoc

                        ' Find and Replace the TAGs with Values.
                        For i = 0 To colParams.Count - 1
                            strKey = colParams.Keys(i).ToString
                            If Not colParams.Get(strKey) Is Nothing Then
                                strValue = colParams.Get(strKey).ToString
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                Dim go As Boolean = True

                                While go
                                    go = False

                                    With WordApp.Selection.Find
                                        .Text = strKey
                                        .Replacement.Text = ""
                                        .Forward = True
                                        .Wrap = Word.WdFindWrap.wdFindContinue
                                        .Execute()
                                    End With

                                    If Not WordApp.Selection.Text Is Nothing AndAlso WordApp.Selection.Text.ToUpper.Trim = strKey.ToUpper.Trim Then
                                        go = True
                                        WordApp.Selection.Text = IIf(strValue = String.Empty, "", strValue)
                                    End If

                                End While
                            Else
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)

                            End If

                        Next

                        If .Tables.Count > 0 Then
                            .Tables.Item(1).Cell(nRow, nColumn).Range.Text = strAddress
                        End If
                    End With
                End With

                ' LetterTemplateSave(WordApp, TemplatePath, DestinationPath)

            Catch ex As Exception

                Throw New Exception(ex.Message)
            End Try
        End Function

        Public Function CreateEnvelope(ByVal strLetter_To_Print As String, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByRef WordApp As Word.Application = Nothing)
            Try
                Dim DocumentPath As String
                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty
                Dim strToday As String = String.Empty
                Dim strDocName As String = String.Empty
                Dim i As Integer = 0


                'Instantiate the Word Object
                Me.TemporaryDocInit(WordApp, TemplatePath)
                'LetterTemplateInit(WordApp, TemplatePath, DestinationPath)

                With WordApp
                    With DestDoc

                        ' Find and Replace the TAGs with Values.
                        For i = 0 To colParams.Count - 1
                            strKey = colParams.Keys(i).ToString
                            If Not colParams.Get(strKey) Is Nothing Then

                                strValue = colParams.Get(strKey).ToString

                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                Dim go As Boolean = True

                                While go
                                    go = False

                                    With WordApp.Selection.Find
                                        .Text = strKey
                                        .Replacement.Text = ""
                                        .Forward = True
                                        .Wrap = Word.WdFindWrap.wdFindContinue
                                        .Execute()
                                    End With

                                    If Not WordApp.Selection.Text Is Nothing AndAlso WordApp.Selection.Text.ToUpper.Trim = strKey.ToUpper.Trim Then
                                        go = True
                                        WordApp.Selection.Text = IIf(strValue = String.Empty, "", strValue)
                                    End If

                                End While
                            Else
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)

                            End If

                        Next
                    End With
                End With

                'LetterTemplateSave(WordApp, TemplatePath, DestinationPath)

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try

        End Function

        Public Function CreateGenericLetter(ByVal strModuleID As String, ByVal strLetter_To_Print As String, ByVal colParams As Specialized.NameValueCollection, ByVal cols As Int16, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal TableFormat As Int16 = 35, Optional ByVal bApplyBorders As Boolean = False, Optional ByVal bApplyShading As Boolean = False, Optional ByVal bApplyFont As Boolean = True, Optional ByVal bApplyColor As Boolean = False, Optional ByVal bApplyHeader As Boolean = True, Optional ByVal bApplyLastRow As Boolean = False, Optional ByVal bApplyFirstCol As Boolean = False, Optional ByVal bApplyLastCol As Boolean = False, Optional ByVal bApplyAutofit As Boolean = True, Optional ByVal attachDocumentInfo As String = "")
            Try

                Dim i As Integer = 0
                Dim oPara As Word.Paragraph
                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty

                'Instantiate the Word Object
                LetterTemplateInit(WordApp, TemplatePath, DestinationPath)

                With WordApp
                    With DestDoc

                        ' Find and Replace the TAGs with Values.
                        If Not colParams.Item("<ERAC Contact>, <ERAC>") Is Nothing Then
                            .Content.Find.Execute(FindText:="<ERAC Contact>, <ERAC>", ReplaceWith:="<ERAC>", Replace:=Word.WdReplace.wdReplaceAll)
                        End If

                        For i = 0 To colParams.Count - 1

                            strKey = colParams.Keys(i).ToString
                            If Not colParams.Get(strKey) Is Nothing Then

                                strValue = colParams.Item(strKey)
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                If strKey = "<DATA>" Then
                                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                    oPara.Range.Text = strValue
                                    oPara.Range.ConvertToTable("|", , cols, , TableFormat, bApplyBorders, bApplyShading, bApplyFont, bApplyColor, bApplyHeader, bApplyLastRow, bApplyFirstCol, bApplyLastCol, bApplyAutofit)
                                    oPara.Format.SpaceAfter = 1
                                    oPara.Range.InsertParagraphAfter()

                                Else
                                    Dim go As Boolean = True

                                    While go
                                        go = False

                                        With WordApp.Selection.Find
                                            .Text = strKey
                                            .Replacement.Text = ""
                                            .Forward = True
                                            .Wrap = Word.WdFindWrap.wdFindContinue
                                            .Execute()
                                        End With

                                        If Not WordApp.Selection.Text Is Nothing AndAlso WordApp.Selection.Text.ToUpper.Trim = strKey.ToUpper.Trim Then
                                            go = True
                                            WordApp.Selection.Text = IIf(strValue = String.Empty, "", strValue)
                                        End If

                                    End While
                                End If
                            Else
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)
                            End If

                        Next

                        ' attach document
                        If attachDocumentInfo <> "" Then
                            If System.IO.File.Exists(attachDocumentInfo) Then
                                Try
                                    WordApp.Selection.GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:=.Bookmarks.Item("\endofdoc").Name)
                                    WordApp.Selection.InsertFile(FileName:=attachDocumentInfo, Range:=.Bookmarks.Item("\endofdoc").Range, ConfirmConversions:=False, Link:=False, Attachment:=False)
                                Catch ex As Exception
                                    ' do nothing
                                End Try
                            End If
                        End If

                    End With
                End With

                LetterTemplateSave(WordApp, TemplatePath, DestinationPath)

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End Function

#End Region

#Region "Registration Letter Operations"

        ' Registration Letters
        Public Function CreateOtherRegistrationLetter(ByVal ownerID As Integer, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, ByVal facsForOwner As DataTable, ByVal oOwner As MUSTER.BusinessLogic.pOwner, Optional ByRef WordApp As Word.Application = Nothing)

            Dim strKey As String = String.Empty
            Dim strValue As String = String.Empty
            Dim i As Integer
            Dim spacingCount As Integer = 0

            Try

                'Instantiate the Word Object
                Me.LetterTemplateInit(WordApp, TemplatePath, DestinationPath)


                With WordApp
                    With DestDoc

                        ' Find and Replace the TAGs with Values.
                        For i = 0 To colParams.Count - 1
                            strKey = colParams.Keys(i).ToString
                            If Not colParams.Get(strKey) Is Nothing Then

                                strValue = colParams.Get(strKey).ToString
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                Dim go As Boolean = True

                                While go
                                    go = False

                                    With WordApp.Selection.Find
                                        .Text = strKey
                                        .Replacement.Text = ""
                                        .Forward = True
                                        .Wrap = Word.WdFindWrap.wdFindContinue
                                        .Execute()
                                    End With

                                    If Not WordApp.Selection.Text Is Nothing AndAlso WordApp.Selection.Text.ToUpper.Trim = strKey.ToUpper.Trim Then
                                        go = True
                                        WordApp.Selection.Text = IIf(strValue = String.Empty, "", strValue)
                                    End If

                                End While
                            Else
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)

                            End If

                        Next

                        ' Facility Table Functionality
                        With .Tables.Item(1)

                            ' if there is more than 1 facility, use inspection format
                            ' i.e. Facility ID # xxx Facility Name Facility Address (line 1, city state)
                            If facsForOwner.Rows.Count > 1 Then

                                ' Add rows
                                For i = 0 To facsForOwner.Rows.Count - 1
                                    .Rows.Add()
                                Next

                                'fill rows
                                For i = 0 To facsForOwner.Rows.Count - 1
                                    .Cell(i + 1, 1).Range.Text = "Facility ID # " + facsForOwner.Rows(i)("ID").ToString + " " + _
                                                                facsForOwner.Rows(i)("Name").ToString + " " + _
                                                                facsForOwner.Rows(i)("Address").ToString + ", " + _
                                                                facsForOwner.Rows(i)("CITY").ToString + " " + _
                                                                facsForOwner.Rows(i)("STATE").ToString
                                Next
                            Else
                                ' Add rows
                                For i = 0 To facsForOwner.Rows.Count - 1
                                    .Rows.Add()
                                    .Rows.Add()
                                    .Rows.Add()
                                    .Rows.Add()
                                    If i < facsForOwner.Rows.Count - 1 Then .Rows.Add()
                                Next

                                'fill rows
                                For i = 0 To facsForOwner.Rows.Count - 1
                                    'With .Tables.Item(1)
                                    .Cell(i + 1 + spacingCount, 1).Range.Text = facsForOwner.Rows(i)("Name").ToString
                                    .Cell(i + 2 + spacingCount, 1).Range.Text = facsForOwner.Rows(i)("Address").ToString
                                    .Cell(i + 3 + spacingCount, 1).Range.Text = facsForOwner.Rows(i)("CITY").ToString + " " + facsForOwner.Rows(i)("STATE").ToString
                                    .Cell(i + 4 + spacingCount, 1).Range.Text = "Facility ID # " + facsForOwner.Rows(i)("ID").ToString

                                    spacingCount += 4
                                    'End With
                                Next
                            End If

                        End With

                    End With
                End With

                LetterTemplateSave(WordApp, TemplatePath, DestinationPath)

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try

        End Function

        Public Function CreateRegistrationLetter(ByVal ownerID As Integer, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, ByVal facsForOwner As DataTable, ByVal oOwner As MUSTER.BusinessLogic.pOwner, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal showTransferOwnerSection As Boolean = False)

            Dim strKey As String = String.Empty
            Dim strValue As String = String.Empty
            Dim i As Integer
            Dim showNewOwnerSection As Boolean = False
            Dim showFeeSection As Boolean = False
            Dim showTOISSection As Boolean = False
            Dim ds As DataSet

            Try

                'Instantiate the Word Object
                LetterTemplateInit(WordApp, TemplatePath, DestinationPath)

                With WordApp
                    With DestDoc


                        ' Find and Replace the TAGs with Values.
                        For i = 0 To colParams.Count - 1
                            strKey = colParams.Keys(i).ToString
                            If Not colParams.Get(strKey) Is Nothing Then
                                strValue = colParams.Get(strKey).ToString
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                Dim go As Boolean = True

                                While go
                                    go = False

                                    With WordApp.Selection.Find
                                        .Text = strKey
                                        .Replacement.Text = ""
                                        .Forward = True
                                        .Wrap = Word.WdFindWrap.wdFindContinue
                                        .Execute()
                                    End With

                                    If Not WordApp.Selection.Text Is Nothing AndAlso WordApp.Selection.Text.ToUpper.Trim = strKey.ToUpper.Trim Then
                                        go = True
                                        WordApp.Selection.Text = IIf(strValue = String.Empty, "", strValue)
                                    End If

                                End While
                            Else
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)

                            End If

                        Next

                        ' Facility Table Functionality
                        With .Tables.Item(1)

                            ' if there is more than 1 facility, use inspection format
                            ' i.e. Facility ID # xxx Facility Name Facility Address (line 1, city state)
                            If facsForOwner.Rows.Count > 1 Then
                                ' Add rows
                                For i = 0 To facsForOwner.Rows.Count - 1
                                    .Rows.Add()
                                Next

                                'fill rows
                                For i = 0 To facsForOwner.Rows.Count - 1
                                    'With .Tables.Item(1)
                                    .Cell(i + 1, 1).Range.Text = "Facility ID # " + facsForOwner.Rows(i)("ID").ToString + " " + _
                                                                facsForOwner.Rows(i)("Name").ToString + " " + _
                                                                facsForOwner.Rows(i)("Address").ToString + ", " + _
                                                                facsForOwner.Rows(i)("CITY").ToString + " " + _
                                                                facsForOwner.Rows(i)("STATE").ToString
                                    'End With
                                Next

                            Else
                                ' Add rows
                                For i = 0 To facsForOwner.Rows.Count - 1
                                    .Rows.Add()
                                    .Rows.Add()
                                    .Rows.Add()
                                    .Rows.Add()
                                    If i < facsForOwner.Rows.Count - 1 Then .Rows.Add()
                                Next

                                'fill rows
                                Dim spacingCount As Integer = 0
                                For i = 0 To facsForOwner.Rows.Count - 1
                                    .Cell(i + 1 + spacingCount, 1).Range.Text = facsForOwner.Rows(i)("Name").ToString
                                    .Cell(i + 2 + spacingCount, 1).Range.Text = facsForOwner.Rows(i)("Address").ToString
                                    .Cell(i + 3 + spacingCount, 1).Range.Text = facsForOwner.Rows(i)("CITY").ToString + " " + facsForOwner.Rows(i)("STATE").ToString
                                    .Cell(i + 4 + spacingCount, 1).Range.Text = "Facility ID # " + facsForOwner.Rows(i)("ID").ToString
                                    spacingCount += 4
                                Next
                            End If
                        End With

                        'Retreive Past Due Fees for Document
                        showNewOwnerSection = oOwner.OwnerL2CSnippet

                        ds = oOwner.RunSQLQuery("SELECT dbo.udfGetOwnerPastDueFees(" + oOwner.ID.ToString + ",0,NULL)")
                        If ds.Tables(0).Rows(0)(0) > 0 Then
                            showFeeSection = True
                        End If

                        If Not colParams.Item("<TOSI Facility IDs>") Is Nothing Then
                            If colParams.Item("<TOSI Facility IDs>") <> String.Empty Then
                                showTOISSection = True
                            End If
                        End If

                        If showTransferOwnerSection Then
                            .Tables.Item(2).Rows.Item(1).Delete()
                        Else
                            .Tables.Item(2).Rows.Item(2).Delete()
                        End If

                        If Not showNewOwnerSection Then
                            .Tables.Item(2).Rows.Item(.Tables.Item(2).Rows.Count - 2).Delete()
                        End If

                        If Not showFeeSection Then
                            .Tables.Item(2).Rows.Item(.Tables.Item(2).Rows.Count - 1).Delete()
                        End If

                        If Not showTOISSection Then
                            .Tables.Item(2).Rows.Item(.Tables.Item(2).Rows.Count).Delete()
                        End If
                    End With
                End With

                'Save template
                LetterTemplateSave(WordApp, TemplatePath, DestinationPath)

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End Function

        Public Function CreateComplianceLetter(ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal facHasAllTOSITanks As Boolean = False)

            Dim strKey As String = String.Empty
            Dim strValue As String = String.Empty
            Dim i As Integer

            Try

                'Instantiate the Word Object
                LetterTemplateInit(WordApp, TemplatePath, DestinationPath)

                With WordApp
                    With DestDoc

                        ' Find and Replace the TAGs with Values.
                        For i = 0 To colParams.Count - 1
                            strKey = colParams.Keys(i).ToString

                            If Not colParams.Get(strKey) Is Nothing Then
                                strValue = colParams.Get(strKey).ToString
                                strValue = strValue.Replace(vbCrLf, Chr(13))
                                Dim go As Boolean = True

                                While go
                                    go = False

                                    With WordApp.Selection.Find
                                        .Text = strKey
                                        .Replacement.Text = ""
                                        .Forward = True
                                        .Wrap = Word.WdFindWrap.wdFindContinue
                                        .Execute()
                                    End With

                                    If Not WordApp.Selection.Text Is Nothing AndAlso WordApp.Selection.Text.ToUpper.Trim = strKey.ToUpper.Trim Then
                                        go = True
                                        WordApp.Selection.Text = IIf(strValue = String.Empty, "", strValue)
                                    End If

                                End While


                            Else
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)
                            End If
                        Next

                        'Remove table if all tanks are TOSI
                        If facHasAllTOSITanks Then
                            .Tables.Item(1).Delete()
                        End If

                        .Content.Select()

                        Dim objSelection As Word.Selection = WordApp.Selection

                        objSelection.Find.Forward = True
                        objSelection.Find.Format = True
                        objSelection.Find.Text = "<DeleteMe>"

                        Do While True
                            objSelection.Find.Execute()
                            If objSelection.Find.Found Then
                                objSelection.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdExtend)
                                objSelection.Delete()
                            Else
                                Exit Do
                            End If
                        Loop

                    End With
                End With

                'save template to doc   
                LetterTemplateSave(WordApp, TemplatePath, DestinationPath)

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End Function
        Public Function CreateTOSILetter(ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal facHasAllTOSITanks As Boolean = False)

            Dim strKey As String = String.Empty
            Dim strValue As String = String.Empty
            Dim i As Integer

            Try

                'Instantiate the Word Object
                LetterTemplateInit(WordApp, TemplatePath, DestinationPath)

                With WordApp
                    With DestDoc

                        ' Find and Replace the TAGs with Values.
                        For i = 0 To colParams.Count - 1
                            strKey = colParams.Keys(i).ToString

                            If Not colParams.Get(strKey) Is Nothing Then
                                strValue = colParams.Get(strKey).ToString
                                strValue = strValue.Replace(vbCrLf, Chr(13))
                                Dim go As Boolean = True

                                While go
                                    go = False

                                    With WordApp.Selection.Find
                                        .Text = strKey
                                        .Replacement.Text = ""
                                        .Forward = True
                                        .Wrap = Word.WdFindWrap.wdFindContinue
                                        .Execute()
                                    End With

                                    If Not WordApp.Selection.Text Is Nothing AndAlso WordApp.Selection.Text.ToUpper.Trim = strKey.ToUpper.Trim Then
                                        go = True
                                        WordApp.Selection.Text = IIf(strValue = String.Empty, "", strValue)
                                    End If

                                End While


                            Else
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)
                            End If
                        Next

                        'Remove table if all tanks are TOSI
                        If facHasAllTOSITanks Then
                            '     .Tables.Item(1).Delete()
                        End If

                        .Content.Select()

                        Dim objSelection As Word.Selection = WordApp.Selection

                        objSelection.Find.Forward = True
                        objSelection.Find.Format = True
                        objSelection.Find.Text = "<DeleteMe>"

                        Do While True
                            objSelection.Find.Execute()
                            If objSelection.Find.Found Then
                                objSelection.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdExtend)
                                objSelection.Delete()
                            Else
                                Exit Do
                            End If
                        Loop

                    End With
                End With

                'save template to doc   
                LetterTemplateSave(WordApp, TemplatePath, DestinationPath)

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End Function
#End Region

#Region "Closure Letter Operations"

        ' Closure Letters
        Public Function CreateClosureDemo(ByVal strModuleID As String, ByVal strLetter_To_Print As String, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByVal dtSample As DataTable = Nothing, Optional ByRef WordApp As Word.Application = Nothing)
            Try

                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty
                Dim i As Integer = 0

                'Instantiate the Word Object
                LetterTemplateInit(WordApp, TemplatePath, DestinationPath)

                With WordApp
                    With DestDoc

                        ' Find and Replace the TAGs with Values.
                        For i = 0 To colParams.Count - 1
                            strKey = colParams.Keys(i).ToString

                            If Not colParams.Get(strKey) Is Nothing Then
                                strValue = colParams.Get(strKey).ToString
                                strValue = strValue.Replace(vbCrLf, Chr(13))
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                            Else
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)
                            End If
                        Next


                        ' Add Table 
                        WordApp.Selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=28)
                        .Tables.Add(WordApp.Selection.Range, dtSample.Rows.Count + 1, dtSample.Columns.Count, Word.WdTableFormat.wdTableFormatList8)

                        With WordApp.Selection.Tables.Item(1)
                            .Style = "Table List 4"
                            .ApplyStyleHeadingRows = True
                            .ApplyStyleLastRow = True
                            .ApplyStyleFirstColumn = True
                            .ApplyStyleLastColumn = True
                        End With

                        'Add Columns 
                        For j As Integer = 0 To dtSample.Columns.Count - 1
                            WordApp.Selection.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            WordApp.Selection.TypeText(dtSample.Columns(j).ColumnName)
                            WordApp.Selection.MoveRight()
                        Next

                        'Fill Values in Table
                        Dim drow As DataRow
                        For Each drow In dtSample.Rows
                            For j As Integer = 0 To dtSample.Columns.Count - 1
                                WordApp.Selection.TypeText(CStr(drow.Item(j)))
                                WordApp.Selection.MoveRight()
                            Next
                        Next

                    End With
                End With

                LetterTemplateSave(WordApp, TemplatePath, DestinationPath)

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End Function

        Public Function CreateClosureInfoNeeded(ByVal TemplatePath As String, ByVal DestinationPath As String, ByVal colParams As Specialized.NameValueCollection, ByVal colInfoNeeded As ArrayList, Optional ByRef WordApp As Word.Application = Nothing)
            Try

                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty
                Dim i As Integer = 0

                'Instantiate the Word Object
                LetterTemplateInit(WordApp, TemplatePath, DestinationPath)

                With WordApp
                    With DestDoc

                        ' Find and Replace the TAGs with Values.
                        For i = 0 To colParams.Count - 1

                            strKey = colParams.Keys(i).ToString
                            If Not colParams.Get(strKey) Is Nothing Then
                                strValue = colParams.Get(strKey).ToString
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                Dim go As Boolean = True

                                While go
                                    go = False

                                    With WordApp.Selection.Find
                                        .Text = strKey
                                        .Replacement.Text = ""
                                        .Forward = True
                                        .Wrap = Word.WdFindWrap.wdFindContinue
                                        .Execute()
                                    End With

                                    If Not WordApp.Selection.Text Is Nothing AndAlso WordApp.Selection.Text.ToUpper.Trim = strKey.ToUpper.Trim Then
                                        go = True
                                        WordApp.Selection.Text = IIf(strValue = String.Empty, "", strValue)
                                    End If

                                End While
                            Else
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)
                            End If
                        Next

                        ' Add Info Needed
                        .Tables.Item(1).Cell(1, 1).Range.Text = " - " + colInfoNeeded.Item(0).ToString

                        For i = 1 To colInfoNeeded.Count - 1
                            With .Tables.Item(1)
                                .Rows.Add()
                                .Cell(i + 1, 1).Range.Text = " - " + colInfoNeeded.Item(i).ToString
                            End With
                        Next
                    End With
                End With

                LetterTemplateSave(WordApp, TemplatePath, DestinationPath)

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try

        End Function
#End Region

#Region "Inspection Letter Operations"

        ' Inspection Announcement Letter
        Public Function CreateInspAnnouncementLetters(ByVal ownerID As Integer, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, ByVal facsForOwner As DataTable, ByVal oOwner As MUSTER.BusinessLogic.pOwner, ByVal progressBarValueIncrement As Single, ByVal moduleID As Integer, ByVal staffID As Integer, ByRef returnVal As String, ByVal UserID As String, Optional ByRef WordApp As Word.Application = Nothing)
            Dim bolDeleteFile As Boolean = False
            Dim strKey As String = String.Empty
            Dim strValue As String = String.Empty
            Dim strAddress As String = String.Empty
            Dim oInspection As New MUSTER.BusinessLogic.pInspection
            Dim oAddressInfo As MUSTER.Info.AddressInfo
            Dim strDate, strTime, strTimes() As String
            Dim i, j As Integer
            Dim oPara, oParaPgBrk As Word.Paragraph
            Dim wordSel As Word.Selection
            Dim bolVars As InspPrepGuidelineVars
            Dim facLastInspDate As Date
            Dim dv As DataView
            Try

                ' check if user has rights to save inspection
                ' will be saving letter generated = true
                If Not oOwner.CheckWriteAccess(moduleID, staffID, SqlHelper.EntityTypes.Inspection) Then
                    returnVal = "You do not have rights to save Inspection."
                    Exit Function
                End If

                'Instantiate the Word Object
                LetterTemplateInit(WordApp, TemplatePath, DestinationPath)

                With WordApp
                    With DestDoc

                        ' Find and Replace the TAGs with Values.
                        For i = 0 To colParams.Count - 1
                            strKey = colParams.Keys(i).ToString

                            If Not colParams.Get(strKey) Is Nothing Then
                                strValue = colParams.Get(strKey).ToString
                                strValue = strValue.Replace(vbCrLf, Chr(13))
                                Dim go As Boolean = True

                                While go
                                    go = False

                                    With WordApp.Selection.Find
                                        .Text = strKey
                                        .Replacement.Text = ""
                                        .Forward = True
                                        .Wrap = Word.WdFindWrap.wdFindContinue
                                        .Execute()
                                    End With

                                    If Not WordApp.Selection.Text Is Nothing AndAlso WordApp.Selection.Text.ToUpper.Trim = strKey.ToUpper.Trim Then
                                        go = True
                                        WordApp.Selection.Text = IIf(strValue = String.Empty, "", strValue)
                                    End If

                                End While
                            Else
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)
                            End If
                        Next

                        ' Facility Table
                        With .Tables.Item(1)
                            ' Add rows
                            For i = 0 To facsForOwner.Rows.Count - 1
                                .Rows.Add()
                            Next
                        End With

                        ' Date Time Table
                        With .Tables.Item(2)
                            ' Add rows
                            For i = 0 To facsForOwner.Rows.Count - 1
                                .Rows.Add()
                            Next
                        End With

                        'fill rows
                        dv = facsForOwner.DefaultView
                        dv.Sort = "SCHEDULE_DATE, SCHEDULE_TIME"

                        For i = 0 To facsForOwner.Rows.Count - 1
                            oOwner.Facility = oOwner.OwnerInfo.facilityCollection.Item(dv.Item(i)("FACILITY_ID"))
                            oAddressInfo = oOwner.Facilities.FacilityAddress
                            strAddress = String.Empty
                            strAddress = oAddressInfo.AddressLine1.Trim + ", "
                            If oAddressInfo.AddressLine2.Trim <> String.Empty Then
                                strAddress += oAddressInfo.AddressLine2.Trim + ", "
                            End If
                            strAddress += oAddressInfo.City.Trim + " " + _
                                            oAddressInfo.State.Trim
                            strDate = CType(dv.Item(i)("SCHEDULE_DATE"), Date).ToString("MMMM d, yyyy")
                            strTime = dv.Item(i)("SCHEDULE_TIME").ToString.Trim
                            strTimes = strTime.Split(":")
                            If CType(strTimes(0), Integer) > 12 Then
                                strTimes(0) = (CType(strTimes(0), Integer) - 12).ToString
                                strTime = strTimes(0) + ":" + strTimes(1) + " p.m."
                            ElseIf CType(strTimes(0), Integer) = 12 Then
                                strTime += " p.m."
                            Else
                                strTime += " a.m."
                            End If
                            With .Tables.Item(1)
                                .Cell(i + 2, 1).Range.Text = dv.Item(i)("FACILITY_NAME").ToString
                                .Cell(i + 2, 2).Range.Text = strAddress
                                .Cell(i + 2, 3).Range.Text = dv.Item(i)("FACILITY_ID").ToString
                            End With
                            With .Tables.Item(2)
                                .Cell(i + 2, 1).Range.Text = dv.Item(i)("FACILITY_ID").ToString
                                .Cell(i + 2, 2).Range.Text = strDate
                                .Cell(i + 2, 3).Range.Text = strTime
                            End With
                        Next

                        ' PrepGuidelines
                        For i = 0 To facsForOwner.Rows.Count - 1
                            WordApp.Selection.GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:=.Bookmarks.Item("\endofdoc").Name)
                            WordApp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)

                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                            oPara.Range.Font.Name = "Arial"
                            oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            oPara.Range.Text = "UST Facility Compliance Inspection Requirements"
                            oPara.Range.Font.Size = 16
                            oPara.Range.Font.Bold = 1
                            oPara.Range.InsertParagraphAfter()

                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                            oPara.Range.Text = ""
                            oPara.Range.Font.Bold = 0
                            oPara.Range.InsertParagraphAfter()

                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                            oPara.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
                            oPara.Range.Text = "<Facility Name>, <Facility Address>, Facility ID # <Facility ID>"
                            oPara.Range.Font.Size = 10
                            oPara.Range.Font.Bold = 1
                            oPara.Range.InsertParagraphAfter()

                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                            oPara.Range.Text = ""
                            oPara.Range.Font.Bold = 0
                            oPara.Range.InsertParagraphAfter()

                            oInspection.Retrieve(dv.Item(i)("INSPECTION_ID"))
                            If dv.Item(i)("LAST_INSPECTED_ON") Is DBNull.Value Then
                                facLastInspDate = CDate("01/01/0001")
                            Else
                                facLastInspDate = dv.Item(i)("LAST_INSPECTED_ON")
                            End If
                            'facLastInspDate = IIf(Date.Compare(oInspection.RescheduledDate, CDate("01/01/0001")) = 0, oInspection.ScheduledDate, oInspection.RescheduledDate)
                            oOwner.Facility = oOwner.OwnerInfo.facilityCollection.Item(dv.Item(i)("FACILITY_ID"))
                            colParams.Clear()
                            colParams.Add("<Facility Name>", oOwner.Facilities.Name.Trim)
                            strAddress = oOwner.Facilities.FacilityAddress.AddressLine1.Trim
                            If oOwner.Facilities.FacilityAddress.AddressLine2 <> String.Empty Then
                                strAddress += ", " + oOwner.Facilities.FacilityAddress.AddressLine2.Trim
                            End If
                            strAddress += ", " + oOwner.Facilities.FacilityAddress.City.Trim
                            strAddress += ", " + oOwner.Facilities.FacilityAddress.State.Trim
                            colParams.Add("<Facility Address>", strAddress)
                            colParams.Add("<Facility ID>", oOwner.Facilities.ID.ToString)

                            ' Find and Replace the TAGs with Values.
                            For j = 0 To colParams.Count - 1
                                strKey = colParams.Keys(j).ToString
                                strValue = colParams.Get(strKey).ToString
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                            Next

                            'InitVars(bolVars)
                            'ProcessVars(bolVars, oOwner.Facility, facLastInspDate)

                            ' records
                            ProcessRecords(bolVars, WordApp, DestDoc, ownerID, dv.Item(i)("FACILITY_ID"))


                            'space
                            If bolVars.RecordsHeader Then
                                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                oPara.Range.Text = ""
                                oPara.Range.Font.Bold = 0
                                oPara.Format.SpaceAfter = 1
                                oPara.Range.InsertParagraphAfter()
                            End If

                            ' components
                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                            oPara.Range.Text = "In addition to the record review, you must also provide access to the following components of the UST during the inspection:"
                            oPara.Range.Font.Bold = 0
                            oPara.Format.SpaceAfter = 1
                            oPara.Range.InsertParagraphAfter()

                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                            oPara.Range.Text = ""
                            oPara.Range.Font.Bold = 0
                            oPara.Format.SpaceAfter = 1
                            oPara.Range.InsertParagraphAfter()

                            ProcessComponents(bolVars, WordApp, DestDoc, ownerID, dv.Item(i)("FACILITY_ID"))

                            oInspection.LetterGenerated = True
                            oInspection.DateLetterGenerated = Now.Date
                            oInspection.ModifiedBy = UserID

                            oInspection.Save(moduleID, staffID, returnVal)
                            If Not returnVal = String.Empty Then
                                bolDeleteFile = True
                                Exit Function
                            End If

                            For Each oPara In .Paragraphs
                                If oPara.ID = "BULLET" Then
                                    oPara.Range.Style = Word.WdBuiltinStyle.wdStyleListBullet
                                End If
                            Next
                            RaiseEvent CheckListProgress(progressBarValueIncrement)
                        Next

                    End With
                End With

                LetterTemplateSave(WordApp, TemplatePath, DestinationPath)

            Catch ex As Exception
                bolDeleteFile = True
                Throw ex

            Finally
                If bolDeleteFile Then
                    RaiseEvent CloseCheckListProgress()
                End If
            End Try
        End Function

        ' CheckList
        Public Function CreateInspCheckList(ByVal facID As Integer, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, ByVal oInsp As MUSTER.BusinessLogic.pInspection, ByVal dtTank As DataSet, ByVal dtPipe As DataSet, ByVal dsTerm As DataSet, ByVal progress As Integer, ByVal sketchPath As String, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal imgs As Collections.ArrayList = Nothing, Optional ByVal msg As String = "", Optional ByVal comments As String = "", Optional ByVal thirdPartyOperator As String = "") As Word.Application

            Try
                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty
                Dim i, j, rowIndex, colIndex, colCount As Integer

                Dim StartTanks As Boolean = True
                Dim StartPipes As Boolean = True
                Dim StarTterms As Boolean = True

                Dim oPara As Word.Paragraph
                Dim oTable As Word.Table

                'Instantiate the Word Object
                LetterTemplateInit(WordApp, TemplatePath, DestinationPath)

                With WordApp
                    .Visible = False

                    With DestDoc



                        ' Find and Replace the TAGs with Values.
                        For i = 0 To colParams.Count - 1
                            strKey = colParams.Keys(i).ToString
                            If Not colParams.Get(strKey) Is Nothing Then
                                strValue = colParams.Get(strKey).ToString
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                            Else
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)
                            End If
                        Next


                        progress += 10
                        RaiseEvent CheckListProgress(progress)


                        ' Add Tanks
                        For Each dt As DataTable In dtTank.Tables
                            'Init Paragraph Spaces 
                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                            oPara.Range.Text = "<Space>"
                            oPara.Format.SpaceAfter = 1
                            oPara.Range.InsertParagraphAfter()

                            InspectionAddTanks(dt, WordApp, DestDoc, StartTanks)
                        Next
                        progress += 10
                        RaiseEvent CheckListProgress(progress)


                        ' Add Pipes


                        'Add Pipes
                        For Each dt As DataTable In dtPipe.Tables
                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                            oPara.Range.Text = "<Space>"
                            oPara.Format.SpaceAfter = 1
                            oPara.Range.InsertParagraphAfter()

                            InspectionAddPipes(dt, WordApp, DestDoc, StartPipes)
                        Next

                        progress += 10
                        RaiseEvent CheckListProgress(progress)


                        ' Add Terminations
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "<Space>"
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()

                        InspectionAddTerms(dsTerm, WordApp, DestDoc, StarTterms)
                        progress += 10
                        RaiseEvent CheckListProgress(progress)

                        ''Init Checklist report
                        DevelopChecklistOnDocument(WordApp, DestDoc, oInsp, imgs, sketchPath, msg, comments, thirdPartyOperator)

                        .Content.Find.Execute(FindText:="<Space>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)

                        LetterTemplateSave(WordApp, TemplatePath, DestinationPath, True)

                        Return WordApp

                    End With
                End With


                RaiseEvent CheckListProgress(100)

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End Function
#End Region


        ' C & E

#Region "C & E Letter Operations"

        Public Function GenerateCAELetter(ByVal colParams As Specialized.NameValueCollection, ByVal strFacsForAgreedOrder() As String, _
                ByVal TemplatePath As String, ByVal COCTemplatePath As String, ByVal xCOITemplatePath As String, ByVal DestinationPath As String, _
                ByVal dtFacs As DataTable, ByVal dtCitations As DataTable, ByVal dtDiscreps As DataTable, ByVal dtCorActions As DataTable, ByVal dtCOIFacs As DataTable, _
                ByVal alFacsIndex As ArrayList, ByVal alCitationsIndex As ArrayList, _
                ByVal alCorrectiveActionIndex As ArrayList, ByVal alCorrectiveActionWithDueDateIndex As ArrayList, _
                ByVal alCorrectiveActionAddOnIndex As ArrayList, ByVal alDiscrepsIndex As ArrayList, ByVal alDiscrepsCorrectiveActionIndex As ArrayList, _
                Optional ByVal bolCOCRequired As Boolean = False, Optional ByRef WordApp As Word.Application = Nothing)
            Dim strKey As String = String.Empty
            Dim strValue As String = String.Empty
            Dim prevFac, i, j, k, nStrKeyEndIndex, lineNumBegin, lineNumEnd As Integer
            Dim dt As DataTable
            Dim dv As DataView
            Dim temp As WordTemplate
            Dim multiFacs As Boolean = False
            Dim totalCnt As Integer
            Dim resetPath As Boolean = True

            Try

                'Instantiate the Word Object
                LetterTemplateInit(WordApp, TemplatePath, DestinationPath)


                Dim wrdGlobal As New Word.[Global]
                temp = New WordTemplate(wrdGlobal.ListGalleries.Item(Word.WdListGalleryType.wdOutlineNumberGallery).ListTemplates.Item(1), TemplatePath.Substring(0, TemplatePath.LastIndexOf("\")))
                wrdGlobal = Nothing
                ' MsgBox("here3")
                With WordApp
                    With DestDoc

                        ' Find and Replace the TAGs with Values.
                        For i = 0 To colParams.Count - 1
                            strKey = colParams.Keys(i).ToString
                            If Not colParams.Get(strKey) Is Nothing Then
                                strValue = colParams.Get(strKey).ToString
                                strValue = strValue.Replace(vbCrLf, Chr(13))
                                ' MsgBox("here31")
                                Dim go As Boolean = True

                                While go

                                    go = False
                                    With WordApp.Selection.Find
                                        .Text = strKey
                                        .Replacement.Text = ""
                                        .Forward = True
                                        .Wrap = Word.WdFindWrap.wdFindContinue
                                        .Execute()
                                    End With
                                    '   WordApp.Selection.Find.Text = strKey
                                    If Not WordApp.Selection.Text Is Nothing AndAlso WordApp.Selection.Text.ToUpper.Trim = strKey.ToUpper.Trim Then
                                        go = True 'tag found
                                        WordApp.Selection.Text = IIf(strValue = String.Empty, "", strValue)
                                        'With WordApp.Selection.Find
                                        '    .Text = strKey
                                        '    .Replacement.Text = ""
                                        '    .Forward = True
                                        '    .Wrap = Word.WdFindWrap.wdFindContinue
                                        '    .Execute()
                                        'End With
                                        ' go = False
                                    Else ' tag not found
                                        '  go = True
                                    End If
                                End While
                            Else
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)
                                ' MsgBox("here34")
                            End If
                            'MsgBox("here35")
                        Next
                        ' MsgBox("here4")
                        ' A,B,C
                        If Not dtFacs Is Nothing Then
                            ' Facility Table
                            If alFacsIndex.Count > 0 Then
                                For Each facIndex As Integer In alFacsIndex.ToArray
                                    With .Tables.Item(facIndex)
                                        ' Add rows
                                        For i = 0 To dtFacs.Rows.Count - 1
                                            .Rows.Add()
                                        Next
                                    End With

                                    dv = dtFacs.DefaultView
                                    dv.Sort = "FACILITY_ID"


                                    For i = 0 To dtFacs.Rows.Count - 1

                                        If i >= 1 Then
                                            multiFacs = True
                                        End If

                                        With .Tables.Item(facIndex)
                                            .Cell(i + 2, 1).Range.Text = "Facility ID #" + dv.Item(i)("FACILITY_ID").ToString + ", " + _
                                                                        dv.Item(i)("FACILITY").ToString + ", " + _
                                                                        dv.Item(i)("ADDRESS").ToString

                                            .Cell(i + 2, 1).Range.Font.Bold = 0


                                        End With
                                    Next ' For i = 0 To dtFacs.Rows.Count - 1
                                Next ' For Each facIndex As Integer In dtFacsIndex.ToArray
                            End If ' If dtFacsIndex.Count > 0 Then
                        End If ' If Not dtFacs Is Nothing Then
                        '  MsgBox("here5")
                        .Content.Find.Execute(FindText:="<Facility>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)
                        '              .Content.Find.Execute(FindText:="<Facility Name>", ReplaceWith:=dv.Item(i)("FACILITY").ToString, Replace:=Word.WdReplace.wdReplaceAll)
                        '             .Content.Find.Execute(FindText:="<Facility Address>", ReplaceWith:=dv.Item(i)("ADDRESS").ToString, Replace:=Word.WdReplace.wdReplaceAll)
                        '            .Content.Find.Execute(FindText:="<Facility ID>", ReplaceWith:=dv.Item(i)("FACILITY_ID").ToString, Replace:=Word.WdReplace.wdReplaceAll)

                        ' facs for agreed order
                        If Not strFacsForAgreedOrder Is Nothing Then
                            If strFacsForAgreedOrder.Length > 0 Then
                                WordApp.Selection.Find.ClearFormatting()
                                With WordApp.Selection.Find
                                    .Text = "<Facilities>"
                                    .Replacement.Text = ""
                                    .Forward = True
                                    .Wrap = Word.WdFindWrap.wdFindContinue
                                    .Format = False
                                    .MatchCase = False
                                    .MatchWholeWord = False
                                    .MatchWildcards = False
                                    .MatchSoundsLike = False
                                    .MatchAllWordForms = False
                                End With
                                WordApp.Selection.Find.Execute()
                                For i = 0 To strFacsForAgreedOrder.Length - 1
                                    If i < strFacsForAgreedOrder.Length - 1 Then
                                        WordApp.Selection.TypeText(strFacsForAgreedOrder(i) + "; ")
                                    Else
                                        WordApp.Selection.TypeText(strFacsForAgreedOrder(i))
                                    End If
                                Next
                                'End If
                            End If
                        End If

                        SetUpList(WordApp, temp)

                        If Not dtCitations Is Nothing Then
                            '  Dim citIndex As Integer
                            ' Citation Table
                            If alCitationsIndex.Count > 0 Then

                                For Each citIndex As Integer In alCitationsIndex.ToArray

                                    '    totalCnt = alCitationsIndex.Count
                                    '   For citIndex = 2 To totalCnt + 1
                                    prevFac = 0
                                    ' Fill Table with Text
                                    ' With .Tables.Item(citindex)
                                    With .Tables.Item(citIndex)


                                        If WordApp.Selection.Range.InStory(.Cell(1, 1).Range) Then

                                            WordApp.Selection.Document.Content.Tables.Item(citIndex).Cell(1, 1).Range.Select()

                                            WordApp.Selection.Range.InsertParagraph()

                                            ApplyToList(WordApp, temp, multiFacs)

                                        End If
                                        dv = dtCitations.DefaultView
                                        dv.Sort = "FACILITy_ID, INDX"
                                        For i = 0 To dv.Count - 1
                                            Threading.Thread.Sleep(100)

                                            If prevFac <> dv.Item(i)("FACILITY_ID") Then

                                                prevFac = dv.Item(i)("FACILITY_ID")
                                                If dtFacs.Rows.Count > 1 Then

                                                    If i > 0 Then
                                                        WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlinePromote()
                                                    End If

                                                    WordApp.Selection.Range.Bold = 1

                                                    WordApp.Selection.TypeText("Facility ID #" + dv.Item(i)("FACILITY_ID").ToString + Chr(13))

                                                    WordApp.Selection.Range.Bold = 0

                                                    WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlineDemote()


                                                End If
                                            End If

                                            WordApp.Selection.TypeText(dv.Item(i)("CITATIONTEXT").ToString.Replace(Chr(13), String.Empty).Replace(Chr(7), String.Empty).Trim)

                                            Dim CCAT As New System.Text.StringBuilder
                                            Dim CCATStr As String = String.Empty

                                            If Not TypeOf dv.Item(i).Item("CCAT_COMMENTS") Is DBNull Then
                                                CCATStr = dv.Item(i).Item("CCAT_COMMENTS")
                                            Else
                                                CCATStr = dv.Item(i).Item("CCAT")
                                            End If

                                            If CCATStr.Length > 0 Then


                                                For Each item As String In CCATStr.Trim.Split(","c)
                                                    item = item.Trim

                                                    If item.Length > 2 AndAlso item.StartsWith("PT") Then
                                                        CCAT.Append("Term #").Append(item.Substring(2)).Append(Chr(13))
                                                    ElseIf item.Length > 1 AndAlso item.StartsWith("PMW") Then
                                                        CCAT.Append("Piping Trench Monitoring Well #").Append(item.Substring(3)).Append(Chr(13))
                                                    ElseIf item.Length > 1 AndAlso item.StartsWith("TMW") Then
                                                        CCAT.Append("Tank Monitoring Well #").Append(item.Substring(3)).Append(Chr(13))
                                                    ElseIf item.Length > 1 AndAlso item.StartsWith("P") Then
                                                        CCAT.Append("Pipe #").Append(item.Substring(1)).Append(Chr(13))
                                                    ElseIf item.Length > 1 AndAlso item.StartsWith("T") Then
                                                        CCAT.Append("Tank #").Append(item.Substring(1)).Append(Chr(13))
                                                    ElseIf item.Length > 0 Then
                                                        CCAT.Append("Entity #").Append(item.ToString).Append(Chr(13))
                                                    End If

                                                Next

                                                WordApp.Selection.TypeText(Chr(13))
                                                WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlineDemote()

                                                Dim ii As Integer = 0

                                                For Each str As String In CCAT.ToString.Split(Chr(13))

                                                    If str.Length > 0 Then

                                                        WordApp.Selection.TypeText(str.Replace(Chr(13), String.Empty).Replace(Chr(7), String.Empty).Trim)

                                                        If ii < CCAT.ToString.Split(Chr(13)).GetUpperBound(0) Then
                                                            WordApp.Selection.TypeText(Chr(13))
                                                        End If

                                                    End If

                                                    ii += 1

                                                Next

                                                WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlinePromote()

                                                CCAT.Length = 0
                                            Else
                                                WordApp.Selection.TypeText(Chr(13))

                                            End If

                                            If i = dtCitations.Rows.Count - 1 Then
                                                WordApp.Selection.TypeParagraph()
                                                WordApp.Selection.TypeBackspace()
                                                WordApp.Selection.TypeBackspace()
                                            End If

                                        Next
                                    End With
                                Next ' For Each citIndex As Integer In dtCitationsIndex.ToArray

                            End If ' If dtCitationsIndex.Count > 0 Then
                        End If





                        If Not dtDiscreps Is Nothing Then

                            ' discrep Table
                            If alDiscrepsIndex.Count > 0 Then
                                For Each discrepIndex As Integer In alDiscrepsIndex.ToArray
                                    prevFac = 0
                                    ' Fill Table with Text
                                    With .Tables.Item(discrepIndex)


                                        If WordApp.Selection.Range.InStory(.Cell(1, 1).Range) Then

                                            WordApp.Selection.Document.Content.Tables.Item(discrepIndex).Cell(1, 1).Range.Select()

                                            WordApp.Selection.Range.InsertParagraph()

                                            ApplyToList(WordApp, temp, multiFacs)


                                        End If
                                        dv = dtDiscreps.DefaultView
                                        dv.Sort = "FACILITy_ID, INDX"
                                        For i = 0 To dv.Count - 1
                                            Threading.Thread.Sleep(100)

                                            If prevFac <> dv.Item(i)("FACILITY_ID") Then

                                                prevFac = dv.Item(i)("FACILITY_ID")
                                                If dtFacs.Rows.Count > 1 Then

                                                    If i > 0 Then
                                                        WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlinePromote()
                                                    End If

                                                    WordApp.Selection.TypeText("Facility ID #" + dv.Item(i)("FACILITY_ID").ToString + Chr(13))

                                                    WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlineDemote()


                                                End If
                                            End If

                                            WordApp.Selection.TypeText(dv.Item(i)("DISCREP TEXT").ToString.Replace(Chr(13), String.Empty).Replace(Chr(7), String.Empty).Trim)

                                            Dim CCAT As New System.Text.StringBuilder
                                            Dim CCATStr As String = String.Empty

                                            If Not TypeOf dv.Item(i).Item("CCAT_COMMENTS") Is DBNull Then
                                                CCATStr = dv.Item(i).Item("CCAT_COMMENTS")
                                            Else
                                                CCATStr = dv.Item(i).Item("CCAT")
                                            End If

                                            If CCATStr.Length > 0 AndAlso (dv.Item(i)("QUESTION_ID") = 36 OrElse dv.Item(i)("QUESTION_ID") = 44 OrElse dv.Item(i)("QUESTION_ID") = 31 OrElse dv.Item(i)("QUESTION_ID") < 0) Then

                                                For Each item As String In CCATStr.Trim.Split(","c)
                                                    item = item.Trim

                                                    If item.Length > 2 AndAlso item.StartsWith("PT") Then
                                                        CCAT.Append("Term #").Append(item.Substring(2)).Append(Chr(13))
                                                    ElseIf item.Length > 1 AndAlso item.StartsWith("PMW") Then
                                                        CCAT.Append("Piping Trench Monitoring Well #").Append(item.Substring(3)).Append(Chr(13))
                                                    ElseIf item.Length > 1 AndAlso item.StartsWith("TMW") Then
                                                        CCAT.Append("Tank Monitoring Well #").Append(item.Substring(3)).Append(Chr(13))

                                                    ElseIf item.Length > 1 AndAlso item.StartsWith("P") Then
                                                        CCAT.Append("Pipe #").Append(item.Substring(1)).Append(Chr(13))
                                                    ElseIf item.Length > 1 AndAlso item.StartsWith("T") Then
                                                        CCAT.Append("Tank #").Append(item.Substring(1)).Append(Chr(13))
                                                    ElseIf item.Length > 0 Then
                                                        CCAT.Append("Entity #").Append(item.ToString).Append(Chr(13))
                                                    End If

                                                Next

                                                WordApp.Selection.TypeText(Chr(13))
                                                WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlineDemote()

                                                Dim ii As Integer = 0

                                                For Each str As String In CCAT.ToString.Split(Chr(13))

                                                    If str.Length > 0 Then

                                                        WordApp.Selection.TypeText(str.Replace(Chr(13), String.Empty).Replace(Chr(7), String.Empty).Trim)

                                                        If ii < CCAT.ToString.Split(Chr(13)).GetUpperBound(0) Then
                                                            WordApp.Selection.TypeText(Chr(13))
                                                        End If

                                                    End If

                                                    ii += 1

                                                Next

                                                WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlinePromote()

                                                CCAT.Length = 0
                                            Else
                                                WordApp.Selection.TypeText(Chr(13))
                                            End If

                                            If i = dtDiscreps.Rows.Count - 1 Then
                                                WordApp.Selection.TypeParagraph()
                                                WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlinePromote()
                                                WordApp.Selection.TypeBackspace()
                                                WordApp.Selection.TypeBackspace()
                                            End If

                                        Next
                                    End With
                                Next ' For Each discrepIndex As Integer In dtCitationsIndex.ToArray

                            End If ' If dtdiscrepIndex.Count > 0 Then
                        End If




                        If Not dtCitations Is Nothing OrElse Not dtCorActions Is Nothing Then

                            Dim subContext As Boolean = False

                            ' Corrective Actions Table (NEW)
                            If alCorrectiveActionIndex.Count > 0 Then


                                Dim pull As String = String.Empty

                                For Each citIndex As Integer In alCorrectiveActionIndex.ToArray
                                    prevFac = 0
                                    ' Fill Table with Text
                                    With .Tables.Item(citIndex)


                                        If WordApp.Selection.Range.InStory(.Cell(1, 1).Range) Then

                                            WordApp.Selection.Document.Content.Tables.Item(citIndex).Cell(1, 1).Range.Select()

                                            pull = WordApp.Selection.Text.Replace(Chr(13), String.Empty).Replace(Chr(7), String.Empty).Trim

                                            WordApp.Selection.Range.InsertParagraph()

                                            ApplyToList(WordApp, temp, multiFacs)

                                        End If

                                        If Not dtCorActions Is Nothing AndAlso dtCorActions.Rows.Count > 0 Then
                                            dt = dtCorActions
                                            dv = dtCorActions.DefaultView
                                            dv.Sort = "FACILITy_ID, INDX"

                                        Else
                                            dt = dtCitations
                                            dv = dtCitations.DefaultView
                                            dv.Sort = "FACILITY_ID, INDX"

                                        End If

                                        Dim InDetailed As Boolean = False

                                        For i = 0 To dv.Count
                                            Threading.Thread.Sleep(100)

                                            If i = dv.Count OrElse prevFac <> dv.Item(i)("FACILITY_ID") Then

                                                If i < dv.Count Then prevFac = dv.Item(i)("FACILITY_ID")

                                                If dtFacs.Rows.Count > 1 OrElse pull.Length > 0 OrElse i = dv.Count Then



                                                    If i > 0 AndAlso dtFacs.Rows.Count > 1 Then
                                                        WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlinePromote()
                                                    End If

                                                    WordApp.Selection.Range.Bold = 1

                                                    If i < dv.Count AndAlso dtFacs.Rows.Count > 1 Then

                                                        WordApp.Selection.TypeText("Facility ID #" + dv.Item(i)("FACILITY_ID").ToString + Chr(13))
                                                    ElseIf i = dv.Count Then

                                                        If pull.Length > 0 Then

                                                            WordApp.Selection.TypeText(pull)
                                                            WordApp.Selection.TypeText(Chr(13))

                                                        End If

                                                        WordApp.Selection.TypeBackspace()

                                                        Exit For
                                                    End If

                                                    WordApp.Selection.Range.Bold = 0

                                                    If dtFacs.Rows.Count > 1 Then

                                                        WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlineDemote()
                                                    End If


                                                ElseIf i = dv.Count Then

                                                    If pull.Length > 0 Then
                                                        WordApp.Selection.TypeText(pull)
                                                        WordApp.Selection.TypeText(Chr(13))
                                                    End If
                                                    WordApp.Selection.TypeBackspace()

                                                    Exit For



                                                End If
                                            End If

                                            Dim HasMoreThanOne As Boolean = False


                                            Dim drr As DataRow() = dt.Select(String.Format("[General Corrective Action] = '{0}' and Facility_ID = {1}", dv.Item(i)("General Corrective Action"), dv.Item(i)("FACILITY_ID")))
                                            Dim n = 0

                                            If Not drr Is Nothing AndAlso drr.GetUpperBound(0) > 0 Then
                                                For Each row As DataRow In drr
                                                    If n >= 1 AndAlso row.Item("CorrectiveAction") <> drr(n - 1).Item("CorrectiveAction") Then
                                                        HasMoreThanOne = True
                                                        Exit For
                                                    End If
                                                    n += 1
                                                Next
                                            End If



                                            If Not HasMoreThanOne Then
                                                WordApp.Selection.TypeText(dv.Item(i)("CorrectiveAction").Replace(Chr(13), String.Empty).Replace(Chr(7), String.Empty).Trim)
                                            End If




                                            If HasMoreThanOne Then



                                                If Not InDetailed Then
                                                    WordApp.Selection.TypeText(dv.Item(i)("General Corrective Action").Replace(Chr(13), String.Empty).Replace(Chr(7), String.Empty).Trim)
                                                End If


                                                If i = 0 OrElse dv.Item(i - 1)("General Corrective Action") <> dv.Item(i)("General Corrective Action") Then

                                                    WordApp.Selection.TypeText(Chr(13))
                                                    WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlineDemote()
                                                    InDetailed = True

                                                End If


                                                WordApp.Selection.TypeText(dv.Item(i)("CorrectiveAction").Replace(Chr(13), String.Empty).Replace(Chr(7), String.Empty).Trim)

                                                If i < dv.Count - 1 AndAlso dv.Item(i + 1)("General Corrective Action") = dv.Item(i)("General Corrective Action") Then
                                                    WordApp.Selection.TypeText(Chr(13))

                                                End If

                                                If (i < dv.Count - 1 AndAlso dv.Item(i + 1)("General Corrective Action") <> dv.Item(i)("General Corrective Action")) OrElse i = dv.Count - 1 Then

                                                    WordApp.Selection.TypeText(Chr(13))
                                                    WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlinePromote()
                                                    InDetailed = False
                                                End If
                                            ElseIf i < dv.Count Then
                                                WordApp.Selection.TypeText(Chr(13))
                                            End If

                                        Next
                                    End With
                                Next ' For Each citCorrIndex As Integer In dtCitationsIndex.ToArray
                            End If ' If dvIndex.Count > 0 Then
                        End If




                        If Not dtCitations Is Nothing OrElse Not dtCorActions Is Nothing Then

                            Dim pull As String = String.Empty

                            ' Corrective Actions DUE DATE Table (NEW)
                            If alCorrectiveActionWithDueDateIndex.Count > 0 Then


                                For Each citIndex As Integer In alCorrectiveActionWithDueDateIndex.ToArray
                                    prevFac = 0
                                    ' Fill Table with Text
                                    With .Tables.Item(citIndex)


                                        If WordApp.Selection.Range.InStory(.Cell(1, 1).Range) Then

                                            WordApp.Selection.Document.Content.Tables.Item(citIndex).Cell(1, 1).Range.Select()

                                            pull = WordApp.Selection.Text.Replace(Chr(13), String.Empty).Replace(Chr(7), String.Empty).Trim

                                            WordApp.Selection.Range.InsertParagraph()

                                            ApplyToList(WordApp, temp, multiFacs)

                                        End If

                                        If Not dtCorActions Is Nothing AndAlso dtCorActions.Rows.Count > 0 Then
                                            dt = dtCorActions
                                            dv = dtCorActions.DefaultView
                                            dv.Sort = "FACILITy_ID, INDX"

                                        Else
                                            dt = dtCitations
                                            dv = dtCitations.DefaultView
                                            dv.Sort = "FACILITY_ID, INDX"

                                        End If

                                        Dim InDetailed As Boolean = False

                                        DestDoc.Content.Find.Execute(FindText:="In lieu of a formal enforcement hearing concerning the violations listed above, Complainant and Respondent agree to settle this matter as follows:", ReplaceWith:=String.Format("In lieu of a formal enforcement hearing concerning the violations listed above, Complainant and Respondent agree to settle this matter by {0:MMMM dd yyyy} as follows:", Convert.ToDateTime(dv.Item(0)("DUE"))), Replace:=Word.WdReplace.wdReplaceAll)

                                        For i = 0 To dv.Count
                                            Threading.Thread.Sleep(100)




                                            If i = dv.Count OrElse prevFac <> dv.Item(i)("FACILITY_ID") Then

                                                If i < dv.Count Then prevFac = dv.Item(i)("FACILITY_ID")

                                                If dtFacs.Rows.Count > 1 OrElse pull.Length > 0 Then

                                                    If i > 0 AndAlso dtFacs.Rows.Count > 1 Then
                                                        WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlinePromote()
                                                    End If

                                                    WordApp.Selection.Range.Bold = 1

                                                    If i < dv.Count And dtFacs.Rows.Count > 1 Then

                                                        WordApp.Selection.TypeText("Facility ID #" + dv.Item(i)("FACILITY_ID").ToString + Chr(13))

                                                    ElseIf i = dv.Count Then

                                                        If pull.Length > 0 Then
                                                            WordApp.Selection.TypeText(pull)
                                                            WordApp.Selection.TypeText(Chr(13))
                                                        End If
                                                        WordApp.Selection.TypeBackspace()

                                                        Exit For
                                                    End If

                                                    WordApp.Selection.Range.Bold = 0

                                                    If dtFacs.Rows.Count > 1 Then

                                                        WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlineDemote()
                                                    End If
                                                ElseIf i = dv.Count Then

                                                    If pull.Length > 0 Then
                                                        WordApp.Selection.TypeText(pull)
                                                        WordApp.Selection.TypeText(Chr(13))
                                                    End If
                                                    WordApp.Selection.TypeBackspace()

                                                    Exit For




                                                End If
                                            End If

                                            Dim HasMoreThanOne As Boolean = False
                                            Dim drr As DataRow() = dt.Select(String.Format("[General Corrective Action] = '{0}' and Facility_ID = {1}", dv.Item(i)("General Corrective Action"), dv.Item(i)("FACILITY_ID")))
                                            Dim n = 0

                                            If Not drr Is Nothing AndAlso drr.GetUpperBound(0) > 0 Then
                                                For Each row As DataRow In drr
                                                    If n >= 1 AndAlso row.Item("CorrectiveAction") <> drr(n - 1).Item("CorrectiveAction") Then
                                                        HasMoreThanOne = True
                                                        Exit For
                                                    End If
                                                    n += 1
                                                Next
                                            End If


                                            Dim textStr As String



                                            If Not HasMoreThanOne Then

                                                textStr = dv.Item(i)("CorrectiveAction").ToString.Trim

                                                If textStr.IndexOf(".") <= -1 Then
                                                    textStr = String.Format("{0}.", textStr)
                                                End If

                                                WordApp.Selection.TypeText(textStr.Replace(Chr(13), String.Empty).Replace(Chr(7), String.Empty).Trim)
                                            End If


                                            If HasMoreThanOne Then

                                                If Not InDetailed Then

                                                    textStr = dv.Item(i)("General Corrective Action").ToString.Trim

                                                    If textStr.IndexOf(".") <= -1 Then
                                                        textStr = String.Format("{0}.", textStr)
                                                    End If

                                                    WordApp.Selection.TypeText(textStr.Replace(Chr(13), String.Empty).Replace(Chr(7), String.Empty).Trim)
                                                End If


                                                If i = 0 OrElse dv.Item(i - 1)("General Corrective Action") <> dv.Item(i)("General Corrective Action") Then

                                                    WordApp.Selection.TypeText(Chr(13))
                                                    WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlineDemote()
                                                    InDetailed = True

                                                End If


                                                WordApp.Selection.TypeText(dv.Item(i)("CorrectiveAction").Replace(Chr(13), String.Empty).Replace(Chr(7), String.Empty).Trim)

                                                If i < dv.Count - 1 AndAlso dv.Item(i + 1)("General Corrective Action") = dv.Item(i)("General Corrective Action") Then
                                                    WordApp.Selection.TypeText(Chr(13))

                                                End If

                                                If (i < dv.Count - 1 AndAlso dv.Item(i + 1)("General Corrective Action") <> dv.Item(i)("General Corrective Action")) OrElse i = dv.Count - 1 Then

                                                    WordApp.Selection.TypeText(Chr(13))
                                                    WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlinePromote()
                                                    InDetailed = False
                                                End If

                                            ElseIf i < dv.Count Then
                                                WordApp.Selection.TypeText(Chr(13))
                                            End If

                                        Next

                                    End With
                                Next ' For Each citCorrIndex As Integer In dtCitationsIndex.ToArray

                            End If ' If dvIndex.Count > 0 Then
                        End If


                        resetPath = True
                        .Content.Find.Execute(FindText:="<Citation>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)
                        .Content.Find.Execute(FindText:="<CorrectiveAction>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)
                        .Content.Find.Execute(FindText:="<CorrectiveActionWithDueDate>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)
                        .Content.Find.Execute(FindText:="<CorrectiveActionWithDueDate", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)



                        If Not dtDiscreps Is Nothing Then

                            ' Discrep corrective action table
                            If alDiscrepsCorrectiveActionIndex.Count > 0 Then
                                For Each discrepCAIndex As Integer In alDiscrepsCorrectiveActionIndex.ToArray
                                    prevFac = 0
                                    ' Fill Table with Text
                                    With .Tables.Item(discrepCAIndex)


                                        If WordApp.Selection.Range.InStory(.Cell(1, 1).Range) Then

                                            WordApp.Selection.Document.Content.Tables.Item(discrepCAIndex).Cell(1, 1).Range.Select()

                                            WordApp.Selection.Range.InsertParagraph()

                                            ApplyToList(WordApp, temp, multiFacs)


                                        End If
                                        dv = dtDiscreps.DefaultView
                                        dv.Sort = "FACILITy_ID, INDX"
                                        For i = 0 To dtDiscreps.Rows.Count - 1
                                            Threading.Thread.Sleep(100)

                                            If prevFac <> dv.Item(i)("FACILITY_ID") Then

                                                prevFac = dv.Item(i)("FACILITY_ID")
                                                If dtFacs.Rows.Count > 1 Then

                                                    If i > 0 Then
                                                        WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlinePromote()
                                                    End If

                                                    WordApp.Selection.TypeText("Facility ID #" + dv.Item(i)("FACILITY_ID").ToString + Chr(13))

                                                    WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlineDemote()


                                                End If
                                            End If

                                            WordApp.Selection.TypeText(dv.Item(i)("CorrectiveAction").ToString.Trim)

                                            WordApp.Selection.TypeText(Chr(13))

                                            If i = dtDiscreps.Rows.Count - 1 Then
                                                WordApp.Selection.TypeParagraph()
                                                WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlinePromote()
                                                WordApp.Selection.TypeBackspace()
                                                WordApp.Selection.TypeBackspace()
                                            End If

                                        Next
                                    End With
                                Next ' For Each discrepIndex As Integer In dtCitationsIndex.ToArray

                            End If ' If dtdiscrepIndex.Count > 0 Then
                        End If


                        .Content.Find.Execute(FindText:="<Discrepancy>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)
                        .Content.Find.Execute(FindText:="<DiscrepCorrectiveAction>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)

                        .Content.Select()

                        Dim objSelection As Word.Selection = WordApp.Selection

                        objSelection.Find.Forward = True
                        objSelection.Find.Format = True
                        objSelection.Find.Text = "<DeleteMe>"

                        Do While True
                            objSelection.Find.Execute()
                            If objSelection.Find.Found Then
                                objSelection.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdExtend)
                                objSelection.Delete()
                            Else
                                Exit Do
                            End If
                        Loop



                        ' remove the extra tables
                        Dim deltblCount As Integer = 0

                        If alCorrectiveActionAddOnIndex.Count > 0 Then
                            For Each i In alCorrectiveActionAddOnIndex.ToArray

                                If .Tables.Count >= (i - deltblCount) Then

                                    .Tables.Item(i - deltblCount).Delete()
                                    deltblCount += 1
                                End If

                            Next
                        End If



                        'While keepsaving

                        'Try

                        'keepsaving = False

                        '.Save()

                        ' Catch ex As Exception
                        '    If ex.ToString.ToUpper.IndexOf(" PERMISSION") > -1 Then

                        '   Threading.Thread.Sleep(1000)

                        '  keepsaving = True

                        ' End If

                        ' End Try

                        ' End While
                        ' append coc
                        If bolCOCRequired Then
                            If COCTemplatePath <> String.Empty Then
                                If System.IO.File.Exists(COCTemplatePath) Then
                                    If Not dtFacs Is Nothing Then
                                        If dtFacs.Rows.Count > 0 Then
                                            dv = dtFacs.DefaultView
                                            dv.Sort = "FACILITY_ID"
                                            For i = 0 To dtFacs.Rows.Count - 1
                                                Threading.Thread.Sleep(100)

                                                ' insert page break
                                                DestDoc.Application.Selection.EndKey(Unit:=Word.WdUnits.wdStory)
                                                DestDoc.Application.Selection.InsertBreak(Word.WdBreakType.wdPageBreak)

                                                ' insert file
                                                DestDoc.Application.Selection.InsertFile(FileName:=COCTemplatePath, ConfirmConversions:=False, Link:=False, Attachment:=False)

                                                ' set tags
                                                ' <COC Facility Name>
                                                strKey = "<COC Facility Name>"
                                                strValue = dv.Item(i)("FACILITY")
                                                'strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                                ' <COC Facility Address 1>
                                                strKey = "<COC Facility Address 1>"
                                                strValue = dv.Item(i)("ADDRESS_LINE_ONE")
                                                'strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                                ' <COC Facility Address 2>
                                                'strKey = "<COC Facility Address 2>"
                                                'strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                                '.Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                                ' <COC Facility City/State/Zip>
                                                strKey = "<COC Facility City/State/Zip>"
                                                strValue = dv.Item(i)("CITY") + ", " + _
                                                            dv.Item(i)("STATE") + " " + _
                                                            dv.Item(i)("ZIP")
                                                'strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                                ' <COC Facility ID>
                                                strKey = "<COC Facility ID>"
                                                strValue = dv.Item(i)("FACILITY_ID").ToString
                                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                                ' <COC Owner Name>
                                                'strKey = "<COC Owner Name>"
                                                strKey = colParams.Keys(5).ToString
                                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                                                ' <COC Owner Address1>
                                                strKey = colParams.Keys(7).ToString
                                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                                strKey = colParams.Keys(8).ToString
                                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                                                ' <COC Owner City/State/Zip>
                                                strKey = colParams.Keys(9).ToString
                                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = "<DeleteMe>", "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                                ' <Date>
                                                strKey = "<Date>"
                                                strValue = IIf(IsNothing(colParams.Get(strKey)), "", colParams.Get(strKey).ToString)
                                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                                ' <Due Date>
                                                strKey = "<Due Date>"
                                                strValue = IIf(IsNothing(colParams.Get(strKey)), "", colParams.Get(strKey).ToString)
                                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                            Next ' For i = 0 To dtFacs.Rows.Count - 1
                                        End If ' If dtFacs.Rows.Count > 0 Then
                                    End If ' If Not dtFacs Is Nothing Then
                                End If ' If System.IO.File.Exists(COCTemplatePath) Then
                            End If ' If COCTemplatePath <> String.Empty Then
                        End If

                        '''' append coi
                        'If COITemplatePath <> String.Empty Then
                        'If System.IO.File.Exists(COITemplatePath) Then
                        '  If Not dtCOIFacs Is Nothing Then
                        '   If dtCOIFacs.Rows.Count > 0 Then
                        '       dv = dtCOIFacs.DefaultView
                        '       dv.Sort = "FACILITY_ID"
                        '       Dim dr As DataRow
                        '           For i = 0 To dtCOIFacs.Rows.Count - 1
                        '            Threading.Thread.Sleep(100)

                        '           dr = dtFacs.Select("FACILITY_ID = " + dv.Item(i)("FACILITY_ID").ToString)(0)
                        ''          ' insert page break
                        '        DestDoc.Application.Selection.EndKey(Unit:=Word.WdUnits.wdStory)
                        '       DestDoc.Application.Selection.InsertBreak(Word.WdBreakType.wdPageBreak)

                        ' insert file
                        '      DestDoc.Application.Selection.InsertFile(FILENAME:=COITemplatePath, ConfirmConversions:=False, Link:=False, Attachment:=False)

                        ' set tags
                        ' <COI Facility Name>
                        '     strKey = "<COI Facility Name>"
                        'strValue = dv.Item(i)("FACILITY")
                        '    strValue = dr("FACILITY")
                        'strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                        '   .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                        ' <COI Facility Address 1>
                        '  strKey = "<COI Facility Address 1>"
                        ' strValue = dr("ADDRESS_LINE_ONE")
                        'strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                        '.Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                        '' <COI Facility Address 2>
                        'strKey = "<COI Facility Address 2>"
                        'strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                        '.Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                        '' <COI Facility City/State/Zip>
                        'strKey = "<COI Facility City/State/Zip>"
                        'strValue = dr("CITY") + ", " + _
                        '            dr("STATE") + " " + _
                        '            dr("ZIP")
                        'strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                        ' .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                        ''' <COI Facility ID>
                        ' strKey = "<COI Facility ID>"
                        ' strValue = dr("FACILITY_ID").ToString
                        ' .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                        ' ' <Due Date>
                        ' strKey = "<Due Date>"
                        ' strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                        ' .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                        ' Next ' For i = 0 To dtCOIFacs.Rows.Count - 1
                        ' End If ' If dtCOIFacs.Rows.Count > 0 Then
                        ' End If ' If Not dtCOIFacs Is Nothing Then
                        'end If ' If System.IO.File.Exists(COITemplatePath) Then
                        'End If ' If COITemplatePath <> String.Empty Then


                        'Special cleaning for Agreed orders with one facility only
                        If dtFacs Is Nothing OrElse dtFacs.Rows.Count <= 1 Then
                            .Content.Find.Execute(FindText:="'(USTs) at establishments", ReplaceWith:="(USTs) at an establishment", Replace:=Word.WdReplace.wdReplaceAll)
                        End If
                    End With
                End With
                'Dim objSelection2 As Word.Selection = WordApp.Selection

                'objSelection2.Find.Forward = False
                'objSelection2.Find.Format = True
                'objSelection2.Find.Text = "<DeleteMe>"

                'Do While True
                '    objSelection2.Find.Execute()
                '    If objSelection2.Find.Found Then
                '        objSelection2.EndKey(Word.WdUnits.wdLine, Word.WdMovementType.wdExtend)
                '        objSelection2.Delete()
                '    Else
                '        Exit Do
                '    End If
                'Loop
                LetterTemplateSave(WordApp, TemplatePath, DestinationPath, True)

                Return WordApp

            Catch ex As Exception
                Throw New Exception(ex.ToString)
            End Try
        End Function
#End Region


#Region "Financial Letter Operation"
        Public Function CreateFinancialLetter(ByVal strModuleID As String, ByVal strLetter_To_Print As String, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal strfile As String = "", Optional ByVal strSignature As String = "") ', Optional ByVal strFiles As String = "")
            Try

                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty

                Dim i As Integer = 0
                Dim bolKeyDeductionReasons As Boolean = False
                Dim bolKeyReimbursementConditions As Boolean = False
                Dim strDeductionReasonsValue As String = String.Empty
                Dim strReimbursementConditionsValue As String = String.Empty

                'Instantiate the Word Object
                LetterTemplateInit(WordApp, TemplatePath, DestinationPath)

                With WordApp

                    With DestDoc

                        If Not colParams.Item("<ERAC Contact>, <ERAC>") Is Nothing Then
                            .Content.Find.Execute(FindText:="<ERAC Contact>, <ERAC>", ReplaceWith:="<ERAC>", Replace:=Word.WdReplace.wdReplaceAll)
                        End If

                        ' Find and Replace the TAGs with Values.
                        For i = 0 To colParams.Count - 1
                            strKey = colParams.Keys(i).ToString
                            If Not colParams.Get(strKey) Is Nothing Then
                                strValue = colParams.Item(strKey).ToString
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                If strKey <> "<Reimbursement Conditions>" And strKey <> "<DEDUCTION REASONS>" Then
                                    Dim go As Boolean = True

                                    While go
                                        go = False

                                        With WordApp.Selection.Find
                                            .Text = strKey
                                            .Replacement.Text = ""
                                            .Forward = True
                                            .Wrap = Word.WdFindWrap.wdFindContinue
                                            .Execute()
                                        End With

                                        If Not WordApp.Selection.Text Is Nothing AndAlso WordApp.Selection.Text.ToUpper.Trim = strKey.ToUpper.Trim Then
                                            go = True
                                            WordApp.Selection.Text = IIf(strValue = String.Empty, "", strValue)
                                        End If

                                    End While
                                ElseIf strKey = "<DEDUCTION REASONS>" Then
                                    bolKeyDeductionReasons = True
                                    strDeductionReasonsValue = strValue
                                ElseIf strKey = "<Reimbursement Conditions>" Then
                                    strReimbursementConditionsValue = strValue
                                    bolKeyReimbursementConditions = True
                                ElseIf strKey = "<ERAC Contact>, <ERAC>" Then
                                    'ignore
                                End If
                            Else
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)
                            End If

                        Next

                        'Modifications for Financial Letters
                        If TemplatePath.EndsWith("MGPTFApprovalFormTemplateNotRSOM.doc") Then
                            If bolKeyReimbursementConditions Then
                                .Content.Find.Execute(FindText:="<Reimbursement Conditions>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)
                                If .Tables.Count > 0 Then
                                    .Tables.Item(1).Cell(3, 1).Tables.Item(1).Cell(1, 1).Range.Text = strReimbursementConditionsValue
                                End If
                            End If
                        End If

                        If TemplatePath.EndsWith("MGPTFApprovalFormTemplateRSOMYr1.doc") Or _
                            TemplatePath.EndsWith("MGPTFApprovalFormTemplateRSOMYr2-7.doc") Then
                            If bolKeyReimbursementConditions Then
                                .Content.Find.Execute(FindText:="<Reimbursement Conditions>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)
                                If .Tables.Count > 0 Then
                                    .Tables.Item(2).Cell(1, 1).Range.Text = strReimbursementConditionsValue
                                End If
                            End If
                        End If

                        If bolKeyReimbursementConditions And Not (TemplatePath.EndsWith("MGPTFApprovalFormTemplateNotRSOM.doc") Or _
                                                                    TemplatePath.EndsWith("MGPTFApprovalFormTemplateRSOMYr1.doc") Or _
                                                                    TemplatePath.EndsWith("MGPTFApprovalFormTemplateRSOMYr2-7.doc")) Then
                            .Content.Find.Execute(FindText:="<Reimbursement Conditions>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)
                            If .Tables.Count > 0 Then
                                .Tables.Item(1).Cell(1, 1).Range.Text = strReimbursementConditionsValue
                            End If
                        End If

                        If bolKeyDeductionReasons And Not (TemplatePath.EndsWith("MGPTFApprovalFormTemplateNotRSOM.doc") Or _
                                                                    TemplatePath.EndsWith("MGPTFApprovalFormTemplateRSOMYr1.doc") Or _
                                                                    TemplatePath.EndsWith("MGPTFApprovalFormTemplateRSOMYr2-7.doc")) Then
                            .Content.Find.Execute(FindText:="<DEDUCTION REASONS>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)
                            If .Tables.Count > 0 Then
                                .Tables.Item(1).Cell(1, 1).Range.Text = strDeductionReasonsValue
                            End If
                        End If

                    End With
                End With

                LetterTemplateSave(WordApp, TemplatePath, DestinationPath)

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End Function

        Public Function CreateFinancialGenericLetter(ByVal strModuleID As String, ByVal strLetter_To_Print As String, ByVal colParams As Specialized.NameValueCollection, ByVal cols As Int16, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal TableFormat As Int16 = 35, Optional ByVal bApplyBorders As Boolean = False, Optional ByVal bApplyShading As Boolean = False, Optional ByVal bApplyFont As Boolean = True, Optional ByVal bApplyColor As Boolean = False, Optional ByVal bApplyHeader As Boolean = True, Optional ByVal bApplyLastRow As Boolean = False, Optional ByVal bApplyFirstCol As Boolean = False, Optional ByVal bApplyLastCol As Boolean = False, Optional ByVal bApplyAutofit As Boolean = True)
            Try
                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty
                Dim i As Integer = 0
                Dim oPara As Word.Paragraph

                'Instantiate the Word Object
                LetterTemplateInit(WordApp, TemplatePath, DestinationPath)

                With WordApp


                    With DestDoc

                        ' Find and Replace the TAGs with Values.
                        If Not colParams.Item("<ERAC Contact>, <ERAC>") Is Nothing Then
                            .Content.Find.Execute(FindText:="<ERAC Contact>, <ERAC>", ReplaceWith:="<ERAC>", Replace:=Word.WdReplace.wdReplaceAll)
                        End If

                        For i = 0 To colParams.Count - 1

                            strKey = colParams.Keys(i).ToString
                            If Not colParams.Get(strKey) Is Nothing Then
                                strValue = colParams.Item(strKey)
                                strValue = strValue.Replace(vbCrLf, Chr(13))


                                If strKey = "<ERAC Contact>, <ERAC>" Then
                                    'ignore

                                ElseIf strKey = "<DATA>" Then
                                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                    oPara.Range.Text = strValue
                                    'Word.WdTableFormat.wdTableFormatContemporary
                                    oPara.Range.ConvertToTable("|", , cols, , TableFormat, bApplyBorders, bApplyShading, bApplyFont, bApplyColor, bApplyHeader, bApplyLastRow, bApplyFirstCol, bApplyLastCol, bApplyAutofit)
                                    oPara.Format.SpaceAfter = 1
                                    oPara.Range.InsertParagraphAfter()

                                Else
                                    Dim go As Boolean = True

                                    While go
                                        go = False

                                        With WordApp.Selection.Find
                                            .Text = strKey
                                            .Replacement.Text = ""
                                            .Forward = True
                                            .Wrap = Word.WdFindWrap.wdFindContinue
                                            .Execute()
                                        End With

                                        If Not WordApp.Selection.Text Is Nothing AndAlso WordApp.Selection.Text.ToUpper.Trim = strKey.ToUpper.Trim Then
                                            go = True
                                            WordApp.Selection.Text = IIf(strValue = String.Empty, "", strValue)
                                        End If

                                    End While
                                End If
                            Else
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:="", Replace:=Word.WdReplace.wdReplaceAll)

                            End If

                        Next
                    End With

                    'Initiate Page setup for generic Letter
                    LetterTemplateGenericPageSetup(WordApp, DestDoc)

                End With

                LetterTemplateSave(WordApp, TemplatePath, DestinationPath)

            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try

        End Function

        Public Function CreateCompanyInfoLetter(ByVal strModuleID As String, ByVal strLetter_To_Print As String, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal strfile As String = "", Optional ByVal strSignature As String = "") ', Optional ByVal strFiles As String = "")

            Try
                Return Me.CreateLetter(strModuleID, strLetter_To_Print, colParams, TemplatePath, DestinationPath, WordApp, strfile, strSignature)

            Catch ex As Exception
                If Not ex.InnerException Is Nothing Then
                    Throw New Exception(ex.InnerException.ToString)
                Else
                    Throw New Exception(ex.ToString)
                End If
            End Try

        End Function
#End Region

#End Region

#Region "Old Commented Code on Exposed Collections Functions"

        'Private Sub CreateEnvelope(ByRef WordApp As Word.Application, ByRef doc As Word.Document, ByVal TemplatePath As String, ByVal strName As String, ByVal strAddress1 As String, ByVal strCity As String, ByVal strZipCode As String, Optional ByVal strAddress2 As String = "")
        '    Dim wrdMergeFields As Word.MailMergeFields
        '    Dim wrdSelection As Word.Selection
        '    Dim file As file
        '    Dim docAddress As Word.Document
        '    wrdSelection = WordApp.Selection
        '    'WordApp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)
        '    doc.MailMerge.MainDocumentType = Word.WdMailMergeMainDocType.wdEnvelopes
        '    If WordApp.ActiveWindow.View.SplitSpecial = Word.WdSpecialPane.wdPaneNone Then
        '        WordApp.ActiveWindow.ActivePane.View.Type = Word.WdViewType.wdPrintView
        '    Else
        '        WordApp.ActiveWindow.View.Type = Word.WdViewType.wdPrintView
        '    End If
        '    WordApp.ActiveWindow.ActivePane.View.Zoom.PageFit = Word.WdPageFit.wdPageFitBestFit
        '    'if file exists, delete to start with fresh file
        '    If (file.Exists("C:\addresses.doc")) Then
        '        docAddress = WordApp.Documents.Open("C:\addresses.doc")
        '        docAddress.Close(False)
        '        System.IO.File.Delete("C:\addresses.doc")
        '        docAddress = Nothing
        '    End If
        '    doc.MailMerge.CreateDataSource(Name:="C:\addresses.doc", _
        '        HeaderRecord:="Name, Address1, Address2, City,State")
        '    docAddress = WordApp.Documents.Open("C:\addresses.doc")
        '    docAddress.Tables.Item(1).Rows.Add()
        '    With docAddress.Tables.Item(1)
        '        ' Insert the data in the specific cell.
        '        .Cell(2, 1).Range.InsertAfter("abcd")
        '        .Cell(2, 2).Range.InsertAfter("Address1")
        '        .Cell(2, 3).Range.InsertAfter("Address2")
        '        .Cell(2, 4).Range.InsertAfter("City")
        '        .Cell(2, 5).Range.InsertAfter("State")
        '    End With
        '    docAddress.Save()
        '    docAddress.Close(False)
        '    docAddress = Nothing
        '    wrdMergeFields = doc.MailMerge.Fields()
        '    wrdMergeFields.Add(wrdSelection.Range, "Name")
        '    wrdSelection.TypeParagraph()
        '    wrdMergeFields.Add(wrdSelection.Range, "Address1")
        '    wrdSelection.TypeParagraph()
        '    wrdMergeFields.Add(wrdSelection.Range, "Address2")
        '    wrdSelection.TypeParagraph()
        '    wrdMergeFields.Add(wrdSelection.Range, "City")
        '    wrdSelection.TypeParagraph()
        '    wrdMergeFields.Add(wrdSelection.Range, "State")
        '    'WordApp.Selection.TypeBackspace()
        '    'WordApp.Selection.TypeBackspace()
        '    'WordApp.Selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=3)
        '    'doc.Fields.Add(Range:=WordApp.Selection.Range, Type:= _
        '    'Word.WdFieldType.wdFieldAddressBlock, Text:="\f ""<<NAME >><<ADDRESS1>><<ADDRESS2>><<CITY>>" & Chr(13) & "<<state" & Chr(13) & ">>")
        '    'doc.wdFieldAddressBlock, 
        '    WordApp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)
        '    WordApp.Selection.InsertFile(FILENAME:=TemplatePath.Trim, Range:="", _
        '                ConfirmConversions:=False, Link:=False, Attachment:=False)
        '    With doc.MailMerge
        '        .Destination = Word.WdMailMergeDestination.wdSendToNewDocument
        '        .SuppressBlankLines = True
        '        .Execute(Pause:=False)
        '    End With

        'End Sub
#End Region

#Region "Private Inspection Annoucement Letter Functions"
        Private Sub InitVars(ByRef bolVars As InspPrepGuidelineVars)
            Try
                bolVars.P_LDhasElecALLD = False
                bolVars.T_LDhasGroundWaterOrVaporMonitoring = False
                bolVars.P_LDhasGroundWaterOrVaporMonitoring = False
                bolVars.T_LDhasAutomaticTankGauging = False
                bolVars.T_LDhasStatisticalInventory = False
                bolVars.P_LDhasStatisticalInventory = False
                bolVars.T_LDhasInventoryControl = False
                bolVars.T_hasInstDtgt12_21_88_neverInspected = False
                bolVars.T_hasInstDtgt12_21_88_neverInspected_MatOfConstructionhasFigerGlassReinforcedPlastic = False
                bolVars.P_LDhasLineTightnessTesting = False
                bolVars.P_hasInstDtgt12_21_88_neverInspected = False
                bolVars.T_LDhasManualTankGauging = False
                bolVars.T_LDhasContinuous_Electronic_Visual_Manual_Monitoring = False
                bolVars.P_LDhasContinuous_Electronic_Visual_Manual_Monitoring = False
                bolVars.T_SteelTank = False
                bolVars.T_ImpressedCurrentCP = False
                bolVars.P_ImpressedCurrentCP = False
                bolVars.P_Steel_MetallicPiping = False
                bolVars.P_LDhasMechanicalAutomaticLine = False
                bolVars.RecordsHeader = False
                ' Components
                bolVars.T_EPGonly = False
                bolVars.T_UsedOilonly = False
                bolVars.P_pressurizedPipeSystem = False
            Catch ex As Exception
                Throw ex
            End Try
        End Sub
        Private Sub ProcessVars(ByRef bolVars As InspPrepGuidelineVars, ByVal facInfo As MUSTER.Info.FacilityInfo, ByVal facLastInspDate As Date)
            Dim tnk As MUSTER.Info.TankInfo
            Dim pipe As MUSTER.Info.PipeInfo
            Dim comp As MUSTER.Info.CompartmentInfo
            Dim nullDate As Date = CDate("01/01/0001")
            Dim pre88 As Date = CDate("12/21/1988")
            Dim tnkCount, epgTnkCount, compTnkCount, usedOilTnkCount As Integer
            Try
                tnkCount = 0
                epgTnkCount = 0
                compTnkCount = 0
                usedOilTnkCount = 0
                For Each tnk In facInfo.TankCollection.Values
                    If tnk.TankStatus = 424 Or _
                        tnk.TankStatus = 425 Or _
                        tnk.TankStatus = 429 Then ' CIU OR TOS OR TOSI
                        tnkCount += 1
                        If Date.Compare(tnk.DateInstalledTank, pre88) > 0 Then
                            If Date.Compare(tnk.DateInstalledTank, facLastInspDate) > 0 Then
                                bolVars.T_hasInstDtgt12_21_88_neverInspected = True
                                If tnk.TankMatDesc = 348 Then
                                    bolVars.T_hasInstDtgt12_21_88_neverInspected_MatOfConstructionhasFigerGlassReinforcedPlastic = True
                                End If
                            End If
                        End If
                        bolVars.T_ImpressedCurrentCP = IIf(tnk.TankCPType = 418, True, bolVars.T_ImpressedCurrentCP)
                        bolVars.T_LDhasAutomaticTankGauging = IIf(tnk.TankLD = 336, True, bolVars.T_LDhasAutomaticTankGauging)
                        bolVars.T_LDhasContinuous_Electronic_Visual_Manual_Monitoring = IIf(tnk.TankLD = 339 Or tnk.TankLD = 343, True, bolVars.T_LDhasContinuous_Electronic_Visual_Manual_Monitoring)
                        bolVars.T_LDhasGroundWaterOrVaporMonitoring = IIf(tnk.TankLD = 335, True, bolVars.T_LDhasGroundWaterOrVaporMonitoring)
                        bolVars.T_LDhasInventoryControl = IIf(tnk.TankLD = 338, True, bolVars.T_LDhasInventoryControl)
                        bolVars.T_LDhasManualTankGauging = IIf(tnk.TankLD = 337, True, bolVars.T_LDhasManualTankGauging)
                        bolVars.T_LDhasStatisticalInventory = IIf(tnk.TankLD = 340, True, bolVars.T_LDhasStatisticalInventory)
                        bolVars.T_SteelTank = IIf(tnk.TankMatDesc = 344 Or _
                                                    tnk.TankMatDesc = 347, True, bolVars.T_SteelTank)
                        epgTnkCount = IIf(tnk.TankEmergen, epgTnkCount + 1, epgTnkCount)
                        For Each comp In tnk.CompartmentCollection.Values
                            compTnkCount += 1
                            If comp.Substance = 314 Then
                                usedOilTnkCount += 1
                            End If
                        Next
                    End If

                    For Each pipe In tnk.pipesCollection.Values
                        If pipe.PipeStatusDesc = 424 Or _
                            pipe.PipeStatusDesc = 425 Or _
                            pipe.PipeStatusDesc = 429 Then ' CIU OR TOS OR TOSI
                            If Date.Compare(pipe.PipeInstallDate, pre88) > 0 Then
                                If Date.Compare(pipe.PipeInstallDate, facLastInspDate) > 0 Then
                                    bolVars.P_hasInstDtgt12_21_88_neverInspected = True
                                End If
                            End If
                            bolVars.P_ImpressedCurrentCP = IIf(pipe.PipeCPType = 478, True, bolVars.P_ImpressedCurrentCP)
                            bolVars.P_LDhasLineTightnessTesting = IIf(pipe.PipeLD = 245, True, bolVars.P_LDhasLineTightnessTesting)
                            bolVars.P_LDhasContinuous_Electronic_Visual_Manual_Monitoring = IIf(pipe.PipeLD = 242 Or pipe.PipeLD = 243, True, bolVars.P_LDhasContinuous_Electronic_Visual_Manual_Monitoring)
                            bolVars.P_LDhasElecALLD = IIf(pipe.PipeLD = 246, True, bolVars.P_LDhasElecALLD)
                            bolVars.P_LDhasGroundWaterOrVaporMonitoring = IIf(pipe.PipeLD = 241, True, bolVars.P_LDhasGroundWaterOrVaporMonitoring)
                            bolVars.P_LDhasMechanicalAutomaticLine = IIf(pipe.ALLDType = 496, True, bolVars.P_LDhasMechanicalAutomaticLine)
                            bolVars.P_LDhasStatisticalInventory = IIf(pipe.PipeLD = 244, True, bolVars.P_LDhasStatisticalInventory)
                            bolVars.P_Steel_MetallicPiping = IIf(pipe.PipeMatDesc = 250 Or _
                                                                    pipe.PipeMatDesc = 251 Or _
                                                                    pipe.PipeMatDesc = 253, True, bolVars.P_Steel_MetallicPiping)
                            bolVars.P_pressurizedPipeSystem = IIf(pipe.PipeTypeDesc = 266, True, bolVars.P_pressurizedPipeSystem)
                        End If
                    Next
                Next
                bolVars.T_EPGonly = IIf(epgTnkCount = tnkCount, True, bolVars.T_EPGonly)
                bolVars.T_UsedOilonly = IIf(usedOilTnkCount = compTnkCount, True, bolVars.T_UsedOilonly)
            Catch ex As Exception
                Throw ex
            End Try
        End Sub
        Private Sub ProcessRecordsHeader(ByRef bolVars As InspPrepGuidelineVars, ByRef WordApp As Word.Application, ByRef DestDoc As Word.Document)
            Dim oPara As Word.Paragraph
            Try
                With DestDoc
                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "The following records must be made available for our review at the time of the inspection:"
                    oPara.Range.Font.Name = "Arial"
                    oPara.Range.Font.Bold = 0
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = ""
                    oPara.Range.Font.Bold = 0
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()

                    bolVars.RecordsHeader = True
                End With
            Catch ex As Exception
            End Try
        End Sub
        Private Sub ProcessRecords(ByRef bolVars As InspPrepGuidelineVars, ByRef WordApp As Word.Application, ByRef DestDoc As Word.Document, Optional ByVal ownerId As Integer = 0, Optional ByVal facID As Integer = 0)
            Dim oPara As Word.Paragraph
            Dim dt As DataTable
            Try
                With DestDoc


                    Dim db As New DataAccess.InspectionChecklistMasterDB
                    dt = db.DBGetAnnouncementLetterProcesses(ownerId, facID)

                    If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then


                        For Each dr As DataRow In dt.Rows
                            If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                            oPara.Range.Text = dr("statement")
                            oPara.ID = "BULLET"
                            oPara.Format.SpaceAfter = 1
                            oPara.Range.InsertParagraphAfter()
                        Next

                        dt.Dispose()

                    End If

                    dt = Nothing
                    db = Nothing

                    Exit Sub


                    ' 1   -- old code
                    If bolVars.P_LDhasElecALLD Then
                        If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "Electronic line leak detector test records (0.2 gph tests)"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                    ' 2
                    If bolVars.T_LDhasGroundWaterOrVaporMonitoring Or bolVars.P_LDhasGroundWaterOrVaporMonitoring Then
                        If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "Groundwater/Vapor Monitoring records (provide copies of the previous 12 months of monitoring well observations)"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                    ' 3
                    If bolVars.T_LDhasAutomaticTankGauging Then
                        If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "Automatic Tank Gauging records (provide copies of the previous 12 months of 'leak test' records)"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                    ' 4
                    If bolVars.T_LDhasStatisticalInventory Or bolVars.P_LDhasStatisticalInventory Then
                        If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.ID = "BULLET"
                        oPara.Range.Text = "Statistical Inventory Reconciliation (provide copies of the previous 12 months of 'Pass/Fail' reports)"
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                    ' 5
                    If bolVars.T_LDhasInventoryControl Then
                        If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "Inventory Control records (provide copies of the previous 12 months)"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                    ' 6
                    If bolVars.T_LDhasInventoryControl Or bolVars.T_hasInstDtgt12_21_88_neverInspected Then
                        If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "Precision Tightness Test record for the tanks (provide a copy of the most recent precision tightness test)"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                    ' 7
                    If bolVars.P_LDhasLineTightnessTesting Or bolVars.P_hasInstDtgt12_21_88_neverInspected Then
                        If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "Precision Tightness Test record for the piping (provide a copy of the most recent precision tightness test)"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                    ' 8
                    If bolVars.T_LDhasManualTankGauging Then
                        If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "Manual Tank Gauging records (provide copies of the previous 12 months of records)"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                    ' 9
                    If bolVars.T_LDhasContinuous_Electronic_Visual_Manual_Monitoring Or bolVars.P_LDhasContinuous_Electronic_Visual_Manual_Monitoring Then
                        If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "Interstitial Monitoring records (provide copies of the previous 12 months of records)"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                    ' 10
                    If bolVars.T_hasInstDtgt12_21_88_neverInspected_MatOfConstructionhasFigerGlassReinforcedPlastic Then
                        If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "Tank deflection test records"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                    ' 11
                    If bolVars.T_SteelTank Then
                        If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "Testing of the tank cathodic protection system (provide copy of the last test record)"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                    ' 12
                    If bolVars.T_ImpressedCurrentCP Then
                        If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "Records showing the impressed current cathodic protection system rectifier has been checked for operation at least once every 60 days"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                    ' 13
                    If bolVars.P_Steel_MetallicPiping Then
                        If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "Testing of the pipe cathodic protection system (provide copy of the last test record)"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                    ' 14
                    If bolVars.P_LDhasMechanicalAutomaticLine Then
                        If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "Testing of the automatic line leak detectors (provide copies of the last two test records)"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If


                    If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "Testing of the spill buckets"
                    oPara.ID = "BULLET"
                    'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()



                    'oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    'oPara.Range.Text = "<REMOVE>"
                End With
            Catch ex As Exception
                Throw ex
            End Try
        End Sub
        Private Sub ProcessComponents(ByVal bolVars As InspPrepGuidelineVars, ByRef WordApp As Word.Application, ByRef DestDoc As Word.Document, Optional ByVal ownerid As Integer = 0, Optional ByVal facid As Integer = 0)
            Dim oPara As Word.Paragraph
            Dim dt As DataTable = Nothing
            Try

                With DestDoc

                    Dim db As New DataAccess.InspectionChecklistMasterDB
                    dt = db.DBGetAnnouncementLetterComponents(ownerid, facid)

                    If Not dt Is Nothing AndAlso dt.Rows.Count > 0 Then


                        For Each dr As DataRow In dt.Rows
                            If Not bolVars.RecordsHeader Then ProcessRecordsHeader(bolVars, WordApp, DestDoc)
                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                            oPara.Range.Text = dr("statement")
                            oPara.ID = "BULLET"
                            oPara.Format.SpaceAfter = 1
                            oPara.Range.InsertParagraphAfter()
                        Next

                        dt.Dispose()

                    End If

                    dt = Nothing
                    db = Nothing

                    Exit Sub


                    ' 15
                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "The tank fill ports (ensure that keys are available for any locks that may be on the fill port caps)"
                    oPara.ID = "BULLET"
                    'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "Compliance Manager's Certificate"
                    oPara.ID = "BULLET"
                    'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "UST Operation Clerk Log"
                    oPara.ID = "BULLET"
                    'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()




                    If Not (bolVars.T_EPGonly Or bolVars.T_UsedOilonly) Then
                        ' 16
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "The fuel dispenser cabinets (ensure that keys are available to open the dispenser cabinet doors)"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If

                    If bolVars.P_pressurizedPipeSystem Then
                        ' 17
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "The submersible pump heads (ensure that any soil and/or water that may be covering the pump heads is removed for the inspection)"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If

                    If Not bolVars.T_UsedOilonly Then
                        ' 18
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "The tank spill prevention devices"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If

                    If Not bolVars.T_UsedOilonly Then
                        ' 19
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "The tank overfill prevention devices (ensure that our inspector is able to physically verify that overfill prevention devices are installed in the tanks"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If

                    If bolVars.T_LDhasGroundWaterOrVaporMonitoring Or bolVars.P_LDhasGroundWaterOrVaporMonitoring Then
                        ' 20
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "The monitoring wells (ensure that keys are available for any locks that may be on the monitoring well caps)"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If

                    If bolVars.T_LDhasAutomaticTankGauging Then
                        ' 21
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "The Automatic Tank Gauging system"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If

                    If bolVars.T_ImpressedCurrentCP Then
                        ' 22
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "The Impressed Current Cathodic Protection System Rectifier (ensure that keys are available for any locks that may be on the rectifier cabinet)"
                        oPara.ID = "BULLET"
                        'oPara.Range.ListFormat.ApplyListTemplate(ListTemplate:=WordApp.ListGalleries(Word.WdListGalleryType.wdBulletGallery).ListTemplates(1), ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                    'oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    'oPara.Range.Text = "<REMOVE>"
                End With
            Catch ex As Exception
                Throw ex
            End Try
        End Sub


#End Region

#Region "Private Inspection Checklist Functions"

        Function GetApproxInches(ByVal percent As Double, ByVal lineInches As Double) As Double
            Return lineInches * percent
        End Function

        Private Sub DevelopChecklistOnDocument(ByVal wordapp As Word.Application, ByVal destdoc As Word.Document, ByVal oInsp As BusinessLogic.pInspection, ByVal imgs As Collections.ArrayList, ByVal sketchpath As String, ByVal msg As String, ByVal comments As String, ByVal ThirdPartyOperator As String)

            Dim dr, drows() As DataRow
            Dim dt, dtSub As DataTable
            Dim dv, dvSub As DataView
            Dim ds As DataSet
            Dim oPara As Word.Paragraph
            Dim oTable As Word.Table
            Dim i As Integer
            Dim hasDesignatedOwner As Boolean = True


            With destdoc
                ''Init Checklist report
                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                oPara.Range.InsertBreak()

                ' registration / testing / construction
                dt = oInsp.CheckListMaster.RegTable()
                dv = dt.DefaultView
                dv.Sort = "CL_POSITION"

                If dt.Rows.Count > 0 Then
                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dt.Rows.Count, 5)
                    oTable.Range.ParagraphFormat.KeepWithNext = True
                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                    oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                    oTable.Columns.Item(2).Width = wordapp.InchesToPoints(5.8)
                    oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)

                    For i = 0 To dt.Rows.Count - 1
                        oTable.Rows.Item(i + 1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(i + 1).Shading
                            oTable.Cell(i + 1, 1).Range.Text = dv.Item(i)("Line#")
                            oTable.Cell(i + 1, 2).Range.Text = dv.Item(i)("Question")
                            If dv.Item(i)("HEADER") Then
                                If dv.Item(i)("Line#").ToString.Length = 1 Then
                                    .BackgroundPatternColor = Word.WdColor.wdColorBlack
                                Else
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                    .Texture = Word.WdTextureIndex.wdTexture30Percent
                                End If
                                oTable.Rows.Item(i + 1).Range.Font.Bold = True
                                oTable.Cell(i + 1, 3).Range.Text = "Yes"
                                oTable.Cell(i + 1, 4).Range.Text = "No"
                                oTable.Cell(i + 1, 5).Range.Text = "N/A"
                            Else
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                oTable.Cell(i + 1, 5).Range.Text = IIf(dv.Item(i)("N/A"), "X", "")

                                If dv.Item(i)("Line#").ToString.Trim = "1.13" AndAlso dv.Item(i)("NO") Then
                                    hasDesignatedOwner = False
                                End If

                            End If
                        End With
                    Next

                    If hasDesignatedOwner Then

                        If ThirdPartyOperator <> "OWNER" AndAlso ThirdPartyOperator <> String.Empty Then
                            hasDesignatedOwner = False
                        End If

                    End If

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()
                End If




                ' Spill and Overfill Prevention
                dt = oInsp.CheckListMaster.SpillTable
                dv = dt.DefaultView
                dv.Sort = "CL_POSITION"

                If dt.Rows.Count > 0 Then
                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dt.Rows.Count, 5)
                    oTable.Range.ParagraphFormat.KeepWithNext = True

                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                    oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                    oTable.Columns.Item(2).Width = wordapp.InchesToPoints(5.8)
                    oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)

                    Dim Show2_6 As Boolean = False
                    For i = 0 To dt.Rows.Count - 1
                        If dv.Item(i)("Line#") <> "2.6" OrElse Show2_6 Then
                            oTable.Rows.Item(i + 1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oTable.Rows.Item(i + 1).Shading
                                oTable.Cell(i + 1, 1).Range.Text = dv.Item(i)("Line#")
                                oTable.Cell(i + 1, 2).Range.Text = dv.Item(i)("Question")
                                If dv.Item(i)("HEADER") Then
                                    If dv.Item(i)("Line#").ToString.Length = 1 Then
                                        .BackgroundPatternColor = Word.WdColor.wdColorBlack
                                    Else
                                        .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                        .Texture = Word.WdTextureIndex.wdTexture30Percent
                                    End If
                                    oTable.Rows.Item(i + 1).Range.Font.Bold = True
                                    oTable.Cell(i + 1, 3).Range.Text = "Yes"
                                    oTable.Cell(i + 1, 4).Range.Text = "No"
                                    oTable.Cell(i + 1, 5).Range.Text = "N/A"
                                Else

                                    If dv.Item(i)("Line#") = "2.5" AndAlso dv.Item(i)("Yes") Then
                                        Show2_6 = True
                                    End If
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                    oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                    oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                    oTable.Cell(i + 1, 5).Range.Text = IIf(dv.Item(i)("N/A"), "X", "")
                                End If


                            End With
                        End If
                    Next

                    Dim r As Integer = 0

                    While r < oTable.Rows.Count

                        If oTable.Cell(r, 1).Range.Text.Replace(Chr(13), "").Replace(Chr(7), "").Trim = String.Empty Then
                            oTable.Rows.Item(r).Delete()
                            Threading.Thread.Sleep(200)
                        Else
                            r += 1
                        End If


                    End While

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()
                End If


                ' Corrosion Protection
                ds = oInsp.CheckListMaster.CPTable
                ds.Tables("CP").Columns("Line#").ColumnName = "LineNum"

                drows = ds.Tables("CP").Select("LineNum <= '3.4'")
                dt = New DataTable
                dt = ds.Tables("CP").Clone
                For Each dr In drows
                    dt.ImportRow(dr)
                Next
                dv = dt.DefaultView
                dv.Sort = "CL_POSITION"

                If drows.Length > 0 Then
                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, drows.Length, 5)
                    oTable.Range.ParagraphFormat.KeepWithNext = True

                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                    oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                    oTable.Columns.Item(2).Width = wordapp.InchesToPoints(5.8)
                    oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)

                    For i = 0 To drows.Length - 1
                        oTable.Rows.Item(i + 1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(i + 1).Shading
                            'dr = drows.GetValue(i)
                            oTable.Cell(i + 1, 1).Range.Text = dv.Item(i)("LineNum")
                            oTable.Cell(i + 1, 2).Range.Text = dv.Item(i)("Question")
                            If dv.Item(i)("HEADER") Then
                                If dv.Item(i)("LineNum").ToString.Length = 1 Then
                                    .BackgroundPatternColor = Word.WdColor.wdColorBlack
                                Else
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                    .Texture = Word.WdTextureIndex.wdTexture30Percent
                                End If
                                oTable.Rows.Item(i + 1).Range.Font.Bold = True
                                oTable.Cell(i + 1, 3).Range.Text = "Yes"
                                oTable.Cell(i + 1, 4).Range.Text = "No"
                                oTable.Cell(i + 1, 5).Range.Text = "N/A"
                            Else
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                oTable.Cell(i + 1, 5).Range.Text = IIf(dv.Item(i)("N/A"), "X", "")
                            End If
                        End With
                    Next

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()

                    drows = ds.Tables("CP").Select("LineNum = '3.4'")
                    If drows.Length > 0 Then
                        dtSub = ds.Tables("CPRect")

                        ' if no values are present for volts / amps / hours, create empty rows
                        If dtSub.Rows.Count = 0 Then
                            dr = dtSub.NewRow
                            dr("Volts") = DBNull.Value
                            dr("Amps") = DBNull.Value
                            dr("Hours") = DBNull.Value
                            dr("How Long") = DBNull.Value
                            dtSub.Rows.Add(dr)
                        End If

                        dvSub = dtSub.DefaultView
                        dvSub.Sort = "ID, QUESTION_ID"

                        oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dtSub.Rows.Count + 1, 5)
                        oTable.Range.ParagraphFormat.KeepWithNext = True

                        oTable.Range.Font.Name = "Arial"
                        oTable.Range.Font.Size = 8
                        oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                        oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                        oTable.Columns.Item(2).Width = wordapp.InchesToPoints(1.7)
                        oTable.Columns.Item(3).Width = wordapp.InchesToPoints(1.7)
                        oTable.Columns.Item(4).Width = wordapp.InchesToPoints(1.7)
                        oTable.Columns.Item(5).Width = wordapp.InchesToPoints(1.7)

                        oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        oTable.Cell(1, 1).Range.Text = ""
                        oTable.Cell(1, 2).Range.Text = "Volts"
                        oTable.Cell(1, 3).Range.Text = "Amps"
                        oTable.Cell(1, 4).Range.Text = "Hours"
                        oTable.Cell(1, 5).Range.Text = "How Long"
                        With oTable.Rows.Item(1).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture30Percent
                        End With

                        For i = 0 To dtSub.Rows.Count - 1
                            oTable.Rows.Item(i + 2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oTable.Rows.Item(i + 2).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 2, 1).Range.Text = ""

                                If dvSub.Item(i)("Volts") Is DBNull.Value Then
                                    oTable.Cell(i + 2, 2).Range.Text = String.Empty
                                ElseIf dvSub.Item(i)("Volts") = 0 Then
                                    oTable.Cell(i + 2, 2).Range.Text = String.Empty
                                Else
                                    oTable.Cell(i + 2, 2).Range.Text = dvSub.Item(i)("Volts")
                                End If

                                If dvSub.Item(i)("Amps") Is DBNull.Value Then
                                    oTable.Cell(i + 2, 3).Range.Text = String.Empty
                                ElseIf dvSub.Item(i)("Amps") = 0 Then
                                    oTable.Cell(i + 2, 3).Range.Text = String.Empty
                                Else
                                    oTable.Cell(i + 2, 3).Range.Text = dvSub.Item(i)("Amps")
                                End If

                                If dvSub.Item(i)("Hours") Is DBNull.Value Then
                                    oTable.Cell(i + 2, 4).Range.Text = String.Empty
                                ElseIf dvSub.Item(i)("Hours") = 0 Then
                                    oTable.Cell(i + 2, 4).Range.Text = String.Empty
                                Else
                                    oTable.Cell(i + 2, 4).Range.Text = dvSub.Item(i)("Hours")
                                End If

                                If dvSub.Item(i)("How Long") Is DBNull.Value Then
                                    oTable.Cell(i + 2, 5).Range.Text = String.Empty
                                ElseIf dvSub.Item(i)("How Long") = String.Empty Then
                                    oTable.Cell(i + 2, 5).Range.Text = String.Empty
                                Else
                                    oTable.Cell(i + 2, 5).Range.Text = dvSub.Item(i)("How Long")
                                End If
                            End With
                        Next

                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "<Space>"
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                End If

                drows = ds.Tables("CP").Select("LineNum > '3.4' and LineNum <= '3.5.4'")
                dt = New DataTable
                dt = ds.Tables("CP").Clone
                For Each dr In drows
                    dt.ImportRow(dr)
                Next
                dv = dt.DefaultView
                dv.Sort = "CL_POSITION"

                If drows.Length > 0 Then
                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, drows.Length, 5)
                    oTable.Range.ParagraphFormat.KeepWithNext = True

                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                    oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                    oTable.Columns.Item(2).Width = wordapp.InchesToPoints(5.8)
                    oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)

                    For i = 0 To drows.Length - 1
                        oTable.Rows.Item(i + 1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(i + 1).Shading
                            'dr = drows.GetValue(i)
                            oTable.Cell(i + 1, 1).Range.Text = dv.Item(i)("LineNum")
                            oTable.Cell(i + 1, 2).Range.Text = dv.Item(i)("Question")
                            If dv.Item(i)("HEADER") Then
                                If dv.Item(i)("LineNum").ToString.Length = 1 Then
                                    .BackgroundPatternColor = Word.WdColor.wdColorBlack
                                Else
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                    .Texture = Word.WdTextureIndex.wdTexture30Percent
                                End If
                                oTable.Rows.Item(i + 1).Range.Font.Bold = True
                                If dv.Item(i)("LineNum") <> "3.5.4" Then
                                    oTable.Cell(i + 1, 3).Range.Text = "Yes"
                                    oTable.Cell(i + 1, 4).Range.Text = "No"
                                    oTable.Cell(i + 1, 5).Range.Text = "N/A"

                                Else
                                    ' add cp readings cp tank/pipe/term tested by inspector
                                    dtSub = ds.Tables("CPTankInspectorTested")
                                    If dtSub.Rows.Count > 0 Then
                                        If dtSub.Rows(0)("Yes") Then
                                            oTable.Cell(i + 1, 3).Range.Text = "X"
                                            oTable.Cell(i + 1, 4).Range.Text = ""
                                            oTable.Cell(i + 1, 5).Range.Text = ""

                                        ElseIf dtSub.Rows(0)("No") Then
                                            oTable.Cell(i + 1, 3).Range.Text = ""
                                            oTable.Cell(i + 1, 4).Range.Text = "X"
                                            oTable.Cell(i + 1, 5).Range.Text = ""


                                        ElseIf dtSub.Rows(0)("N/A") Then
                                            oTable.Cell(i + 1, 3).Range.Text = ""
                                            oTable.Cell(i + 1, 4).Range.Text = ""
                                            oTable.Cell(i + 1, 5).Range.Text = "X"
                                        End If
                                    End If
                                End If
                                'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                            Else
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                oTable.Cell(i + 1, 5).Range.Text = IIf(dv.Item(i)("N/A"), "X", "")

                            End If
                        End With
                    Next

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()


                    ' add cp readings description of remote reference cell placement row
                    dtSub = ds.Tables("CPTankRemote")
                    dvSub = dtSub.DefaultView
                    If dtSub.Rows.Count > 0 Then
                        ' CP READINGS DESCRIPTION
                        oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dtSub.Rows.Count + 1, 1)
                        oTable.Range.ParagraphFormat.KeepWithNext = True

                        oTable.Range.Font.Name = "Arial"
                        oTable.Range.Font.Size = 8
                        oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                        oTable.Columns.Item(1).Width = wordapp.InchesToPoints(7.5)

                        oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        oTable.Cell(1, 1).Range.Text = "Description of Remote Reference Cell Placement"
                        With oTable.Rows.Item(1).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture30Percent
                        End With
                        For i = 0 To dtSub.Rows.Count - 1
                            oTable.Rows.Item(i + 2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oTable.Rows.Item(i + 2).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 2, 1).Range.Text = dvSub.Item(i)("Description of Remote Reference Cell Placement")
                            End With
                        Next

                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "<Space>"
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If

                    ' add cp readings rows
                    dtSub = ds.Tables("CPTank")
                    dvSub = dtSub.DefaultView
                    dvSub.Sort = "TANK_INDEX, LINE_NUMBER"
                    'dvSub.Sort = "Tank#, QUESTION_ID"

                    If dtSub.Rows.Count > 0 Then
                        oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dtSub.Rows.Count + 1, 10)
                        oTable.Range.ParagraphFormat.KeepWithNext = True

                        oTable.Range.Font.Name = "Arial"
                        oTable.Range.Font.Size = 8
                        oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                        oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                        oTable.Columns.Item(2).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.8)
                        oTable.Columns.Item(4).Width = wordapp.InchesToPoints(1.0)
                        oTable.Columns.Item(5).Width = wordapp.InchesToPoints(1.0)
                        oTable.Columns.Item(6).Width = wordapp.InchesToPoints(1.0)
                        oTable.Columns.Item(7).Width = wordapp.InchesToPoints(1.0)
                        oTable.Columns.Item(8).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(9).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(10).Width = wordapp.InchesToPoints(0.5)

                        oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        oTable.Cell(1, 1).Range.Text = ""
                        oTable.Cell(1, 2).Range.Text = "Tank#"
                        oTable.Cell(1, 3).Range.Text = "Fuel Type"
                        oTable.Cell(1, 4).Range.Text = "Contact Point"
                        oTable.Cell(1, 5).Range.Text = "Local Reference Cell Placement"
                        oTable.Cell(1, 6).Range.Text = "Local/On"
                        oTable.Cell(1, 7).Range.Text = "Remote/Off"
                        oTable.Cell(1, 8).Range.Text = "Pass"
                        oTable.Cell(1, 9).Range.Text = "Fail"
                        oTable.Cell(1, 10).Range.Text = "Incon"
                        With oTable.Rows.Item(1).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture30Percent
                        End With
                        For i = 0 To dtSub.Rows.Count - 1
                            oTable.Rows.Item(i + 2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oTable.Rows.Item(i + 2).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 2, 1).Range.Text = dvSub.Item(i)("Line#")
                                oTable.Cell(i + 2, 2).Range.Text = dvSub.Item(i)("Tank#")
                                oTable.Cell(i + 2, 3).Range.Text = dvSub.Item(i)("Fuel Type")
                                oTable.Cell(i + 2, 4).Range.Text = dvSub.Item(i)("Contact Point")
                                oTable.Cell(i + 2, 5).Range.Text = dvSub.Item(i)("Local Reference Cell Placement")
                                oTable.Cell(i + 2, 6).Range.Text = dvSub.Item(i)("Local/On")
                                oTable.Cell(i + 2, 7).Range.Text = dvSub.Item(i)("Remote/Off")
                                oTable.Cell(i + 2, 8).Range.Text = IIf(dvSub.Item(i)("Pass"), "X", "")
                                oTable.Cell(i + 2, 9).Range.Text = IIf(dvSub.Item(i)("Fail"), "X", "")
                                oTable.Cell(i + 2, 10).Range.Text = IIf(dvSub.Item(i)("Incon"), "X", "")
                            End With
                        Next

                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "<Space>"
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                End If

                drows = ds.Tables("CP").Select("LineNum > '3.5.4' and LineNum <= '3.6.3'")
                dt = New DataTable
                dt = ds.Tables("CP").Clone
                For Each dr In drows
                    dt.ImportRow(dr)
                Next
                dv = dt.DefaultView
                dv.Sort = "CL_POSITION"

                If drows.Length > 0 Then
                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, drows.Length, 5)
                    oTable.Range.ParagraphFormat.KeepWithNext = True

                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                    oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                    oTable.Columns.Item(2).Width = wordapp.InchesToPoints(5.8)
                    oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)

                    For i = 0 To drows.Length - 1
                        oTable.Rows.Item(i + 1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(i + 1).Shading
                            'dr = drows.GetValue(i)
                            oTable.Cell(i + 1, 1).Range.Text = dv.Item(i)("LineNum")
                            oTable.Cell(i + 1, 2).Range.Text = dv.Item(i)("Question")
                            If dv.Item(i)("HEADER") Then
                                If dv.Item(i)("LineNum").ToString.Length = 1 Then
                                    .BackgroundPatternColor = Word.WdColor.wdColorBlack
                                Else
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                    .Texture = Word.WdTextureIndex.wdTexture30Percent
                                End If
                                oTable.Rows.Item(i + 1).Range.Font.Bold = True
                                If dv.Item(i)("LineNum") <> "3.6.3" Then
                                    oTable.Cell(i + 1, 3).Range.Text = "Yes"
                                    oTable.Cell(i + 1, 4).Range.Text = "No"
                                    oTable.Cell(i + 1, 5).Range.Text = "N/A"

                                Else
                                    ' add cp readings cp tank/pipe/term tested by inspector
                                    dtSub = ds.Tables("CPPipeInspectorTested")
                                    If dtSub.Rows.Count > 0 Then
                                        If dtSub.Rows(0)("Yes") Then
                                            oTable.Cell(i + 1, 3).Range.Text = "X"
                                            oTable.Cell(i + 1, 4).Range.Text = ""
                                            oTable.Cell(i + 1, 5).Range.Text = ""

                                        ElseIf dtSub.Rows(0)("No") Then
                                            oTable.Cell(i + 1, 3).Range.Text = ""
                                            oTable.Cell(i + 1, 4).Range.Text = "X"
                                            oTable.Cell(i + 1, 5).Range.Text = ""

                                        ElseIf dtSub.Rows(0)("N/A") Then

                                            oTable.Cell(i + 1, 3).Range.Text = ""
                                            oTable.Cell(i + 1, 4).Range.Text = ""
                                            oTable.Cell(i + 1, 5).Range.Text = "X"


                                        End If
                                    End If
                                End If
                                'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                            Else
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                oTable.Cell(i + 1, 5).Range.Text = IIf(dv.Item(i)("N/A"), "X", "")

                            End If
                        End With
                    Next

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()


                    ' add cp readings description of remote reference cell placement row
                    dtSub = ds.Tables("CPPipeRemote")
                    dvSub = dtSub.DefaultView
                    If dtSub.Rows.Count > 0 Then
                        ' CP READINGS DESCRIPTION
                        oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dtSub.Rows.Count + 1, 1)
                        oTable.Range.ParagraphFormat.KeepWithNext = True

                        oTable.Range.Font.Name = "Arial"
                        oTable.Range.Font.Size = 8
                        oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                        oTable.Columns.Item(1).Width = wordapp.InchesToPoints(7.5)

                        oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        oTable.Cell(1, 1).Range.Text = "Description of Remote Reference Cell Placement"
                        With oTable.Rows.Item(1).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture30Percent
                        End With
                        For i = 0 To dtSub.Rows.Count - 1
                            oTable.Rows.Item(i + 2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oTable.Rows.Item(i + 2).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 2, 1).Range.Text = dvSub.Item(i)("Description of Remote Reference Cell Placement")
                            End With
                        Next

                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "<Space>"
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If

                    ' add cp readings rows
                    dtSub = ds.Tables("CPPipe")
                    dvSub = dtSub.DefaultView
                    dvSub.Sort = "PIPE_INDEX, LINE_NUMBER"
                    'dvSub.Sort = "Pipe#, QUESTION_ID"

                    If dtSub.Rows.Count > 0 Then
                        oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dtSub.Rows.Count + 1, 10)
                        oTable.Range.ParagraphFormat.KeepWithNext = True

                        oTable.Range.Font.Name = "Arial"
                        oTable.Range.Font.Size = 8
                        oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                        oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                        oTable.Columns.Item(2).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.8)
                        oTable.Columns.Item(4).Width = wordapp.InchesToPoints(1.0)
                        oTable.Columns.Item(5).Width = wordapp.InchesToPoints(1.0)
                        oTable.Columns.Item(6).Width = wordapp.InchesToPoints(1.0)
                        oTable.Columns.Item(7).Width = wordapp.InchesToPoints(1.0)
                        oTable.Columns.Item(8).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(9).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(10).Width = wordapp.InchesToPoints(0.5)

                        oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        oTable.Cell(1, 1).Range.Text = ""
                        oTable.Cell(1, 2).Range.Text = "Pipe#"
                        oTable.Cell(1, 3).Range.Text = "Fuel Type"
                        oTable.Cell(1, 4).Range.Text = "Contact Point"
                        oTable.Cell(1, 5).Range.Text = "Local Reference Cell Placement"
                        oTable.Cell(1, 6).Range.Text = "Local/On"
                        oTable.Cell(1, 7).Range.Text = "Remote/Off"
                        oTable.Cell(1, 8).Range.Text = "Pass"
                        oTable.Cell(1, 9).Range.Text = "Fail"
                        oTable.Cell(1, 10).Range.Text = "Incon"
                        With oTable.Rows.Item(1).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture30Percent
                        End With
                        For i = 0 To dtSub.Rows.Count - 1
                            oTable.Rows.Item(i + 2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oTable.Rows.Item(i + 2).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 2, 1).Range.Text = dvSub.Item(i)("Line#")
                                oTable.Cell(i + 2, 2).Range.Text = dvSub.Item(i)("Pipe#")
                                oTable.Cell(i + 2, 3).Range.Text = dvSub.Item(i)("Fuel Type")
                                oTable.Cell(i + 2, 4).Range.Text = dvSub.Item(i)("Contact Point")
                                oTable.Cell(i + 2, 5).Range.Text = dvSub.Item(i)("Local Reference Cell Placement")
                                oTable.Cell(i + 2, 6).Range.Text = dvSub.Item(i)("Local/On")
                                oTable.Cell(i + 2, 7).Range.Text = dvSub.Item(i)("Remote/Off")
                                oTable.Cell(i + 2, 8).Range.Text = IIf(dvSub.Item(i)("Pass"), "X", "")
                                oTable.Cell(i + 2, 9).Range.Text = IIf(dvSub.Item(i)("Fail"), "X", "")
                                oTable.Cell(i + 2, 10).Range.Text = IIf(dvSub.Item(i)("Incon"), "X", "")
                            End With
                        Next

                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "<Space>"
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                End If

                drows = ds.Tables("CP").Select("LineNum > '3.6.3' and LineNum <= '3.7.6'")
                dt = New DataTable
                dt = ds.Tables("CP").Clone
                For Each dr In drows
                    dt.ImportRow(dr)
                Next
                dv = dt.DefaultView
                dv.Sort = "CL_POSITION"

                If drows.Length > 0 Then
                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, drows.Length, 5)
                    oTable.Range.ParagraphFormat.KeepWithNext = True

                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                    oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                    oTable.Columns.Item(2).Width = wordapp.InchesToPoints(5.8)
                    oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)

                    For i = 0 To drows.Length - 1
                        oTable.Rows.Item(i + 1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(i + 1).Shading
                            'dr = drows.GetValue(i)
                            oTable.Cell(i + 1, 1).Range.Text = dv.Item(i)("LineNum")
                            oTable.Cell(i + 1, 2).Range.Text = dv.Item(i)("Question")
                            If dv.Item(i)("HEADER") Then
                                If dv.Item(i)("LineNum").ToString.Length = 1 Then
                                    .BackgroundPatternColor = Word.WdColor.wdColorBlack
                                Else
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                    .Texture = Word.WdTextureIndex.wdTexture30Percent
                                End If
                                oTable.Rows.Item(i + 1).Range.Font.Bold = True
                                If dv.Item(i)("LineNum") <> "3.7.6" Then
                                    oTable.Cell(i + 1, 3).Range.Text = "Yes"
                                    oTable.Cell(i + 1, 4).Range.Text = "No"
                                    oTable.Cell(i + 1, 5).Range.Text = "N/A"

                                Else
                                    ' add cp readings cp tank/pipe/term tested by inspector
                                    dtSub = ds.Tables("CPTermInspectorTested")
                                    If dtSub.Rows.Count > 0 Then
                                        If dtSub.Rows(0)("Yes") Then
                                            oTable.Cell(i + 1, 3).Range.Text = "X"
                                            oTable.Cell(i + 1, 4).Range.Text = ""
                                            oTable.Cell(i + 1, 5).Range.Text = ""

                                        ElseIf dtSub.Rows(0)("No") Then
                                            oTable.Cell(i + 1, 3).Range.Text = ""
                                            oTable.Cell(i + 1, 4).Range.Text = "X"
                                            oTable.Cell(i + 1, 5).Range.Text = ""

                                        ElseIf dtSub.Rows(0)("N/A") Then
                                            oTable.Cell(i + 1, 3).Range.Text = ""
                                            oTable.Cell(i + 1, 4).Range.Text = ""
                                            oTable.Cell(i + 1, 5).Range.Text = "X"
                                        End If
                                    End If
                                End If
                                'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                            Else
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                oTable.Cell(i + 1, 5).Range.Text = IIf(dv.Item(i)("N/A"), "X", "")

                                'oTable.Cell(i + 1, 5).Range.Text = dv.Item(i)("CCAT")
                            End If
                        End With
                    Next

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()

                    dtSub = ds.Tables("CPTermRemote")
                    dvSub = dtSub.DefaultView
                    If dtSub.Rows.Count > 0 Then
                        ' CP READINGS DESCRIPTION
                        oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dtSub.Rows.Count + 1, 1)
                        oTable.Range.ParagraphFormat.KeepWithNext = True

                        oTable.Range.Font.Name = "Arial"
                        oTable.Range.Font.Size = 8
                        oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                        oTable.Columns.Item(1).Width = wordapp.InchesToPoints(7.5)

                        oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        oTable.Cell(1, 1).Range.Text = "Description of Remote Reference Cell Placement"
                        With oTable.Rows.Item(1).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture30Percent
                        End With
                        For i = 0 To dtSub.Rows.Count - 1
                            oTable.Rows.Item(i + 2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oTable.Rows.Item(i + 2).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 2, 1).Range.Text = dvSub.Item(i)("Description of Remote Reference Cell Placement")
                            End With
                        Next

                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "<Space>"
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If

                    ' add cp readings rows
                    dtSub = ds.Tables("CPTerm")
                    dvSub = dtSub.DefaultView
                    dvSub.Sort = "TERM_INDEX, LINE_NUMBER"
                    'dvSub.Sort = "Term#, QUESTION_ID"

                    If dtSub.Rows.Count > 0 Then
                        oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dtSub.Rows.Count + 1, 10)
                        oTable.Range.ParagraphFormat.KeepWithNext = True

                        oTable.Range.Font.Name = "Arial"
                        oTable.Range.Font.Size = 8
                        oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                        oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                        oTable.Columns.Item(2).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.8)
                        oTable.Columns.Item(4).Width = wordapp.InchesToPoints(1.0)
                        oTable.Columns.Item(5).Width = wordapp.InchesToPoints(1.0)
                        oTable.Columns.Item(6).Width = wordapp.InchesToPoints(1.0)
                        oTable.Columns.Item(7).Width = wordapp.InchesToPoints(1.0)
                        oTable.Columns.Item(8).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(9).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(10).Width = wordapp.InchesToPoints(0.5)

                        oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        oTable.Cell(1, 1).Range.Text = ""
                        oTable.Cell(1, 2).Range.Text = "Term#"
                        oTable.Cell(1, 3).Range.Text = "Fuel Type"
                        oTable.Cell(1, 4).Range.Text = "Contact Point"
                        oTable.Cell(1, 5).Range.Text = "Local Reference Cell Placement"
                        oTable.Cell(1, 6).Range.Text = "Local/On"
                        oTable.Cell(1, 7).Range.Text = "Remote/Off"
                        oTable.Cell(1, 8).Range.Text = "Pass"
                        oTable.Cell(1, 9).Range.Text = "Fail"
                        oTable.Cell(1, 10).Range.Text = "Incon"
                        With oTable.Rows.Item(1).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture30Percent
                        End With
                        For i = 0 To dtSub.Rows.Count - 1
                            oTable.Rows.Item(i + 2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oTable.Rows.Item(i + 2).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 2, 1).Range.Text = dvSub.Item(i)("Line#")
                                oTable.Cell(i + 2, 2).Range.Text = dvSub.Item(i)("Term#")
                                oTable.Cell(i + 2, 3).Range.Text = dvSub.Item(i)("Fuel Type")
                                oTable.Cell(i + 2, 4).Range.Text = dvSub.Item(i)("Contact Point")
                                oTable.Cell(i + 2, 5).Range.Text = dvSub.Item(i)("Local Reference Cell Placement")
                                oTable.Cell(i + 2, 6).Range.Text = dvSub.Item(i)("Local/On")
                                oTable.Cell(i + 2, 7).Range.Text = dvSub.Item(i)("Remote/Off")
                                oTable.Cell(i + 2, 8).Range.Text = IIf(dvSub.Item(i)("Pass"), "X", "")
                                oTable.Cell(i + 2, 9).Range.Text = IIf(dvSub.Item(i)("Fail"), "X", "")
                                oTable.Cell(i + 2, 10).Range.Text = IIf(dvSub.Item(i)("Incon"), "X", "")
                            End With
                        Next
                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "<Space>"
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                End If

                ' Tank Leak
                ds = oInsp.CheckListMaster.TankLeakTable
                ds.Tables("TankLeak").Columns("Line#").ColumnName = "LineNum"

                drows = ds.Tables("TankLeak").Select("LineNum <= '4.2.8'")
                dt = New DataTable
                dt = ds.Tables("TankLeak").Clone
                For Each dr In drows
                    dt.ImportRow(dr)
                Next
                dv = dt.DefaultView
                dv.Sort = "CL_POSITION"

                If drows.Length > 0 Then
                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, drows.Length, 5)
                    oTable.Range.ParagraphFormat.KeepWithNext = True

                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                    oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                    oTable.Columns.Item(2).Width = wordapp.InchesToPoints(5.8)
                    oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)

                    For i = 0 To drows.Length - 1
                        oTable.Rows.Item(i + 1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(i + 1).Shading
                            'dr = drows.GetValue(i)
                            oTable.Cell(i + 1, 1).Range.Text = dv.Item(i)("LineNum")
                            oTable.Cell(i + 1, 2).Range.Text = dv.Item(i)("Question")
                            If dv.Item(i)("HEADER") Then
                                If dv.Item(i)("LineNum").ToString.Length = 1 Then
                                    .BackgroundPatternColor = Word.WdColor.wdColorBlack
                                Else
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                    .Texture = Word.WdTextureIndex.wdTexture30Percent
                                End If
                                oTable.Rows.Item(i + 1).Range.Font.Bold = True
                                If dv.Item(i)("LineNum") <> "4.2.8" Then
                                    oTable.Cell(i + 1, 3).Range.Text = "Yes"
                                    oTable.Cell(i + 1, 4).Range.Text = "No"
                                    oTable.Cell(i + 1, 5).Range.Text = "N/A"

                                End If
                                'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                            Else
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                oTable.Cell(i + 1, 5).Range.Text = IIf(dv.Item(i)("N/A"), "X", "")

                                'oTable.Cell(i + 1, 5).Range.Text = dv.Item(i)("CCAT")
                            End If
                        End With
                    Next

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()

                    dtSub = ds.Tables("Well")
                    dvSub = dtSub.DefaultView
                    dvSub.Sort = "Well#, LINE_NUMBER"

                    If dtSub.Rows.Count > 0 Then
                        oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dtSub.Rows.Count + 1, 10)
                        oTable.Range.ParagraphFormat.KeepWithNext = True

                        oTable.Range.Font.Name = "Arial"
                        oTable.Range.Font.Size = 8
                        oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                        oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                        oTable.Columns.Item(2).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(6).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(7).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(8).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(9).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(10).Width = wordapp.InchesToPoints(2.8)

                        oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        oTable.Cell(1, 1).Range.Text = ""
                        oTable.Cell(1, 2).Range.Text = "Well#"
                        oTable.Cell(1, 2).Range.Font.Size = 7
                        oTable.Cell(1, 3).Range.Text = "Well" + vbCrLf + "Depth"
                        oTable.Cell(1, 3).Range.Font.Size = 7
                        oTable.Cell(1, 4).Range.Text = "Depth" + vbCrLf + "to" + vbCrLf + "Water"
                        oTable.Cell(1, 4).Range.Font.Size = 7
                        oTable.Cell(1, 5).Range.Text = "Depth" + vbCrLf + "to" + vbCrLf + "Slots"
                        oTable.Cell(1, 5).Range.Font.Size = 7
                        oTable.Cell(1, 6).Range.Text = "Surface" + vbCrLf + "Sealed" + vbCrLf + "Yes"
                        oTable.Cell(1, 6).Range.Font.Size = 7
                        oTable.Cell(1, 7).Range.Text = "Surface" + vbCrLf + "Sealed" + vbCrLf + "No"
                        oTable.Cell(1, 7).Range.Font.Size = 7
                        oTable.Cell(1, 8).Range.Text = "Well" + vbCrLf + "Caps" + vbCrLf + "Yes"
                        oTable.Cell(1, 8).Range.Font.Size = 7
                        oTable.Cell(1, 9).Range.Text = "Well" + vbCrLf + "Caps" + vbCrLf + "No"
                        oTable.Cell(1, 9).Range.Font.Size = 7
                        oTable.Cell(1, 10).Range.Text = "Inspector's Observations"
                        oTable.Cell(1, 10).Range.Font.Size = 7
                        With oTable.Rows.Item(1).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture30Percent
                        End With
                        For i = 0 To dtSub.Rows.Count - 1
                            oTable.Rows.Item(i + 2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oTable.Rows.Item(i + 2).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 2, 1).Range.Text = dvSub.Item(i)("Line#")
                                oTable.Cell(i + 2, 2).Range.Text = IIf(dvSub.Item(i)("Well#") = 0, String.Empty, dvSub.Item(i)("Well#"))
                                oTable.Cell(i + 2, 3).Range.Text = dvSub.Item(i)("Well Depth")
                                oTable.Cell(i + 2, 4).Range.Text = dvSub.Item(i)("Depth to" + vbCrLf + "Water")
                                oTable.Cell(i + 2, 5).Range.Text = dvSub.Item(i)("Depth to" + vbCrLf + "Slots")
                                oTable.Cell(i + 2, 6).Range.Text = IIf(dvSub.Item(i)("Surface Sealed" + vbCrLf + "Yes"), "X", "")
                                oTable.Cell(i + 2, 7).Range.Text = IIf(dvSub.Item(i)("Surface Sealed" + vbCrLf + "No"), "X", "")
                                oTable.Cell(i + 2, 8).Range.Text = IIf(dvSub.Item(i)("Well Caps" + vbCrLf + "Yes"), "X", "")
                                oTable.Cell(i + 2, 9).Range.Text = IIf(dvSub.Item(i)("Well Caps" + vbCrLf + "No"), "X", "")
                                oTable.Cell(i + 2, 10).Range.Text = dvSub.Item(i)("Inspector's Observations")
                            End With
                        Next

                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "<Space>"
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                End If

                drows = ds.Tables("TankLeak").Select("LineNum > '4.2.8'")
                dt = New DataTable
                dt = ds.Tables("TankLeak").Clone
                For Each dr In drows
                    dt.ImportRow(dr)
                Next
                dv = dt.DefaultView
                dv.Sort = "CL_POSITION"

                If drows.Length > 0 Then
                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, drows.Length, 5)
                    oTable.Range.ParagraphFormat.KeepWithNext = True

                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                    oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                    oTable.Columns.Item(2).Width = wordapp.InchesToPoints(5.8)
                    oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)

                    For i = 0 To drows.Length - 1
                        oTable.Rows.Item(i + 1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(i + 1).Shading
                            'dr = drows.GetValue(i)
                            oTable.Cell(i + 1, 1).Range.Text = dv.Item(i)("LineNum")
                            oTable.Cell(i + 1, 2).Range.Text = dv.Item(i)("Question")
                            If dv.Item(i)("HEADER") Then
                                If dv.Item(i)("LineNum").ToString.Length = 1 Then
                                    .BackgroundPatternColor = Word.WdColor.wdColorBlack
                                Else
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                    .Texture = Word.WdTextureIndex.wdTexture30Percent
                                End If
                                oTable.Rows.Item(i + 1).Range.Font.Bold = True
                                oTable.Cell(i + 1, 3).Range.Text = "Yes"
                                oTable.Cell(i + 1, 4).Range.Text = "No"
                                oTable.Cell(i + 1, 5).Range.Text = "N/A"
                            Else
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                oTable.Cell(i + 1, 5).Range.Text = IIf(dv.Item(i)("N/A"), "X", "")

                            End If
                        End With
                    Next

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()
                End If

                ' Pipe Leak
                ds = oInsp.CheckListMaster.PipeLeakTable
                ds.Tables("PipeLeak").Columns("Line#").ColumnName = "LineNum"

                drows = ds.Tables("PipeLeak").Select("LineNum <= '5.2.8'")
                dt = New DataTable
                dt = ds.Tables("PipeLeak").Clone
                For Each dr In drows
                    dt.ImportRow(dr)
                Next
                dv = dt.DefaultView
                dv.Sort = "CL_POSITION"

                If drows.Length > 0 Then
                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, drows.Length, 5)
                    oTable.Range.ParagraphFormat.KeepWithNext = True

                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                    oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                    oTable.Columns.Item(2).Width = wordapp.InchesToPoints(5.8)
                    oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)

                    For i = 0 To drows.Length - 1
                        oTable.Rows.Item(i + 1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(i + 1).Shading
                            'dr = drows.GetValue(i)
                            oTable.Cell(i + 1, 1).Range.Text = dv.Item(i)("LineNum")
                            oTable.Cell(i + 1, 2).Range.Text = dv.Item(i)("Question")
                            If dv.Item(i)("HEADER") Then
                                If dv.Item(i)("LineNum").ToString.Length = 1 Then
                                    .BackgroundPatternColor = Word.WdColor.wdColorBlack
                                Else
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                    .Texture = Word.WdTextureIndex.wdTexture30Percent
                                End If
                                oTable.Rows.Item(i + 1).Range.Font.Bold = True
                                If dv.Item(i)("LineNum") <> "5.2.8" Then
                                    oTable.Cell(i + 1, 3).Range.Text = "Yes"
                                    oTable.Cell(i + 1, 4).Range.Text = "No"
                                    oTable.Cell(i + 1, 5).Range.Text = "N/A"

                                End If
                            Else
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                oTable.Cell(i + 1, 5).Range.Text = IIf(dv.Item(i)("N/A"), "X", "")
                            End If
                        End With
                    Next

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()

                    dtSub = ds.Tables("Well")
                    dvSub = dtSub.DefaultView
                    dvSub.Sort = "Well#, LINE_NUMBER"

                    If dtSub.Rows.Count > 0 Then
                        oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dtSub.Rows.Count + 1, 10)
                        oTable.Range.ParagraphFormat.KeepWithNext = True

                        oTable.Range.Font.Name = "Arial"
                        oTable.Range.Font.Size = 8
                        oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                        oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                        oTable.Columns.Item(2).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(6).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(7).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(8).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(9).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(10).Width = wordapp.InchesToPoints(2.8)

                        oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        oTable.Cell(1, 1).Range.Text = ""
                        oTable.Cell(1, 2).Range.Text = "Well#"
                        oTable.Cell(1, 2).Range.Font.Size = 7
                        oTable.Cell(1, 3).Range.Text = "Well" + vbCrLf + "Depth"
                        oTable.Cell(1, 3).Range.Font.Size = 7
                        oTable.Cell(1, 4).Range.Text = "Depth" + vbCrLf + "to" + vbCrLf + "Water"
                        oTable.Cell(1, 4).Range.Font.Size = 7
                        oTable.Cell(1, 5).Range.Text = "Depth" + vbCrLf + "to" + vbCrLf + "Slots"
                        oTable.Cell(1, 5).Range.Font.Size = 7
                        oTable.Cell(1, 6).Range.Text = "Surface" + vbCrLf + "Sealed" + vbCrLf + "Yes"
                        oTable.Cell(1, 6).Range.Font.Size = 7
                        oTable.Cell(1, 7).Range.Text = "Surface" + vbCrLf + "Sealed" + vbCrLf + "No"
                        oTable.Cell(1, 7).Range.Font.Size = 7
                        oTable.Cell(1, 8).Range.Text = "Well" + vbCrLf + "Caps" + vbCrLf + "Yes"
                        oTable.Cell(1, 8).Range.Font.Size = 7
                        oTable.Cell(1, 9).Range.Text = "Well" + vbCrLf + "Caps" + vbCrLf + "No"
                        oTable.Cell(1, 9).Range.Font.Size = 7
                        oTable.Cell(1, 10).Range.Text = "Inspector's Observations"
                        oTable.Cell(1, 10).Range.Font.Size = 7
                        With oTable.Rows.Item(1).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture30Percent
                        End With
                        For i = 0 To dtSub.Rows.Count - 1
                            oTable.Rows.Item(i + 2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oTable.Rows.Item(i + 2).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 2, 1).Range.Text = dvSub.Item(i)("Line#")
                                oTable.Cell(i + 2, 2).Range.Text = IIf(dvSub.Item(i)("Well#") = 0, String.Empty, dvSub.Item(i)("Well#"))
                                oTable.Cell(i + 2, 3).Range.Text = dvSub.Item(i)("Well Depth")
                                oTable.Cell(i + 2, 4).Range.Text = dvSub.Item(i)("Depth to" + vbCrLf + "Water")
                                oTable.Cell(i + 2, 5).Range.Text = dvSub.Item(i)("Depth to" + vbCrLf + "Slots")
                                oTable.Cell(i + 2, 6).Range.Text = IIf(dvSub.Item(i)("Surface Sealed" + vbCrLf + "Yes"), "X", "")
                                oTable.Cell(i + 2, 7).Range.Text = IIf(dvSub.Item(i)("Surface Sealed" + vbCrLf + "No"), "X", "")
                                oTable.Cell(i + 2, 8).Range.Text = IIf(dvSub.Item(i)("Well Caps" + vbCrLf + "Yes"), "X", "")
                                oTable.Cell(i + 2, 9).Range.Text = IIf(dvSub.Item(i)("Well Caps" + vbCrLf + "No"), "X", "")
                                oTable.Cell(i + 2, 10).Range.Text = dvSub.Item(i)("Inspector's Observations")
                            End With
                        Next

                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "<Space>"
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    End If
                End If

                drows = ds.Tables("PipeLeak").Select("LineNum > '5.2.8'")
                dt = New DataTable
                dt = ds.Tables("PipeLeak").Clone
                For Each dr In drows
                    dt.ImportRow(dr)
                Next
                dv = dt.DefaultView
                dv.Sort = "CL_POSITION"

                If drows.Length > 0 Then
                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, drows.Length, 5)
                    oTable.Range.ParagraphFormat.KeepWithNext = True

                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                    oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                    oTable.Columns.Item(2).Width = wordapp.InchesToPoints(5.8)
                    oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)

                    For i = 0 To drows.Length - 1
                        oTable.Rows.Item(i + 1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(i + 1).Shading
                            'dr = drows.GetValue(i)
                            oTable.Cell(i + 1, 1).Range.Text = dv.Item(i)("LineNum")
                            oTable.Cell(i + 1, 2).Range.Text = dv.Item(i)("Question")
                            If dv.Item(i)("HEADER") Then
                                If dv.Item(i)("LineNum").ToString.Length = 1 Then
                                    .BackgroundPatternColor = Word.WdColor.wdColorBlack
                                Else
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                    .Texture = Word.WdTextureIndex.wdTexture30Percent
                                End If
                                oTable.Rows.Item(i + 1).Range.Font.Bold = True
                                oTable.Cell(i + 1, 3).Range.Text = "Yes"
                                oTable.Cell(i + 1, 4).Range.Text = "No"
                                oTable.Cell(i + 1, 5).Range.Text = "N/A"
                            Else
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                oTable.Cell(i + 1, 5).Range.Text = IIf(dv.Item(i)("N/A"), "X", "")
                            End If
                        End With
                    Next

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()
                End If

                ' CAT Leak
                dt = oInsp.CheckListMaster.CATLeakTable
                dv = dt.DefaultView
                dv.Sort = "CL_POSITION"

                If dt.Rows.Count > 0 Then
                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dt.Rows.Count, 5)
                    oTable.Range.ParagraphFormat.KeepWithNext = True

                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                    oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                    oTable.Columns.Item(2).Width = wordapp.InchesToPoints(5.8)
                    oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)

                    For i = 0 To dt.Rows.Count - 1
                        oTable.Rows.Item(i + 1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(i + 1).Shading
                            oTable.Cell(i + 1, 1).Range.Text = dv.Item(i)("Line#")
                            oTable.Cell(i + 1, 2).Range.Text = dv.Item(i)("Question")
                            If dv.Item(i)("HEADER") Then
                                If dv.Item(i)("Line#").ToString.Length = 1 Then
                                    .BackgroundPatternColor = Word.WdColor.wdColorBlack
                                Else
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                    .Texture = Word.WdTextureIndex.wdTexture30Percent
                                End If
                                oTable.Rows.Item(i + 1).Range.Font.Bold = True
                                oTable.Cell(i + 1, 3).Range.Text = "Yes"
                                oTable.Cell(i + 1, 4).Range.Text = "No"
                                oTable.Cell(i + 1, 5).Range.Text = "N/A"
                            Else
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                oTable.Cell(i + 1, 5).Range.Text = IIf(dv.Item(i)("N/A"), "X", "")
                            End If
                        End With
                    Next

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()
                End If

                ' Visual
                dt = oInsp.CheckListMaster.VisualTable
                dv = dt.DefaultView
                dv.Sort = "CL_POSITION"

                If dt.Rows.Count > 0 Then
                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dt.Rows.Count, 5)
                    oTable.Range.ParagraphFormat.KeepWithNext = True

                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                    oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                    oTable.Columns.Item(2).Width = wordapp.InchesToPoints(5.8)
                    oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)

                    For i = 0 To dt.Rows.Count - 1
                        oTable.Rows.Item(i + 1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(i + 1).Shading
                            oTable.Cell(i + 1, 1).Range.Text = dv.Item(i)("Line#")
                            oTable.Cell(i + 1, 2).Range.Text = dv.Item(i)("Question")
                            If dv.Item(i)("HEADER") Then
                                If dv.Item(i)("Line#").ToString.Length = 1 Then
                                    .BackgroundPatternColor = Word.WdColor.wdColorBlack
                                Else
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                    .Texture = Word.WdTextureIndex.wdTexture30Percent
                                End If
                                oTable.Rows.Item(i + 1).Range.Font.Bold = True
                                oTable.Cell(i + 1, 3).Range.Text = "Yes"
                                oTable.Cell(i + 1, 4).Range.Text = "No"
                                oTable.Cell(i + 1, 5).Range.Text = "N/A"
                            Else
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                oTable.Cell(i + 1, 5).Range.Text = IIf(dv.Item(i)("N/A"), "X", "")
                            End If
                        End With
                    Next

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()
                End If

                ' TOS Tanks
                dt = oInsp.CheckListMaster.TOSTable
                dv = dt.DefaultView
                dv.Sort = "CL_POSITION"

                If dt.Rows.Count > 0 Then
                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dt.Rows.Count, 5)
                    oTable.Range.ParagraphFormat.KeepWithNext = True

                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                    oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                    oTable.Columns.Item(2).Width = wordapp.InchesToPoints(5.8)
                    oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)

                    For i = 0 To dt.Rows.Count - 1
                        oTable.Rows.Item(i + 1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(i + 1).Shading
                            oTable.Cell(i + 1, 1).Range.Text = dv.Item(i)("Line#")
                            oTable.Cell(i + 1, 2).Range.Text = dv.Item(i)("Question")
                            If dv.Item(i)("HEADER") Then
                                If dv.Item(i)("Line#").ToString.Length = 1 Then
                                    .BackgroundPatternColor = Word.WdColor.wdColorBlack
                                Else
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                    .Texture = Word.WdTextureIndex.wdTexture30Percent
                                End If
                                oTable.Rows.Item(i + 1).Range.Font.Bold = True
                                oTable.Cell(i + 1, 3).Range.Text = "Yes"
                                oTable.Cell(i + 1, 4).Range.Text = "No"
                                oTable.Cell(i + 1, 5).Range.Text = "N/A"
                            Else
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                oTable.Cell(i + 1, 5).Range.Text = IIf(dv.Item(i)("N/A"), "X", "")
                            End If
                        End With
                    Next

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()


                    ' page break
                    wordapp.Selection.GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:=.Bookmarks.Item("\endofdoc").Name)
                    wordapp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)


                End If



                ' Inspection Comments
                'Dim strComments As String = oInsp.CheckListMaster.InspectionComments.InsComments
                Dim nRows As Integer = 0
                If oInsp.CheckListMaster.InspectionComments.InsComments = String.Empty Then
                    'nRows = 15
                    nRows = 38
                Else
                    nRows = 3
                End If

                ' page break
                wordapp.Selection.GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:=.Bookmarks.Item("\endofdoc").Name)
                wordapp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)

                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, nRows, 1)
                oTable.Range.ParagraphFormat.KeepWithNext = True

                oTable.Range.Font.Name = "Arial"
                oTable.Range.Font.Size = 8
                oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                oTable.Columns.Item(1).Width = wordapp.InchesToPoints(7.5)
                With oTable.Rows.Shading
                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                End With

                ' Heading
                oTable.Rows.Item(1).Range.Text = "8             INSPECTION COMMENTS"
                oTable.Range.Font.Name = "Arial"
                oTable.Range.Font.Size = 8
                oTable.Rows.Item(1).Range.Font.Bold = True
                oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                With oTable.Rows.Item(1).Shading
                    .BackgroundPatternColor = Word.WdColor.wdColorBlack
                End With

                ' content
                With oTable.Rows.Item(2).Shading
                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                End With
                oTable.Rows.Item(2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                If oInsp.CheckListMaster.InspectionComments.InsComments <> String.Empty Then
                    oTable.Rows.Item(2).Range.Text = oInsp.CheckListMaster.InspectionComments.InsComments
                    oTable.Rows.Item(2).Range.Font.Bold = False
                End If


                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                oPara.Range.Text = "<Space>"
                oPara.Format.SpaceAfter = 1
                oPara.Range.InsertParagraphAfter()

                ' Discrep
                dt = oInsp.CheckListMaster.DiscrepTable
                dv = dt.DefaultView
                dv.Sort = "CL_POSITION"

                If dt.Rows.Count <= 0 Then
                    dr = dt.NewRow
                    dr("CL_POSITION") = 1
                    dr("Line#") = String.Empty
                    dr("Description") = ""
                    dt.Rows.Add(dr)
                End If


                ''' Inspection CCAT


                ' content
                Dim g As Integer = 0

                Dim list As New Collections.ArrayList

                While g <= oInsp.CheckListMaster.InspectionInfo.CitationsCollection.Count - 1

                    Dim key As String = oInsp.CheckListMaster.InspectionInfo.CitationsCollection.GetKeys(g)

                    If oInsp.CheckListMaster.InspectionInfo.CitationsCollection.Item(key).CCAT <> String.Empty Then

                        Dim CCAT As String = oInsp.CheckListMaster.InspectionInfo.CitationsCollection(key).CCAT


                        dv.RowFilter = String.Format("QUESTION_ID = {0}", oInsp.CheckListMaster.InspectionInfo.CitationsCollection(key).QuestionID)

                        If dv.Count > 0 AndAlso Not dv.Item(0)("QUESTION_ID").Equals(DBNull.Value) Then
                            list.Add(String.Format("Citation {0}:   {1}", dv.Item(0)("Line#"), CCAT))
                        End If

                    End If

                    g = g + 1
                End While

                nRows = 0

                If oInsp.CheckListMaster.InspectionInfo.CitationsCollection.Count > 1 Then
                    nRows = Math.Max(38, list.Count + 2)
                Else
                    nRows = 3
                End If




                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, nRows, 1)
                oTable.Range.ParagraphFormat.KeepWithNext = True

                oTable.Range.Font.Name = "Arial"
                oTable.Range.Font.Size = 8
                oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                oTable.Columns.Item(1).Width = wordapp.InchesToPoints(7.5)
                With oTable.Rows.Shading
                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                End With

                ' Heading
                oTable.Rows.Item(1).Range.Text = "9             INSPECTION CCAT"
                oTable.Range.Font.Name = "Arial"
                oTable.Range.Font.Size = 8
                oTable.Rows.Item(1).Range.Font.Bold = True
                oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                With oTable.Rows.Item(1).Shading
                    .BackgroundPatternColor = Word.WdColor.wdColorBlack
                End With


                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                oPara.Range.Text = "<Space>"
                oPara.Format.SpaceAfter = 1
                oPara.Range.InsertParagraphAfter()

                g = 0

                For Each item As String In list

                    With oTable.Rows.Item(2 + g).Shading
                        .BackgroundPatternColor = Word.WdColor.wdColorWhite
                    End With

                    oTable.Rows.Item(2 + g).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    oTable.Rows.Item(2 + g).Range.Text = item
                    oTable.Rows.Item(2 + g).Range.Font.Bold = False

                    g = g + 1

                Next

                oTable.Sort(FieldNumber:=1)


                ' page break
                wordapp.Selection.GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:=.Bookmarks.Item("\endofdoc").Name)
                wordapp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)

                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, IIf(dt.Rows.Count + 2 > 33, dt.Rows.Count + 2, 33), 2)
                oTable.Range.ParagraphFormat.KeepWithNext = True

                oTable.Range.Font.Name = "Arial"
                oTable.Range.Font.Size = 8
                oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                oTable.Columns.Item(2).Width = wordapp.InchesToPoints(6.8)

                oTable.Cell(1, 1).Range.Text = "10"
                oTable.Cell(1, 2).Range.Text = "INSPECTION DISCREPANCIES"
                oTable.Cell(2, 1).Range.Text = "Line#"
                oTable.Cell(2, 2).Range.Text = "Description"

                With oTable.Rows.Item(1).Shading
                    .BackgroundPatternColor = Word.WdColor.wdColorBlack
                End With
                With oTable.Rows.Item(2).Shading
                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                    .Texture = Word.WdTextureIndex.wdTexture30Percent
                End With

                Dim k = 0


                ' Discrep part 2
                dt = oInsp.CheckListMaster.DiscrepTable
                dv = dt.DefaultView
                dv.Sort = "CL_POSITION"

                If dt.Rows.Count <= 0 Then
                    dr = dt.NewRow
                    dr("CL_POSITION") = 1
                    dr("Line#") = String.Empty
                    dr("Description") = ""
                    dt.Rows.Add(dr)
                End If

                For i = 0 To dt.Rows.Count - 1


                    oTable.Rows.Item(i + 3).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    With oTable.Rows.Item(i + 3).Shading
                        .BackgroundPatternColor = Word.WdColor.wdColorWhite

                        oTable.Cell(i + 3, 1).Range.Text = dv.Item(i)("Line#")
                        oTable.Cell(i + 3, 2).Range.Text = dv.Item(i)("Description")


                    End With
                Next





                ' page break
                wordapp.Selection.GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:=.Bookmarks.Item("\endofdoc").Name)
                wordapp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)


                ' Monitor Wells
                ds = oInsp.CheckListMaster.MWellTable
                dt = ds.Tables("TankPipeMW")
                If dt.Rows.Count > 0 Then
                    dv = dt.DefaultView
                    dv.Sort = "CL_POSITION"

                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, 1, 4)
                    oTable.Range.ParagraphFormat.KeepWithNext = True

                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                    oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                    oTable.Columns.Item(2).Width = wordapp.InchesToPoints(5.8)
                    oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                    'oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)

                    ' Header Line
                    oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    With oTable.Rows.Item(1).Shading
                        oTable.Cell(1, 1).Range.Text = dv.Item(0)("Line#")
                        oTable.Cell(1, 2).Range.Text = dv.Item(0)("Question")
                        .BackgroundPatternColor = Word.WdColor.wdColorBlack
                        oTable.Rows.Item(1).Range.Font.Bold = True
                    End With

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()

                    dtSub = ds.Tables("Well")
                    dvSub = dtSub.DefaultView
                    dvSub.Sort = "Well#, LINE_NUMBER"

                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, IIf(dtSub.Rows.Count <= 35, 35, dtSub.Rows.Count), 10)
                    oTable.Range.ParagraphFormat.KeepWithNext = True

                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                    oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                    oTable.Columns.Item(2).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(6).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(7).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(8).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(9).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(10).Width = wordapp.InchesToPoints(2.8)

                    oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    oTable.Cell(1, 1).Range.Text = ""
                    oTable.Cell(1, 2).Range.Text = "Well#"
                    oTable.Cell(1, 2).Range.Font.Size = 7
                    oTable.Cell(1, 3).Range.Text = "Well" + vbCrLf + "Depth"
                    oTable.Cell(1, 3).Range.Font.Size = 7
                    oTable.Cell(1, 4).Range.Text = "Depth" + vbCrLf + "to" + vbCrLf + "Water"
                    oTable.Cell(1, 4).Range.Font.Size = 7
                    oTable.Cell(1, 5).Range.Text = "Depth" + vbCrLf + "to" + vbCrLf + "Slots"
                    oTable.Cell(1, 5).Range.Font.Size = 7
                    oTable.Cell(1, 6).Range.Text = "Surface" + vbCrLf + "Sealed" + vbCrLf + "Yes"
                    oTable.Cell(1, 6).Range.Font.Size = 7
                    oTable.Cell(1, 7).Range.Text = "Surface" + vbCrLf + "Sealed" + vbCrLf + "No"
                    oTable.Cell(1, 7).Range.Font.Size = 7
                    oTable.Cell(1, 8).Range.Text = "Well" + vbCrLf + "Caps" + vbCrLf + "Yes"
                    oTable.Cell(1, 8).Range.Font.Size = 7
                    oTable.Cell(1, 9).Range.Text = "Well" + vbCrLf + "Caps" + vbCrLf + "No"
                    oTable.Cell(1, 9).Range.Font.Size = 7
                    oTable.Cell(1, 10).Range.Text = "Inspector's Observations"
                    oTable.Cell(1, 10).Range.Font.Size = 7
                    With oTable.Rows.Item(1).Shading
                        .BackgroundPatternColor = Word.WdColor.wdColorWhite
                        .Texture = Word.WdTextureIndex.wdTexture30Percent
                    End With
                    For i = 0 To dtSub.Rows.Count - 1
                        oTable.Rows.Item(i + 2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(i + 2).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            oTable.Cell(i + 2, 1).Range.Text = dvSub.Item(i)("Line#")
                            oTable.Cell(i + 2, 2).Range.Text = IIf(dvSub.Item(i)("Well#") = 0, String.Empty, dvSub.Item(i)("Well#"))
                            oTable.Cell(i + 2, 3).Range.Text = dvSub.Item(i)("Well Depth")
                            oTable.Cell(i + 2, 4).Range.Text = dvSub.Item(i)("Depth to" + vbCrLf + "Water")
                            oTable.Cell(i + 2, 5).Range.Text = dvSub.Item(i)("Depth to" + vbCrLf + "Slots")
                            oTable.Cell(i + 2, 6).Range.Text = IIf(dvSub.Item(i)("Surface Sealed" + vbCrLf + "Yes"), "X", "")
                            oTable.Cell(i + 2, 7).Range.Text = IIf(dvSub.Item(i)("Surface Sealed" + vbCrLf + "No"), "X", "")
                            oTable.Cell(i + 2, 8).Range.Text = IIf(dvSub.Item(i)("Well Caps" + vbCrLf + "Yes"), "X", "")
                            oTable.Cell(i + 2, 9).Range.Text = IIf(dvSub.Item(i)("Well Caps" + vbCrLf + "No"), "X", "")
                            oTable.Cell(i + 2, 10).Range.Text = dvSub.Item(i)("Inspector's Observations")
                        End With
                    Next

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()
                End If




                If hasDesignatedOwner Then
                    ' page break
                    wordapp.Selection.GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:=.Bookmarks.Item("\endofdoc").Name)
                    wordapp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)

                    ' Operator Questions
                    dt = oInsp.CheckListMaster.OtherQusetionsTable
                    If dt.Rows.Count > 0 Then
                        dv = dt.DefaultView
                        dv.Sort = "CL_POSITION"

                        oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dv.Count + 1, 4)
                        oTable.Range.ParagraphFormat.KeepWithNext = True

                        oTable.Range.Font.Name = "Arial"
                        oTable.Range.Font.Size = 8
                        oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                        oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                        oTable.Columns.Item(2).Width = wordapp.InchesToPoints(5.8)
                        oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                        oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)




                        For i = 0 To dv.Count - 1
                            oTable.Rows.Item(i + 1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oTable.Rows.Item(i + 1).Shading
                                oTable.Cell(i + 1, 1).Range.Text = dv.Item(i)("Line#")
                                oTable.Cell(i + 1, 2).Range.Text = dv.Item(i)("Question")
                                If dv.Item(i)("HEADER") Then
                                    If dv.Item(i)("Line#").ToString.Length = 1 Then
                                        .BackgroundPatternColor = Word.WdColor.wdColorBlack
                                    Else
                                        .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                        .Texture = Word.WdTextureIndex.wdTextureSolid
                                    End If
                                    oTable.Rows.Item(i + 1).Range.Font.Bold = True
                                    oTable.Cell(i + 1, 3).Range.Text = "True"
                                    oTable.Cell(i + 1, 4).Range.Text = "False"
                                Else
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                    oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                    oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                End If
                            End With
                        Next


                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Font.Name = "Arial"
                        oPara.Range.Font.Size = 8
                        oPara.Range.Text = vbCrLf + vbCrLf + vbCrLf + vbCrLf + _
                                           "BY YOUR SIGNATURE, IT IS ACKNOWLEDGED THAT YOU UNDERSTAND AND" + vbCrLf + _
                                           "AGREE WITH THE REQUIREMENTS FOR COMPLIANCE WITH THIS FACILITY." + vbCrLf + vbCrLf + _
                                           "OWNER/OWNER'S REPRESENTATIVE PRINTED NAME ________________________________________" + vbCrLf + vbCrLf + _
                                           "OWNER/OWNER'S REPRESENTATIVE SIGNATURE ________________________________________" + vbCrLf + vbCrLf + _
                                           "DATE ____________________"
                        oPara.Range.InsertParagraphAfter()


                    End If


                End If

                ' OPER
                dt = oInsp.CheckListMaster.OtherQusetionsTable
                dv = dt.DefaultView
                dv.Sort = "CL_POSITION"

                If dt.Rows.Count > 0 Then
                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dt.Rows.Count, 5)
                    oTable.Range.ParagraphFormat.KeepWithNext = True

                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)
                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                    oTable.Columns.Item(1).Width = wordapp.InchesToPoints(0.7)
                    oTable.Columns.Item(2).Width = wordapp.InchesToPoints(5.8)
                    oTable.Columns.Item(3).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(4).Width = wordapp.InchesToPoints(0.5)
                    oTable.Columns.Item(5).Width = wordapp.InchesToPoints(0.5)

                    For i = 0 To dt.Rows.Count - 1
                        oTable.Rows.Item(i + 1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(i + 1).Shading
                            oTable.Cell(i + 1, 1).Range.Text = dv.Item(i)("Line#")
                            oTable.Cell(i + 1, 2).Range.Text = dv.Item(i)("Question")
                            If dv.Item(i)("HEADER") Then
                                If dv.Item(i)("Line#").ToString.Length = 1 Then
                                    .BackgroundPatternColor = Word.WdColor.wdColorBlack
                                Else
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                    .Texture = Word.WdTextureIndex.wdTexture30Percent
                                End If
                                oTable.Rows.Item(i + 1).Range.Font.Bold = True
                                oTable.Cell(i + 1, 3).Range.Text = "Yes"
                                oTable.Cell(i + 1, 4).Range.Text = "No"
                                oTable.Cell(i + 1, 5).Range.Text = "N/A"
                            Else
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                oTable.Cell(i + 1, 5).Range.Text = IIf(dv.Item(i)("N/A"), "X", "")
                            End If
                        End With
                    Next

                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceAfter = 1
                    oPara.Range.InsertParagraphAfter()
                End If


                '' Sketch

                Dim osketch As New BusinessLogic.pInspectionSketch
                osketch.Retrieve(oInsp.InspectionInfo)

                If osketch.ID > 0 AndAlso osketch.SketchFileName.Length > 0 AndAlso File.Exists(String.Format("{0}\{1}", sketchpath, osketch.SketchFileName).ToUpper.Replace(".PPT", "\slide1.BMP").Replace("Sketch_", String.Empty)) Then

                    ' page break
                    wordapp.Selection.GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:=.Bookmarks.Item("\endofdoc").Name)
                    wordapp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)


                    ' delete spacing between tables
                    oPara.Range.Text = "<Space>"
                    For Each para As Word.Paragraph In .Content.Paragraphs
                        If para.Range.Text = oPara.Range.Text Then
                            para.Range.Delete()
                        End If
                    Next


                    ' Heading
                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, 2, 1)
                    oTable.Range.ParagraphFormat.KeepWithNext = True

                    oTable.Rows.Item(1).Range.Text = "12             SKETCH OF FACILITY"
                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Rows.Height = wordapp.InchesToPoints(0.25)

                    oTable.Rows.Item(1).Range.Font.Bold = True
                    oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    With oTable.Rows.Item(1).Shading
                        .BackgroundPatternColor = Word.WdColor.wdColorBlack
                    End With


                    'Insert an image at the end of the document.
                    oTable.Rows.Item(2).Height = 600
                    oTable.Rows.Item(2).Range.InlineShapes.AddPicture(String.Format("{0}\{1}", sketchpath, osketch.SketchFileName).ToUpper.Replace(".PPT", "\slide1.BMP").Replace("Sketch_", String.Empty))

                    With oTable.Rows.Item(2).Range.InlineShapes.Item(oTable.Rows.Item(2).Range.InlineShapes.Count)


                    End With
                End If

                osketch = Nothing


                '' Maps
                If Not imgs Is Nothing AndAlso imgs.Count > 0 Then




                    ' page break
                    wordapp.Selection.GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:=.Bookmarks.Item("\endofdoc").Name)
                    wordapp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)


                    ' delete spacing between tables
                    oPara.Range.Text = "<Space>"
                    For Each para As Word.Paragraph In .Content.Paragraphs
                        If para.Range.Text = oPara.Range.Text Then
                            para.Range.Delete()
                        End If
                    Next


                    ' Heading

                    Dim shrink As Integer = 0
                    Dim rows, cols As Integer

                    shrink = CInt(Math.Sqrt(imgs.Count)) + IIf(Math.Sqrt(imgs.Count) <> Int(Math.Sqrt(imgs.Count)), 1, 0)

                    rows = Int((800.0 / shrink) * 0.95)
                    cols = Int((600.0 / shrink) * 0.95)

                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, 1 + shrink, shrink)

                    oTable.Range.ParagraphFormat.KeepWithNext = True


                    oTable.Rows.Item(1).Range.Text = "13           FACILITY PICS"


                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8

                    oTable.Rows.Item(1).Range.Font.Bold = True
                    oTable.Rows.Item(1).Cells.Merge()
                    oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                    With oTable.Rows.Item(1).Shading
                        .BackgroundPatternColor = Word.WdColor.wdColorBlack
                    End With


                    Dim x, y As Integer
                    x = 1
                    y = 2
                    For Each item As Bitmap In imgs


                        'Insert an image at the end of the document.
                        If x = 1 Then
                            oTable.Rows.Item(y).Height = rows
                        End If
                        oTable.Rows.Item(y).Cells.Item(x).Width = cols

                        System.Windows.Forms.Clipboard.SetDataObject(item)
                        oTable.Rows.Item(y).Cells.Item(x).Range.Paste()
                        oTable.Rows.Item(y).Cells.Item(x).Range.InlineShapes.Item(1).Height = oTable.Rows.Item(y).Cells.Item(x).Height
                        oTable.Rows.Item(y).Cells.Item(x).Range.InlineShapes.Item(1).Width = oTable.Rows.Item(y).Cells.Item(x).Width

                        x += 1
                        If x >= (shrink + 1) Then
                            x = 1
                            y += 1
                        End If
                    Next

                End If

            End With

        End Sub


        Private Sub InspectionAddTanks(ByVal dt As DataTable, ByRef WordApp As Word.Application, ByRef DestDoc As Word.Document, ByRef start As Boolean)

            Dim i, colCount1, colCount2, colCount3, colIndex As Integer
            Dim inches As Double = 7.5
            Dim oTable As Word.Table
            Dim oPara As Word.Paragraph
            Try
                With DestDoc



                    inches = 7.5


                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    'oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceBefore = 0
                    oPara.Format.SpaceAfter = 0
                    oPara.KeepWithNext = True
                    oPara.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle


                    oPara.Range.InsertParagraphAfter()

                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dt.Columns.Count + 1, dt.Rows.Count + 1)


                    oTable.Rows.WrapAroundText = False

                    oTable.Range.ParagraphFormat.KeepWithNext = True


                    oTable.AllowPageBreaks = False
                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Range.Font.Bold = True
                    oTable.Range.ParagraphFormat.SpaceAfter = 1
                    oTable.Range.ParagraphFormat.LineSpacing = 9
                    oTable.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle

                    oTable.Rows.Height = WordApp.InchesToPoints(0.18)

                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True


                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).LineStyle = Word.WdLineStyle.wdLineStyleSingle

                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).LineWidth = Word.WdLineWidth.wdLineWidth050pt
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).LineWidth = Word.WdLineWidth.wdLineWidth050pt
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).LineWidth = Word.WdLineWidth.wdLineWidth050pt
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).LineWidth = Word.WdLineWidth.wdLineWidth050pt

                    With oTable.Rows.Item(1)

                        .Cells.Merge()
                        .Range.Font.Size = 12

                        .Range.Text = "UST TANKS"
                        .Alignment = Word.WdRowAlignment.wdAlignRowLeft
                        .Shading.BackgroundPatternColor = Word.WdColor.wdColorBlack
                        .Range.Font.Color = Word.WdColor.wdColorWhite
                        .Shading.Texture = Word.WdTextureIndex.wdTextureSolid
                    End With

                    For Each col As DataColumn In dt.Columns



                        With oTable.Rows.Item(col.Ordinal + 2).Cells.Item(1)
                            .VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop


                            .Range.Font.Size = 10

                            With .Shading
                                '.BackgroundPatternColor = Word.WdColor.wdColorBlack
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                .Texture = Word.WdTextureIndex.wdTexture35Percent
                            End With
                        End With



                        oTable.Rows.Item(col.Ordinal + 2).Cells.Item(1).Range.Text = col.ColumnName
                        oTable.Rows.Item(col.Ordinal + 2).Alignment = Word.WdRowAlignment.wdAlignRowLeft

                        For g As Integer = 1 To dt.Rows.Count

                            If col.Ordinal = 0 Then
                                With oTable.Rows.Item(col.Ordinal + 2).Cells.Item(g + 1)
                                    .VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop

                                    .Range.Font.Size = 10

                                    With .Shading
                                        '.BackgroundPatternColor = Word.WdColor.wdColorBlack
                                        .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                        .Texture = Word.WdTextureIndex.wdTexture35Percent
                                    End With
                                End With
                            End If


                            If IIf(dt.Rows(g - 1).Item(col.Ordinal) Is DBNull.Value, "", dt.Rows(g - 1).Item(col.Ordinal).ToString) = "<CAPNEEDED>" Then
                                oTable.Rows.Item(col.Ordinal + 2).Cells.Item(g + 1).Range.Text = String.Empty
                                oTable.Rows.Item(col.Ordinal + 2).Cells.Item(g + 1).Shading.BackgroundPatternColor = Word.WdColor.wdColorYellow
                            Else
                                oTable.Rows.Item(col.Ordinal + 2).Cells.Item(g + 1).Range.Text = IIf(dt.Rows(g - 1).Item(col.Ordinal) Is DBNull.Value, "", dt.Rows(g - 1).Item(col.Ordinal).ToString)
                            End If

                        Next

                    Next

                End With


            Catch ex As Exception
                Throw ex
            End Try



        End Sub

        Private Sub InspectionAddTanksOld(ByVal ds As DataSet, ByRef WordApp As Word.Application, ByRef DestDoc As Word.Document, ByRef start As Boolean)
            Dim dt As DataTable
            Dim i, colCount1, colCount2, colCount3, colIndex As Integer
            Dim inches As Double = 7.5
            Dim oTable As Word.Table
            Dim oPara As Word.Paragraph
            Try
                With DestDoc
                    ' Add Tank Table
                    For i = 0 To ds.Tables.Count - 1

                        colCount1 = 0
                        colCount2 = 0
                        colCount3 = 0


                        Dim backOne As Integer = 0

                        inches = 7.5

                        dt = ds.Tables(i)
                        ' First Row
                        ' to determine how many columns to display in first row of a given tank
                        colCount1 = 10

                        Dim a, b, c As Integer
                        If i > 0 Then
                            a = 1
                            b = 2
                            c = 2
                        Else
                            a = 2
                            b = 3
                            c = 3

                        End If

                        If dt.Columns.Count > 11 Then

                            If dt.Columns.Count <= 15 Then
                                c = 2
                            End If
                        Else
                            c = 1
                        End If

                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        'oPara.Range.Text = "<Space>"
                        oPara.Format.SpaceBefore = 0
                        oPara.Format.SpaceAfter = 0
                        oPara.KeepWithNext = True
                        oPara.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast


                        oPara.Range.InsertParagraphAfter()

                        If oPara Is Nothing Then
                            oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, c + 3, colCount1)
                        Else
                            oTable = .Tables.Add(oPara.Range, c + 3, colCount1)


                        End If

                        oTable.Rows.WrapAroundText = True
                        oTable.Range.ParagraphFormat.KeepWithNext = True


                        oTable.AllowPageBreaks = False
                        oTable.Range.Font.Name = "Arial"
                        oTable.Range.Font.Size = 8
                        oTable.Range.Font.Bold = True
                        oTable.Range.ParagraphFormat.SpaceAfter = 1
                        oTable.Range.ParagraphFormat.LineSpacing = 9
                        oTable.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly

                        oTable.Rows.Height = WordApp.InchesToPoints(0.15)

                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True


                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).LineStyle = Word.WdLineStyle.wdLineStyleSingle

                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).LineWidth = Word.WdLineWidth.wdLineWidth050pt
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).LineWidth = Word.WdLineWidth.wdLineWidth050pt
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).LineWidth = Word.WdLineWidth.wdLineWidth050pt
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).LineWidth = Word.WdLineWidth.wdLineWidth050pt

                        oTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.08, inches))
                        oTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.09, inches))

                        oTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                        oTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                        oTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                        oTable.Columns.Item(6).Width = WordApp.InchesToPoints(GetApproxInches(0.07, inches))
                        oTable.Columns.Item(7).Width = WordApp.InchesToPoints(GetApproxInches(0.09, inches))
                        oTable.Columns.Item(8).Width = WordApp.InchesToPoints(GetApproxInches(0.09, inches))
                        oTable.Columns.Item(9).Width = WordApp.InchesToPoints(GetApproxInches(0.14, inches))
                        oTable.Columns.Item(10).Width = WordApp.InchesToPoints(GetApproxInches(0.14, inches))

                        ' row 1
                        If i = 0 Then
                            oTable.Rows.Item(1).Cells.Merge()
                            oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            oTable.Rows.Item(1).Range.Text = "TANKS"
                            oTable.Rows.Item(1).Range.Font.Size = 10
                            With oTable.Rows.Item(1).Shading
                                '.BackgroundPatternColor = Word.WdColor.wdColorBlack
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                .Texture = Word.WdTextureIndex.wdTexture35Percent
                            End With
                            start = False
                        End If



                        ' row 2
                        oTable.Rows.Item(a).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(a).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture25Percent
                        End With


                        ' row 3
                        oTable.Rows.Item(b).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(b).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                        End With

                        oTable.Rows.Item(a).Range.Font.Size = 9
                        oTable.Rows.Item(a).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                        oTable.Rows.Item(b).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                        ' Fill Data
                        If (colCount1 - 1) > dt.Columns.Count - 1 Then
                            colCount1 -= ((colCount1 - 1) - (dt.Columns.Count - 1))
                        End If

                        For colIndex = 1 To colCount1
                            ' row 2 - header
                            oTable.Rows.Item(a).Cells().Item(colIndex).Range.Text = dt.Columns(colIndex - 1).ColumnName.Replace("Contents", "Content")

                            ' row 3 - data
                            oTable.Rows.Item(b).Cells().Item(colIndex).Range.Text = dt.Rows(0)(colIndex - 1)

                        Next


                        oTable.Rows.Item(a).LeftIndent = 20.0
                        oTable.Rows.Item(b).LeftIndent = 20.0
                        If oTable.Rows.Count >= (b + 1) Then
                            oTable.Rows.Item(b + 1).LeftIndent = 20.0
                            oTable.Rows.Item(b + 1).Cells.Merge()
                        End If

                        If oTable.Rows.Count >= (b + 2) Then

                            oTable.Rows.Item(b + 2).LeftIndent = 20.0
                            oTable.Rows.Item(b + 2).Cells.Merge()

                        End If


                        If oTable.Rows.Count >= (b + 3) Then
                            oTable.Rows.Item(b + 3).LeftIndent = 20.0
                            oTable.Rows.Item(b + 3).Cells.Merge()
                            oTable.Rows.Item(b + 3).Shading.BackgroundPatternColor = Word.WdColor.wdColorWhite
                            oTable.Rows.Item(b + 3).Shading.ForegroundPatternColor = Word.WdColor.wdColorWhite

                            oTable.Rows.Item(b + 3).Borders.Item(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleNone
                            oTable.Rows.Item(b + 3).Borders.Item(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleNone
                            oTable.Rows.Item(b + 3).Borders.Item(Word.WdBorderType.wdBorderLeft).LineStyle = Word.WdLineStyle.wdLineStyleNone
                            oTable.Rows.Item(b + 3).Borders.Item(Word.WdBorderType.wdBorderRight).LineStyle = Word.WdLineStyle.wdLineStyleNone

                        End If




                        ' Second row
                        ' to determine how many cells to display in second row of a given tank
                        colCount2 = dt.Columns.Count - 11

                        If colCount2 > 4 Then

                            colCount3 = colCount2 - 4
                            colCount2 = 4

                        End If


                        If colCount2 > 1 Then


                            If (colCount1 + colCount2 - 1) > (dt.Columns.Count - 1) Then
                                colCount2 -= ((colCount1 + colCount2 - 1) - (dt.Columns.Count - 1))
                            End If

                            Dim oSubTable As Word.Table

                            oTable.Rows.Item(b + 1).Cells.Item(1).Select()


                            oSubTable = oTable.Rows.Item(b + 1).Cells.Item(1).Tables.Add(WordApp.Selection.Range, 2, colCount2 - 1)
                            inches = WordApp.PointsToInches(oTable.Rows.Item(b + 1).Cells.Item(1).Width)

                            If oSubTable.Rows.Count <= 1 Then
                                oSubTable.Rows.Add()
                            End If

                            oSubTable.LeftPadding = 0
                            oSubTable.RightPadding = 0
                            oSubTable.Rows.WrapAroundText = True


                            oSubTable.AllowPageBreaks = False

                            oSubTable.Range.Font.Name = "Arial"
                            oSubTable.Range.Font.Size = 8
                            oSubTable.Range.Font.Bold = True
                            oSubTable.Rows.Height = WordApp.InchesToPoints(0.15)
                            oSubTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                            oSubTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                            oSubTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                            oSubTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True



                            '   oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)

                            Try
                                If colCount2 = 2 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(inches)
                                ElseIf colCount2 = 3 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.5, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.5, inches))
                                ElseIf colCount2 = 4 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.43, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.43, inches))
                                    oSubTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.14, inches))
                                ElseIf colCount2 = 5 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.375, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.375, inches))
                                    oSubTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.125, inches))
                                    oSubTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.125, inches))
                                ElseIf colCount2 = 6 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.25, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.25, inches))
                                    oSubTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                                    oSubTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                                    oSubTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.3, inches))
                                ElseIf colCount2 = 7 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.24, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.24, inches))
                                    oSubTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.13, inches))
                                    oSubTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.13, inches))
                                    oSubTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.13, inches))
                                    oSubTable.Columns.Item(6).Width = WordApp.InchesToPoints(GetApproxInches(0.13, inches))
                                Else
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.2, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.2, inches))
                                    oSubTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.2, inches))
                                    oSubTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                                    oSubTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                                    oSubTable.Columns.Item(6).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                                    oSubTable.Columns.Item(7).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                                End If
                            Catch
                            End Try


                            ' row 2
                            oSubTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oSubTable.Rows.Item(1).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                .Texture = Word.WdTextureIndex.wdTexture25Percent
                            End With

                            ' row 3
                            oSubTable.Rows.Item(2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oSubTable.Rows.Item(2).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            End With

                            ' Fill Data

                            oSubTable.Rows.Item(1).Range.Font.Size = 9
                            oSubTable.Rows.Item(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            oSubTable.Rows.Item(2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter


                            For colIndex = 1 To colCount2
                                Try
                                    ' row 2 - header
                                    oSubTable.Rows.Item(1).Cells().Item(colIndex).Range.Text = dt.Columns(colCount1 + colIndex - 1).ColumnName

                                    ' row 3 - data
                                    oSubTable.Rows.Item(2).Cells().Item(colIndex).Range.Text = dt.Rows(0)(colCount1 + colIndex - 1)
                                Catch
                                End Try


                            Next


                        End If



                        If colCount3 > 1 Then


                            If (colCount1 + colCount2 + colCount3 - 1) > (dt.Columns.Count - 1) Then
                                colCount3 -= ((colCount1 + colCount2 + colCount3 - 1) - (dt.Columns.Count - 1))
                            End If

                            Dim oSubTable As Word.Table

                            oTable.Rows.Item(b + 2).Cells.Item(1).Select()

                            oSubTable = oTable.Rows.Item(b + 2).Cells.Item(1).Tables.Add(WordApp.Selection.Range, 2, colCount3)
                            inches = WordApp.PointsToInches(oTable.Rows.Item(b + 2).Cells.Item(1).Width)

                            If oSubTable.Rows.Count <= 1 Then
                                oSubTable.Rows.Add()
                            End If

                            oSubTable.LeftPadding = 0
                            oSubTable.RightPadding = 0
                            oSubTable.Rows.WrapAroundText = True



                            oSubTable.AllowPageBreaks = False

                            oSubTable.Range.Font.Name = "Arial"
                            oSubTable.Range.Font.Size = 8
                            oSubTable.Range.Font.Bold = True
                            oSubTable.Rows.Height = WordApp.InchesToPoints(0.15)
                            oSubTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                            oSubTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                            oSubTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                            oSubTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True


                            Try

                                If colCount3 = 2 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(inches)
                                ElseIf colCount3 = 3 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.5, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.5, inches))
                                ElseIf colCount3 = 4 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.43, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.43, inches))
                                    oSubTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.14, inches))
                                ElseIf colCount3 = 5 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.375, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.375, inches))
                                    oSubTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.125, inches))
                                    oSubTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.125, inches))
                                ElseIf colCount3 = 6 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.25, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.25, inches))
                                    oSubTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                                    oSubTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                                    oSubTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.3, inches))


                                    ' row 2
                                    oSubTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                    With oSubTable.Rows.Item(1).Shading
                                        .BackgroundPatternColor = Word.WdColor.wdColorGray50
                                        .Texture = Word.WdTextureIndex.wdTexture25Percent
                                    End With

                                    ' row 3
                                    oSubTable.Rows.Item(2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                    With oSubTable.Rows.Item(2).Shading
                                        .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                    End With
                                End If
                            Catch
                            End Try



                            ' Fill Data

                            oSubTable.Rows.Item(1).Range.Font.Size = 9
                            oSubTable.Rows.Item(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            oSubTable.Rows.Item(2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            oSubTable.Rows.Item(1).Shading.BackgroundPatternColor = Word.WdColor.wdColorGray20



                            For colIndex = 0 To (colCount3)
                                Try
                                    ' row 2 - header
                                    oSubTable.Rows.Item(1).Cells().Item(colIndex + 1).Range.Text = dt.Columns(colCount1 + (colCount2) + colIndex - 1).ColumnName

                                    ' row 3 - data
                                    oSubTable.Rows.Item(2).Cells().Item(colIndex + 1).Range.Text = dt.Rows(0)(colCount1 + (colCount2) + colIndex - 1)
                                Catch
                                End Try


                            Next

                        End If

                    Next
                End With
            Catch ex As Exception
                Throw ex
            End Try
        End Sub


        Private Sub InspectionAddPipes(ByVal dt As DataTable, ByRef WordApp As Word.Application, ByRef DestDoc As Word.Document, ByRef start As Boolean)

            Dim i, colCount1, colCount2, colCount3, colIndex As Integer
            Dim inches As Double = 7.5
            Dim oTable As Word.Table
            Dim oPara As Word.Paragraph
            Try
                With DestDoc


                    inches = 7.5



                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    'oPara.Range.Text = "<Space>"
                    oPara.Format.SpaceBefore = 0
                    oPara.Format.SpaceAfter = 0
                    oPara.KeepWithNext = True
                    oPara.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle


                    oPara.Range.InsertParagraphAfter()

                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dt.Columns.Count + 1, dt.Rows.Count + 1)

                    oTable.Rows.WrapAroundText = True
                    oTable.Range.ParagraphFormat.KeepWithNext = True


                    oTable.AllowPageBreaks = False
                    oTable.Range.Font.Name = "Arial"
                    oTable.Range.Font.Size = 8
                    oTable.Range.Font.Bold = True
                    oTable.Range.ParagraphFormat.SpaceAfter = 1
                    oTable.Range.ParagraphFormat.LineSpacing = 9
                    oTable.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle

                    oTable.Rows.Height = WordApp.InchesToPoints(0.18)

                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True


                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).LineStyle = Word.WdLineStyle.wdLineStyleSingle

                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).LineWidth = Word.WdLineWidth.wdLineWidth050pt
                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).LineWidth = Word.WdLineWidth.wdLineWidth050pt
                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).LineWidth = Word.WdLineWidth.wdLineWidth050pt
                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).LineWidth = Word.WdLineWidth.wdLineWidth050pt

                    With oTable.Rows.Item(1)

                        .Cells.Merge()
                        .Range.Font.Size = 12

                        .Range.Text = "UST PIPES"
                        .Alignment = Word.WdRowAlignment.wdAlignRowLeft

                        .Shading.BackgroundPatternColor = Word.WdColor.wdColorBlack
                        .Range.Font.Color = Word.WdColor.wdColorWhite
                        .Shading.Texture = Word.WdTextureIndex.wdTextureSolid
                    End With


                    For Each col As DataColumn In dt.Columns
                        oTable.Rows.Item(col.Ordinal + 2).Cells.Item(1).Range.Text = col.ColumnName
                        oTable.Rows.Item(col.Ordinal + 2).Alignment = Word.WdRowAlignment.wdAlignRowLeft



                        With oTable.Rows.Item(col.Ordinal + 2).Cells.Item(1)
                            .VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop


                            .Range.Font.Size = 10

                            With .Shading
                                '.BackgroundPatternColor = Word.WdColor.wdColorBlack
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                .Texture = Word.WdTextureIndex.wdTexture35Percent
                            End With
                        End With





                        For g As Integer = 1 To dt.Rows.Count

                            If col.Ordinal = 0 Then
                                With oTable.Rows.Item(col.Ordinal + 2).Cells.Item(g + 1)
                                    .VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalTop
                                    .Range.Font.Size = 10

                                    With .Shading
                                        '.BackgroundPatternColor = Word.WdColor.wdColorBlack
                                        .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                        .Texture = Word.WdTextureIndex.wdTexture35Percent
                                    End With
                                End With
                            End If

                            If IIf(dt.Rows(g - 1).Item(col.Ordinal) Is DBNull.Value, "", dt.Rows(g - 1).Item(col.Ordinal).ToString) = "<CAPNEEDED>" Then
                                oTable.Rows.Item(col.Ordinal + 2).Cells.Item(g + 1).Range.Text = String.Empty
                                oTable.Rows.Item(col.Ordinal + 2).Cells.Item(g + 1).Shading.BackgroundPatternColor = Word.WdColor.wdColorYellow
                            Else

                                oTable.Rows.Item(col.Ordinal + 2).Cells.Item(g + 1).Range.Text = IIf(dt.Rows(g - 1).Item(col.Ordinal) Is DBNull.Value, "", dt.Rows(g - 1).Item(col.Ordinal).ToString)
                            End If

                        Next

                    Next

                End With


            Catch ex As Exception
                Throw ex
            End Try


        End Sub

        Private Sub InspectionAddPipesOld(ByVal ds As DataSet, ByRef WordApp As Word.Application, ByRef DestDoc As Word.Document, ByRef start As Boolean)
            Dim dt As DataTable
            Dim i, colCount1, colCount2, colIndex As Integer
            Dim inches As Double = 7.5
            Dim oTable As Word.Table
            Dim oPara As Word.Paragraph
            Try
                With DestDoc
                    ' Add Pipe Table
                    For i = 0 To ds.Tables.Count - 1

                        inches = 7.5

                        colCount1 = 0
                        colCount2 = 0

                        dt = ds.Tables(i)
                        ' First Row
                        ' to determine how many cells to display in first row of a given pipe
                        colCount1 = 8
                        If Not dt.Columns.Contains("Brand") Then
                            colCount1 -= 1
                        End If
                        If Not dt.Columns.Contains("CP Type") Then
                            colCount1 -= 1
                        End If

                        Dim a, b, c As Integer
                        If i > 0 Then
                            a = 1
                            b = 2
                            c = 2
                        Else
                            a = 2
                            b = 3
                            c = 3

                        End If

                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        '                          oPara.Range.Text = "<Space>"
                        oPara.Format.SpaceBefore = 0
                        oPara.Format.SpaceAfter = 0
                        oPara.KeepWithNext = True

                        oPara.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceAtLeast


                        oPara.Range.InsertParagraphAfter()


                        oTable = .Tables.Add(oPara.Range, c + 1, colCount1)
                        oTable.AllowPageBreaks = False
                        oTable.Range.ParagraphFormat.KeepWithNext = True


                        oTable.Rows.WrapAroundText = True
                        oTable.Range.Font.Name = "Arial"
                        oTable.Range.Font.Size = 8
                        oTable.Range.Font.Bold = True
                        oTable.Rows.Height = WordApp.InchesToPoints(0.15)
                        oTable.Range.ParagraphFormat.SpaceAfter = 1
                        oTable.Range.ParagraphFormat.LineSpacing = 9
                        oTable.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly


                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleSingle
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).LineStyle = Word.WdLineStyle.wdLineStyleSingle

                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).LineWidth = Word.WdLineWidth.wdLineWidth050pt
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).LineWidth = Word.WdLineWidth.wdLineWidth050pt
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).LineWidth = Word.WdLineWidth.wdLineWidth050pt
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).LineWidth = Word.WdLineWidth.wdLineWidth050pt

                        If (colCount1 - 1) > (dt.Columns.Count - 1) Then
                            colCount1 -= ((colCount1 - 1) - (dt.Columns.Count - 1))
                        End If

                        oTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                        oTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                        oTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.12, inches))

                        If colCount1 = 6 Then
                            oTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.2266, inches))
                            oTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.2266, inches))
                            oTable.Columns.Item(6).Width = WordApp.InchesToPoints(GetApproxInches(0.2267, inches))
                        ElseIf colCount1 = 7 Then
                            oTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.17, inches))
                            oTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.17, inches))
                            oTable.Columns.Item(6).Width = WordApp.InchesToPoints(GetApproxInches(0.17, inches))
                            oTable.Columns.Item(7).Width = WordApp.InchesToPoints(GetApproxInches(0.17, inches))
                        Else
                            oTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.136, inches))
                            oTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.136, inches))
                            oTable.Columns.Item(6).Width = WordApp.InchesToPoints(GetApproxInches(0.136, inches))
                            oTable.Columns.Item(7).Width = WordApp.InchesToPoints(GetApproxInches(0.136, inches))
                            oTable.Columns.Item(8).Width = WordApp.InchesToPoints(GetApproxInches(0.136, inches))
                        End If

                        ' row 1
                        If i = 0 Then
                            oTable.Rows.Item(1).Cells.Merge()
                            oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            oTable.Rows.Item(1).Range.Text = "PIPING"
                            oTable.Rows.Item(1).Range.Font.Size = 10

                            With oTable.Rows.Item(1).Shading
                                '.BackgroundPatternColor = Word.WdColor.wdColorBlack
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                .Texture = Word.WdTextureIndex.wdTexture35Percent
                            End With
                            start = False
                        End If


                        ' row 2
                        With oTable.Rows.Item(a).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture25Percent
                        End With

                        ' row 3
                        With oTable.Rows.Item(b).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                        End With

                        oTable.Rows.Item(a).LeftIndent = 20.0
                        oTable.Rows.Item(b + 1).LeftIndent = 20.0
                        oTable.Rows.Item(b + 1).Cells.Merge()
                        oTable.Rows.Item(b).LeftIndent = 20.0

                        ' Fill Data
                        oTable.Rows.Item(a).Range.Font.Size = 9
                        oTable.Rows.Item(a).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                        oTable.Rows.Item(b).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                        For colIndex = 1 To colCount1
                            ' row 2 - header
                            oTable.Rows.Item(a).Cells().Item(colIndex).Range.Text = dt.Columns(colIndex - 1).ColumnName

                            ' row 3 - data
                            oTable.Rows.Item(b).Cells().Item(colIndex).Range.Text = dt.Rows(0)(colIndex - 1)

                        Next


                        ' Second row
                        ' to determine how many cells to display in second row of a given pipe
                        colCount2 = 10
                        If Not dt.Columns.Contains("Primary LD") Then
                            colCount2 -= 1
                        End If
                        If Not dt.Columns.Contains("Secondary LD") Then
                            colCount2 -= 1
                        End If
                        If Not dt.Columns.Contains("ALLD Tested") Then
                            colCount2 -= 1
                        End If
                        If Not dt.Columns.Contains("PTT") Then
                            colCount2 -= 1
                        End If
                        If Not dt.Columns.Contains("CP Installed") Then
                            colCount2 -= 1
                        End If
                        If Not dt.Columns.Contains("CP Tested") Then
                            colCount2 -= 1
                        End If
                        If Not dt.Columns.Contains("Shear Tested") Then
                            colCount2 -= 1
                        End If
                        If Not dt.Columns.Contains("Secondary Inspected") Then
                            colCount2 -= 1
                        End If

                        If Not dt.Columns.Contains("Electronic Inspected") Then
                            colCount2 -= 1
                        End If



                        If colCount2 > 1 Then

                            If (colCount1 + colCount2 - 2) > (dt.Columns.Count - 1) Then
                                colCount2 -= ((colCount1 = colCount2 - 2) - (dt.Columns.Count - 1))
                            End If

                            Dim oSubTable As Word.Table

                            oTable.Rows.Item(b + 1).Cells.Item(1).Select()

                            oSubTable = oTable.Rows.Item(b + 1).Cells.Item(1).Tables.Add(WordApp.Selection.Range, 2, colCount2 - 1)
                            inches = WordApp.PointsToInches(oTable.Rows.Item(b + 1).Cells.Item(1).Width)

                            If oSubTable.Rows.Count <= 1 Then
                                oSubTable.Rows.Add()
                            End If

                            oSubTable.LeftPadding = 0
                            oSubTable.RightPadding = 0
                            oSubTable.Rows.WrapAroundText = True



                            oSubTable.AllowPageBreaks = False
                            oSubTable.Range.Font.Name = "Arial"
                            oSubTable.Range.Font.Size = 7
                            oSubTable.Range.Font.Bold = True
                            oSubTable.Rows.Height = WordApp.InchesToPoints(0.15)
                            oSubTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                            oSubTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                            oSubTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                            oSubTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                            '              oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)

                            Try

                                If colCount2 = 2 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(inches)
                                ElseIf colCount2 = 3 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.5, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.5, inches))
                                ElseIf colCount2 = 3 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.43, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.43, inches))
                                    oSubTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.14, inches))
                                ElseIf colCount2 = 5 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.38, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.38, inches))
                                    oSubTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.12, inches))
                                    oSubTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.12, inches))
                                ElseIf colCount2 = 6 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.35, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.35, inches))
                                    oSubTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                                    oSubTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                                    oSubTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                                ElseIf colCount2 = 7 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.25, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.25, inches))
                                    oSubTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.125, inches))
                                    oSubTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.125, inches))
                                    oSubTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.125, inches))
                                    oSubTable.Columns.Item(6).Width = WordApp.InchesToPoints(GetApproxInches(0.125, inches))
                                ElseIf colCount2 = 8 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.25, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.25, inches))
                                    oSubTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.0833, inches))
                                    oSubTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.0833, inches))
                                    oSubTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.0833, inches))
                                    oSubTable.Columns.Item(6).Width = WordApp.InchesToPoints(GetApproxInches(0.0833, inches))
                                    oSubTable.Columns.Item(7).Width = WordApp.InchesToPoints(GetApproxInches(0.0834, inches))

                                ElseIf colCount2 = 9 Then
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.2, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.2, inches))
                                    oSubTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                                    oSubTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                                    oSubTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                                    oSubTable.Columns.Item(6).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                                    oSubTable.Columns.Item(7).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))
                                    oSubTable.Columns.Item(8).Width = WordApp.InchesToPoints(GetApproxInches(0.1, inches))

                                Else
                                    oSubTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.2, inches))
                                    oSubTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.2, inches))
                                    oSubTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.0855, inches))
                                    oSubTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.0855, inches))
                                    oSubTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.08555, inches))
                                    oSubTable.Columns.Item(6).Width = WordApp.InchesToPoints(GetApproxInches(0.08555, inches))
                                    oSubTable.Columns.Item(7).Width = WordApp.InchesToPoints(GetApproxInches(0.0855, inches))
                                    oSubTable.Columns.Item(8).Width = WordApp.InchesToPoints(GetApproxInches(0.0855, inches))
                                    oSubTable.Columns.Item(9).Width = WordApp.InchesToPoints(GetApproxInches(0.08545, inches))

                                End If
                            Catch
                            End Try



                            ' row 2
                            With oSubTable.Rows.Item(1).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                .Texture = Word.WdTextureIndex.wdTexture25Percent
                            End With

                            ' row 3
                            With oSubTable.Rows.Item(2).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            End With

                            ' Fill Data
                            oSubTable.Rows.Item(1).Range.Font.Size = 9
                            oSubTable.Rows.Item(1).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                            oSubTable.Rows.Item(2).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter


                            For colIndex = 2 To colCount2
                                Try
                                    ' row 2 - header
                                    oSubTable.Rows.Item(1).Cells().Item(colIndex - 1).Range.Text = dt.Columns(colCount1 + colIndex - 2).ColumnName

                                    ' row 3 - data
                                    oSubTable.Rows.Item(2).Cells().Item(colIndex - 1).Range.Text = dt.Rows(0)(colCount1 + colIndex - 2)
                                Catch
                                End Try


                            Next

                        End If
                    Next
                End With
            Catch ex As Exception
                Throw ex
            End Try
        End Sub
        Private Sub InspectionAddTerms(ByVal ds As DataSet, ByRef WordApp As Word.Application, ByRef DestDoc As Word.Document, ByRef start As Boolean)
            Dim dt As DataTable
            Dim i, colCount, colIndex As Integer
            Dim inches As Double = 7.5
            Dim oTable As Word.Table
            Dim oPara As Word.Paragraph
            Try
                With DestDoc
                    ' Add Term Table
                    For i = 0 To ds.Tables.Count - 1

                        dt = ds.Tables(i)
                        ' to determine how many cells to display in first row of a given term
                        colCount = 8
                        If Not dt.Columns.Contains("Tank Term. CP") Then
                            colCount -= 1
                        End If
                        If Not dt.Columns.Contains("Dispenser Term. CP") Then
                            colCount -= 1
                        End If
                        If Not dt.Columns.Contains("Term. CP Tested") Then
                            colCount -= 1
                        End If

                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        'oPara.Range.Text = "<Space>"
                        oPara.Format.SpaceBefore = 1
                        oPara.Format.SpaceAfter = 0
                        oPara.KeepWithNext = True


                        oPara.Range.InsertParagraphAfter()
                        oTable = .Tables.Add(oPara.Range, IIf(i = 0, 3, 2), colCount)
                        oTable.Range.ParagraphFormat.KeepWithNext = True
                        oTable.Range.Font.Name = "Arial"
                        oTable.Range.Font.Size = 8
                        oTable.Range.Font.Bold = True
                        oTable.Range.ParagraphFormat.SpaceAfter = 1
                        oTable.Range.ParagraphFormat.LineSpacing = 9
                        oTable.Range.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceExactly

                        oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True


                        oTable.Columns.Item(1).Width = WordApp.InchesToPoints(GetApproxInches(0.08, inches))
                        oTable.Columns.Item(2).Width = WordApp.InchesToPoints(GetApproxInches(0.09, inches))
                        oTable.Columns.Item(3).Width = WordApp.InchesToPoints(GetApproxInches(0.09, inches))

                        If colCount = 5 Then
                            oTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.37, inches))
                            oTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.37, inches))
                        ElseIf colCount = 6 Then
                            oTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.25, inches))
                            oTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.25, inches))
                            oTable.Columns.Item(6).Width = WordApp.InchesToPoints(GetApproxInches(0.24, inches))
                        ElseIf colCount = 7 Then
                            oTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.22, inches))
                            oTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.22, inches))
                            oTable.Columns.Item(6).Width = WordApp.InchesToPoints(GetApproxInches(0.15, inches))
                            oTable.Columns.Item(7).Width = WordApp.InchesToPoints(GetApproxInches(0.15, inches))
                        ElseIf colCount = 8 Then
                            oTable.Columns.Item(4).Width = WordApp.InchesToPoints(GetApproxInches(0.18, inches))
                            oTable.Columns.Item(5).Width = WordApp.InchesToPoints(GetApproxInches(0.18, inches))
                            oTable.Columns.Item(6).Width = WordApp.InchesToPoints(GetApproxInches(0.12, inches))
                            oTable.Columns.Item(7).Width = WordApp.InchesToPoints(GetApproxInches(0.12, inches))
                            oTable.Columns.Item(8).Width = WordApp.InchesToPoints(GetApproxInches(0.14, inches))
                        End If

                        ' row 1
                        If i = 0 Then
                            oTable.Rows.Item(1).Cells.Merge()
                            oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            oTable.Rows.Item(1).Range.Text = "PIPING TERMINATIONS"
                            oTable.Rows.Item(1).Range.Font.Size = 12

                            With oTable.Rows.Item(1).Shading
                                '.BackgroundPatternColor = Word.WdColor.wdColorBlack
                                .BackgroundPatternColor = Word.WdColor.wdColorBlack

                                .Texture = Word.WdTextureIndex.wdTextureSolid
                            End With
                        End If

                        Dim a As Integer = 2
                        Dim b As Integer = 3

                        If i >= 1 Then
                            a = 1
                            b = 2
                        End If


                        ' row 2
                        With oTable.Rows.Item(a).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture30Percent
                        End With

                        ' row 3
                        With oTable.Rows.Item(b).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                        End With

                        ' Fill Data
                        oTable.Rows.Item(a).Range.Font.Size = 9
                        oTable.Rows.Item(a).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
                        oTable.Rows.Item(b).Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter

                        For colIndex = 1 To colCount
                            ' row 2 - header
                            oTable.Rows.Item(a).Cells().Item(colIndex).Range.Text = dt.Columns(colIndex - 1).ColumnName

                            ' row 3 - data

                            If dt.Rows(0)(colIndex - 1) = "<CAPNEEDED>" Then
                                oTable.Rows.Item(b).Cells().Item(colIndex).Range.Text = String.Empty
                                oTable.Rows.Item(b).Cells().Item(colIndex).Shading.BackgroundPatternColor = Word.WdColor.wdColorYellow
                            Else

                                oTable.Rows.Item(b).Cells().Item(colIndex).Range.Text = dt.Rows(0)(colIndex - 1)
                            End If


                        Next


                        oTable.Rows.Item(a).LeftIndent = 20.0
                        oTable.Rows.Item(b).LeftIndent = 20.0

                    Next

                End With
            Catch ex As Exception
                Throw ex
            End Try
        End Sub


#End Region

#Region "General Operations"

        Private Sub LetterTemplateGenericPageSetup(ByRef wordApp As Word.Application, ByRef destDoc As Word.Document)

            With destDoc.PageSetup
                .LineNumbering.Active = False
                '.Orientation = Word.WdOrientation.wdOrientLandscape
                .Orientation = Word.WdOrientation.wdOrientPortrait
                .TopMargin = wordApp.InchesToPoints(0.88)
                .BottomMargin = wordApp.InchesToPoints(1.25)
                .LeftMargin = wordApp.InchesToPoints(1)
                .RightMargin = wordApp.InchesToPoints(1)
                .Gutter = wordApp.InchesToPoints(0)
                .HeaderDistance = wordApp.InchesToPoints(0.5)
                .FooterDistance = wordApp.InchesToPoints(0.5)
                '.PageWidth = WordApp.InchesToPoints(11)
                .PageHeight = wordApp.InchesToPoints(8.5)
                .FirstPageTray = Word.WdPaperTray.wdPrinterDefaultBin
                .OtherPagesTray = Word.WdPaperTray.wdPrinterDefaultBin
                .SectionStart = Word.WdSectionStart.wdSectionNewPage
                .OddAndEvenPagesHeaderFooter = False
                .DifferentFirstPageHeaderFooter = False
                .VerticalAlignment = Word.WdVerticalAlignment.wdAlignVerticalTop
                .SuppressEndnotes = False
                .MirrorMargins = False
                .TwoPagesOnOne = False
                .BookFoldPrinting = False
                .BookFoldRevPrinting = False
                .BookFoldPrintingSheets = 1
                .GutterPos = Word.WdGutterStyle.wdGutterPosLeft
            End With
        End Sub

        Private Function GetWordApp() As Word.Application
            Dim WordApp As Word.Application
            Try
                If IsNothing(WordApp) Then
                    WordApp = GetObject(, "Word.Application")
                End If

            Catch ex As Exception
                If ex.Message.ToUpper = "Cannot Create ActiveX Component.".ToUpper Then
                    WordApp = New Word.Application
                ElseIf ex.Message.ToUpper = "The RPC server is unavailable.".ToUpper Then
                    WordApp = New Word.Application
                ElseIf ex.Message.ToUpper.Contains("The RPC server is unavailable.".ToUpper) Then
                    WordApp = New Word.Application
                Else
                    Throw ex
                End If
            End Try

            Return WordApp

        End Function


        Private Function LetterTemplateSave(ByRef WordApp As Word.Application, ByVal templatePath As String, ByVal DestinationPath As String, Optional ByVal KeepClosed As Boolean = False) As Word.Application
            Dim lastSavedDoc As String = String.Empty
            Dim Keepsaving As Boolean = True
            Dim cnt As Integer = 1

            Try

                With WordApp

                    DestDoc = .Documents.Open(DestinationPath.Replace(".doc", "_TEMPLATE.doc"))
                    DestDoc = WordApp.ActiveDocument

                    If Not DestDoc Is Nothing Then
                        With DestDoc

                            While Keepsaving

                                Try

                                    Keepsaving = False

                                    If cnt <= 1 Then
                                        lastSavedDoc = DestinationPath
                                    Else
                                        lastSavedDoc = DestinationPath.Replace(".doc", String.Format("_{0}.doc", cnt))
                                    End If

                                    .SaveAs(lastSavedDoc)

                                Catch ex As Exception

                                    If ex.ToString.ToUpper.IndexOf(" PERMISSION") > -1 Then

                                        Threading.Thread.Sleep(1000)
                                        cnt += 1
                                        Keepsaving = True
                                    Else
                                        .Close()

                                    End If
                                End Try
                            End While

                            .Close()

                        End With
                    End If

                    WordApp.Documents.Open(lastSavedDoc)

                    If Not KeepClosed Then
                        .Visible = True
                    End If


                    Try

                        If File.Exists(DestinationPath.Replace(".doc", "_TEMPLATE.doc")) Then
                            System.IO.File.Delete(DestinationPath.Replace(".doc", "_TEMPLATE.doc"))
                        End If

                    Catch ex As Exception
                        Throw New FieldAccessException("Unable to remove template copy of  " & DestinationPath & " in pLetterGen object.")
                    End Try


                    Return WordApp

                End With


            Catch ex As Exception
                SrcDoc = Nothing
                If Not WordApp Is Nothing Then
                    If Not WordApp.ActiveDocument Is Nothing Then WordApp.ActiveDocument.Close(False)
                End If

                If ex.InnerException Is Nothing OrElse Not TypeOf ex.InnerException Is FieldAccessException Then

                    If File.Exists(DestinationPath) Then
                        System.IO.File.Delete(DestinationPath)
                    End If

                    If File.Exists(DestinationPath.Replace(".doc", "_TEMPLATE.doc")) Then
                        System.IO.File.Delete(DestinationPath.Replace(".doc", "_TEMPLATE.doc"))
                    End If
                End If

                Throw ex

            End Try

        End Function

        Private Function LetterTemplateInit(ByRef WordApp As Word.Application, ByVal templatePath As String, ByVal DestinationPath As String)

            'Instantiate the Word Object
            If IsNothing(WordApp) Then
                WordApp = GetWordApp()
            End If

            If Not System.IO.File.Exists(templatePath) Then
                Throw New Exception("File Not Found: " + templatePath)
            End If
            System.IO.File.Copy(templatePath, DestinationPath.Replace(".doc", "_TEMPLATE.doc"), True)

            If System.IO.File.Exists(DestinationPath.Replace(".doc", "_TEMPLATE.doc")) Then
                With WordApp


                    DestDoc = .Documents.Open(DestinationPath.Replace(".doc", "_TEMPLATE.doc"))


                    DestDoc = WordApp.ActiveDocument
                    .Visible = False
                    DestDoc.Activate()
                End With
            Else
                Throw New Exception("Unable to copy template " & templatePath & " to " & DestinationPath & " in pLetterGen object.")

            End If

        End Function

        Private Function TemporaryDocInit(ByRef WordApp As Word.Application, ByVal templatePath As String)

            'Instantiate the Word Object
            WordApp = Nothing
            If IsNothing(WordApp) Then
                WordApp = GetWordApp()
            End If

            If Not System.IO.File.Exists(templatePath) Then
                Throw New Exception("File Not Found: " + templatePath)
            End If

            'With WordApp


            Try

                If Not DestDoc Is Nothing Then
                    DestDoc.Close()
                End If
                If WordApp.Visible Then
                    WordApp.Visible = False
                End If
                WordApp.Visible = True
                DestDoc = WordApp.Documents.Open(FileName:=templatePath, ReadOnly:=True)
                DestDoc.Activate()

            Catch ex As Exception
                If Not DestDoc Is Nothing Then
                    DestDoc.Close()
                End If

                Throw New Exception(String.Format("Unale to set up document from {0}", templatePath))

            End Try


            'End With

        End Function

        Public Sub SetUpList(ByRef WordApp As Word.Application, ByRef temp As WordTemplate)
            WordTemplate.ExtractTemplateFromFile(temp, temp.path)
        End Sub

        Public Sub ApplyToList(ByRef WordApp As Word.Application, ByVal temp As WordTemplate, ByVal multiFacs As Boolean)

            Dim wrdglobal As New Word.[Global]
            Dim wtemp As Word.ListTemplate

            wtemp = wrdglobal.ListGalleries.Item(Word.WdListGalleryType.wdOutlineNumberGallery).ListTemplates.Item(1)
            temp.ExtractTemplateDatafromRecord(wtemp, multiFacs)

            WordApp.Selection.Range.ListFormat.ApplyListTemplate(ListTemplate:=wtemp, _
               ContinuePreviousList:=False, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, _
               DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)

            wrdglobal = Nothing
        End Sub
#End Region

#End Region
    End Class
End Namespace
