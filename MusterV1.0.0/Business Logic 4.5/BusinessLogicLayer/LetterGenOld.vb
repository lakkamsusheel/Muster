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
Namespace MUSTER.BusinessLogic
    <Serializable()> _
        Public Class pLetterGen

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
            MyBase.new()
        End Sub
#End Region
#Region "Exposed Operations"
#Region "Collection Operations"
        Public Function CreateLetter(ByVal strModuleID As String, ByVal strLetter_To_Print As String, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal strfile As String = "", Optional ByVal strSignature As String = "") ', Optional ByVal strFiles As String = "")
            Try
                Dim DocumentPath As String
                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty
                Dim strToday As String = String.Empty
                Dim strDocName As String = String.Empty

                Dim i As Integer = 0
                'Dim strEnvelopes() As String = strFiles.Split(",")

                'Instantiate the Word Object
                If IsNothing(WordApp) Then
                    WordApp = GetWordApp()
                End If

                If Not System.IO.File.Exists(TemplatePath) Then
                    Throw New Exception("File Not Found: " + TemplatePath)
                End If
                System.IO.File.Copy(TemplatePath, DestinationPath)

                If System.IO.File.Exists(DestinationPath) Then
                    With WordApp

                        DestDoc = .Documents.Open(DestinationPath)
                        DestDoc = WordApp.ActiveDocument

                        With DestDoc
                            .Activate()

                            ' Find and Replace the TAGs with Values.
                            For i = 0 To colParams.Count - 1
                                strKey = colParams.Keys(i).ToString
                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)

                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                If strKey = "<Reasons>" Then
                                    If .Tables.Count > 0 Then
                                        .Tables.Item(1).Cell(1, 1).Range.Text = strValue
                                    End If
                                Else
                                    .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                                End If
                            Next



                            'WordApp.Run("EndOfDocument")
                            'If Not strFiles = String.Empty Then
                            '    'WordApp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)
                            '    For j As Integer = 0 To UBound(strEnvelopes)
                            '        'WordApp.Selection.InsertFile(FILENAME:=strEnvelopes(j), Range:="", _
                            '        'ConfirmConversions:=False, Link:=False, Attachment:=False)
                            '        'strEnvelopes(j)

                            '    Next

                            'End If

                            ' code to insert envelopes to an existing document
                            'WordApp.ActiveDocument.Envelope.Insert(ExtractAddress:=False, OmitReturnAddress:= _
                            'False, PrintBarCode:=False, PrintFIMA:=False, Height:=WordApp.InchesToPoints(4.13 _
                            '), Width:=WordApp.InchesToPoints(9.5), Address:="875 william blvd" & Chr(13) & _
                            '"ridgeland, ms 39157", AutoText:="", ReturnAddress:="", ReturnAutoText:= _
                            '"", AddressFromLeft:=WordApp.wdAutoPosition, AddressFromTop:=WordApp.wdAutoPosition, _
                            'ReturnAddressFromLeft:=WordApp.wdAutoPosition, ReturnAddressFromTop:= _
                            'WordApp.wdAutoPosition, DefaultOrientation:=WordApp.wdCenterLandscape, DefaultFaceUp:= _
                            'True, PrintEPostage:=False)
                            Dim strPhoto As String = String.Empty
                            If strfile <> String.Empty Then
                                strPhoto = UCase(strfile)
                            End If
                            If strfile <> String.Empty And strPhoto.IndexOf(UCase("\\Opcgw\MUSTER\Images\Licensees\Nophoto.gif")) < 0 Then
                                '"Z:\Images\Licensees\1002.gif"
                                'WordApp.ActiveDocument.Tables(1).Rows(1).Cells(1).Tables(1)

                                With .Tables.Item(1).Cell(1, 1).Tables.Item(1).Cell(1, 1)
                                    Dim oPic As Word.InlineShape
                                    oPic = .Range.InlineShapes.AddPicture(FILENAME:=strfile _
                                        , LinkToFile:=False, SaveWithDocument:=True)
                                    oPic.Height = 55
                                    oPic.Width = 45
                                End With

                                'oPic = .InlineShapes.AddPicture(FILENAME:=strfile _
                                '        , LinkToFile:=False, SaveWithDocument:=True)
                                'oPic.Height = 70
                                'oPic.Width = 70
                            End If

                            If strSignature <> String.Empty And strSignature.IndexOf("\\Opcgw\MUSTER\Images\Licensees\NoSignature.gif") < 0 Then

                                With .Tables.Item(1).Cell(1, 1).Tables.Item(1).Cell(3, 2)
                                    Dim oPic As Word.InlineShape
                                    oPic = .Range.InlineShapes.AddPicture(FILENAME:=strSignature _
                                            , LinkToFile:=False, SaveWithDocument:=True)
                                    oPic.Height = 25
                                    oPic.Width = 70
                                End With

                                'WordApp.Selection.MoveDown(Word.WdUnits.wdLine, Count:=7)
                                'Dim oPicture As Word.InlineShape
                                ''"Z:\Images\Licensees\1002.gif"
                                'oPicture = .InlineShapes.AddPicture(FILENAME:=strSignature _
                                '        , LinkToFile:=False, SaveWithDocument:=True)
                                'oPicture.Height = 70
                                'oPicture.Width = 70
                            End If
                            ' DestDoc.Envelope.Insert(ExtractAddress:=True, OmitReturnAddress:= _
                            'False, PrintBarCode:=False, PrintFIMA:=False, Height:=WordApp.InchesToPoints(4.13 _
                            '), Width:=WordApp.InchesToPoints(9.5), Address:="Mr James " & Chr(13) & "1234", AutoText:= _
                            '    "ToolsCreateLabels", ReturnAddress:="", ReturnAutoText:= _
                            '"ToolsCreateLabels", AddressFromLeft:=Word.WdConstants.wdAutoPosition, AddressFromTop:= _
                            'Word.WdConstants.wdAutoPosition, ReturnAddressFromLeft:=Word.WdConstants.wdAutoPosition, _
                            'ReturnAddressFromTop:=Word.WdConstants.wdAutoPosition, DefaultOrientation:= _
                            'Word.WdEnvelopeOrientation.wdCenterLandscape, DefaultFaceUp:=True, PrintEPostage:=False)
                            ' .Save()
                            .Save()

                        End With

                        'per Danny do not close the Word App when creating Letters.
                        'DestDoc = Nothing
                        '.Quit(False)
                        .Visible = True
                        '.ScreenRefresh()
                        '.ShowMe()
                    End With
                Else
                    Throw New Exception("Unable to copy template " & TemplatePath & " to " & DestinationPath & " in pLetterGen object.")
                End If
            Catch ex As Exception
                SrcDoc = Nothing
                If Not WordApp Is Nothing Then
                    If Not WordApp.ActiveDocument Is Nothing Then WordApp.ActiveDocument.Close(False)
                End If
                System.IO.File.Delete(DestinationPath)
                Throw ex
            End Try
        End Function
        Public Function CreateLabels(ByVal strModuleID As String, ByVal strLetter_To_Print As String, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal strAddress As String = "", Optional ByVal nRow As Integer = 1, Optional ByVal nColumn As Integer = 1) ', Optional ByVal strFiles As String = "")
            Try
                Dim DocumentPath As String
                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty
                Dim strToday As String = String.Empty
                Dim strDocName As String = String.Empty

                Dim i As Integer = 0
                'Dim strEnvelopes() As String = strFiles.Split(",")

                'Instantiate the Word Object
                If IsNothing(WordApp) Then
                    WordApp = GetWordApp()
                End If

                If Not System.IO.File.Exists(TemplatePath) Then
                    Throw New Exception("File Not Found: " + TemplatePath)
                End If
                System.IO.File.Copy(TemplatePath, DestinationPath)

                If System.IO.File.Exists(DestinationPath) Then
                    With WordApp

                        DestDoc = .Documents.Open(DestinationPath)
                        DestDoc = WordApp.ActiveDocument

                        With DestDoc
                            .Activate()

                            ' Find and Replace the TAGs with Values.
                            For i = 0 To colParams.Count - 1
                                strKey = colParams.Keys(i).ToString
                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                            Next

                            If .Tables.Count > 0 Then
                                .Tables.Item(1).Cell(nRow, nColumn).Range.Text = strAddress
                            End If
                            .Save()

                        End With

                        'per Danny do not close the Word App when creating Letters.
                        'DestDoc = Nothing
                        '.Quit(False)
                        .Visible = True
                    End With
                Else
                    Throw New Exception("Unable to copy template " & TemplatePath & " to " & DestinationPath & " in pLetterGen object.")
                End If
            Catch ex As Exception
                SrcDoc = Nothing
                If Not WordApp Is Nothing Then
                    If Not WordApp.ActiveDocument Is Nothing Then WordApp.ActiveDocument.Close(False)
                End If
                System.IO.File.Delete(DestinationPath)
                Throw ex
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
                If IsNothing(WordApp) Then
                    WordApp = GetWordApp()
                End If

                If Not System.IO.File.Exists(TemplatePath) Then
                    Throw New Exception("File Not Found: " + TemplatePath)
                End If
                System.IO.File.Copy(TemplatePath, DestinationPath)

                If System.IO.File.Exists(DestinationPath) Then
                    With WordApp

                        DestDoc = .Documents.Open(DestinationPath)
                        DestDoc = WordApp.ActiveDocument

                        With DestDoc
                            .Activate()

                            ' Find and Replace the TAGs with Values.
                            For i = 0 To colParams.Count - 1
                                strKey = colParams.Keys(i).ToString
                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                            Next

                            .Save()

                        End With

                        'per Danny do not close the Word App when creating Letters.
                        'DestDoc = Nothing
                        '.Quit(False)
                        .Visible = False
                        '.ScreenRefresh()
                        '.ShowMe()
                    End With
                Else
                    Throw New Exception("Unable to copy template " & TemplatePath & " to " & DestinationPath & " in pLetterGen object.")
                End If
            Catch ex As Exception
                SrcDoc = Nothing
                If Not WordApp Is Nothing Then
                    If Not WordApp.ActiveDocument Is Nothing Then WordApp.ActiveDocument.Close(False)
                End If
                System.IO.File.Delete(DestinationPath)
                Throw ex
            End Try
        End Function
        Public Function CreateGenericLetter(ByVal strModuleID As String, ByVal strLetter_To_Print As String, ByVal colParams As Specialized.NameValueCollection, ByVal cols As Int16, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal TableFormat As Int16 = 35, Optional ByVal bApplyBorders As Boolean = False, Optional ByVal bApplyShading As Boolean = False, Optional ByVal bApplyFont As Boolean = True, Optional ByVal bApplyColor As Boolean = False, Optional ByVal bApplyHeader As Boolean = True, Optional ByVal bApplyLastRow As Boolean = False, Optional ByVal bApplyFirstCol As Boolean = False, Optional ByVal bApplyLastCol As Boolean = False, Optional ByVal bApplyAutofit As Boolean = True, Optional ByVal attachDocumentInfo As String = "")
            Try
                Dim DocumentPath As String
                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty
                Dim strToday As String = String.Empty
                Dim strDocName As String = String.Empty
                Dim i As Integer = 0
                Dim oPara As Word.Paragraph

                'Instantiate the Word Object
                If IsNothing(WordApp) Then
                    WordApp = GetWordApp()
                End If

                If Not System.IO.File.Exists(TemplatePath) Then
                    Throw New Exception("File Not Found: " + TemplatePath)
                End If
                System.IO.File.Copy(TemplatePath, DestinationPath)

                If System.IO.File.Exists(DestinationPath) Then
                    With WordApp

                        DestDoc = .Documents.Open(DestinationPath)
                        DestDoc = WordApp.ActiveDocument

                        With DestDoc
                            .Activate()

                            ' Find and Replace the TAGs with Values.
                            For i = 0 To colParams.Count - 1

                                strKey = colParams.Keys(i).ToString
                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                If strKey = "<DATA>" Then
                                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                    oPara.Range.Text = strValue
                                    'Word.WdTableFormat.wdTableFormatContemporary
                                    oPara.Range.ConvertToTable("|", , cols, , TableFormat, bApplyBorders, bApplyShading, bApplyFont, bApplyColor, bApplyHeader, bApplyLastRow, bApplyFirstCol, bApplyLastCol, bApplyAutofit)
                                    oPara.Format.SpaceAfter = 1
                                    oPara.Range.InsertParagraphAfter()

                                Else
                                    .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                                End If
                            Next



                            ' attach document
                            If attachDocumentInfo <> "" Then
                                If System.IO.File.Exists(attachDocumentInfo) Then
                                    Try
                                        WordApp.Selection.GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:=.Bookmarks.Item("\endofdoc").Name)
                                        WordApp.Selection.InsertFile(FILENAME:=attachDocumentInfo, Range:=.Bookmarks.Item("\endofdoc").Range, ConfirmConversions:=False, Link:=False, Attachment:=False)
                                    Catch ex As Exception
                                        ' do nothing
                                    End Try
                                End If
                            End If

                            .Save()

                        End With

                        'per Danny do not close the Word App when creating Letters.
                        'DestDoc = Nothing
                        '.Quit(False)
                    End With
                Else
                    Throw New Exception("Unable to copy template " & TemplatePath & " to " & DestinationPath & " in pLetterGen object.")
                End If

            Catch ex As Exception
                SrcDoc = Nothing
                If Not WordApp Is Nothing Then WordApp.Quit(False)
                Throw ex
            End Try

        End Function

        ' Registration Letters
        Public Function CreateOtherRegistrationLetter(ByVal ownerID As Integer, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, ByVal facsForOwner As DataTable, ByVal oOwner As MUSTER.BusinessLogic.pOwner, Optional ByRef WordApp As Word.Application = Nothing)
            Dim strKey As String = String.Empty
            Dim strValue As String = String.Empty
            Dim i, j As Integer
            Try
                'Instantiate the Word Object
                If IsNothing(WordApp) Then
                    WordApp = GetWordApp()
                End If

                If Not System.IO.File.Exists(TemplatePath) Then
                    Throw New Exception("File Not Found: " + TemplatePath)
                End If
                System.IO.File.Copy(TemplatePath, DestinationPath)

                If System.IO.File.Exists(DestinationPath) Then
                    With WordApp

                        DestDoc = .Documents.Open(DestinationPath)
                        DestDoc = WordApp.ActiveDocument

                        With DestDoc
                            .Activate()

                            ' Find and Replace the TAGs with Values.
                            For i = 0 To colParams.Count - 1
                                strKey = colParams.Keys(i).ToString
                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                            Next


                            ' Facility Table
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

                            .Save()

                        End With

                        DestDoc = Nothing

                    End With
                Else
                    Throw New Exception("Unable to copy template " & TemplatePath & " to " & DestinationPath & " in pLetterGen object.")
                End If
            Catch ex As Exception
                SrcDoc = Nothing
                If Not WordApp Is Nothing Then
                    If Not WordApp.ActiveDocument Is Nothing Then WordApp.ActiveDocument.Close(False)
                End If
                System.IO.File.Delete(DestinationPath)
                Throw ex
            End Try
        End Function
        Public Function CreateRegistrationLetter(ByVal ownerID As Integer, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, ByVal facsForOwner As DataTable, ByVal oOwner As MUSTER.BusinessLogic.pOwner, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal showTransferOwnerSection As Boolean = False)
            Dim strKey As String = String.Empty
            Dim strValue As String = String.Empty
            Dim i, j As Integer
            Dim showNewOwnerSection As Boolean = False
            Dim showFeeSection As Boolean = False
            Dim showTOISSection As Boolean = False
            Try
                'Instantiate the Word Object
                If IsNothing(WordApp) Then
                    WordApp = GetWordApp()
                End If

                If Not System.IO.File.Exists(TemplatePath) Then
                    Throw New Exception("File Not Found: " + TemplatePath)
                End If
                System.IO.File.Copy(TemplatePath, DestinationPath)

                If System.IO.File.Exists(DestinationPath) Then
                    With WordApp

                        DestDoc = .Documents.Open(DestinationPath)
                        DestDoc = WordApp.ActiveDocument

                        With DestDoc
                            .Activate()

                            ' Find and Replace the TAGs with Values.
                            For i = 0 To colParams.Count - 1
                                strKey = colParams.Keys(i).ToString
                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                            Next


                            ' Facility Table
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

                            showNewOwnerSection = oOwner.OwnerL2CSnippet

                            Dim ds As DataSet
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

                            .Save()

                        End With

                        DestDoc = Nothing

                    End With
                Else
                    Throw New Exception("Unable to copy template " & TemplatePath & " to " & DestinationPath & " in pLetterGen object.")
                End If
            Catch ex As Exception
                SrcDoc = Nothing
                If Not WordApp Is Nothing Then
                    If Not WordApp.ActiveDocument Is Nothing Then WordApp.ActiveDocument.Close(False)
                End If
                System.IO.File.Delete(DestinationPath)
                Throw ex
            End Try
        End Function
        Public Function CreateComplianceLetter(ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal facHasAllTOSITanks As Boolean = False)
            Dim strKey As String = String.Empty
            Dim strValue As String = String.Empty
            Dim i, j As Integer
            Try
                'Instantiate the Word Object
                If IsNothing(WordApp) Then
                    WordApp = GetWordApp()
                End If

                If Not System.IO.File.Exists(TemplatePath) Then
                    Throw New Exception("File Not Found: " + TemplatePath)
                End If
                System.IO.File.Copy(TemplatePath, DestinationPath)

                If System.IO.File.Exists(DestinationPath) Then
                    With WordApp

                        DestDoc = .Documents.Open(DestinationPath)
                        DestDoc = WordApp.ActiveDocument

                        With DestDoc
                            .Activate()

                            ' Find and Replace the TAGs with Values.
                            For i = 0 To colParams.Count - 1
                                strKey = colParams.Keys(i).ToString
                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                            Next


                            If facHasAllTOSITanks Then
                                .Tables.Item(1).Delete()
                            End If

                            .Save()

                        End With

                        DestDoc = Nothing

                    End With
                Else
                    Throw New Exception("Unable to copy template " & TemplatePath & " to " & DestinationPath & " in pLetterGen object.")
                End If
            Catch ex As Exception
                SrcDoc = Nothing
                If Not WordApp Is Nothing Then
                    If Not WordApp.ActiveDocument Is Nothing Then WordApp.ActiveDocument.Close(False)
                End If
                System.IO.File.Delete(DestinationPath)
                Throw ex
            End Try
        End Function

        ' Closure Letters
        Public Function CreateClosureDemo(ByVal strModuleID As String, ByVal strLetter_To_Print As String, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByVal dtSample As DataTable = Nothing, Optional ByRef WordApp As Word.Application = Nothing)
            Try
                Dim DocumentPath As String
                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty
                Dim strToday As String = String.Empty
                Dim strDocName As String = String.Empty
                Dim i As Integer = 0

                'Instantiate the Word Object
                If IsNothing(WordApp) Then
                    WordApp = GetWordApp()
                End If

                If Not System.IO.File.Exists(TemplatePath) Then
                    Throw New Exception("File Not Found: " + TemplatePath)
                End If
                System.IO.File.Copy(TemplatePath, DestinationPath)

                If System.IO.File.Exists(DestinationPath) Then
                    With WordApp

                        DestDoc = .Documents.Open(DestinationPath)
                        DestDoc = WordApp.ActiveDocument

                        With DestDoc
                            .Activate()

                            ' Find and Replace the TAGs with Values.
                            For i = 0 To colParams.Count - 1
                                strKey = colParams.Keys(i).ToString
                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
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

                            .Save()
                        End With
                        DestDoc = Nothing
                    End With
                Else
                    Throw New Exception("Unable to copy template " & TemplatePath & " to " & DestinationPath & " in pLetterGen object.")
                End If
            Catch ex As Exception
                SrcDoc = Nothing
                If Not WordApp Is Nothing Then
                    If Not WordApp.ActiveDocument Is Nothing Then WordApp.ActiveDocument.Close(False)
                End If
                System.IO.File.Delete(DestinationPath)
                Throw ex
            End Try
        End Function
        Public Function CreateClosureInfoNeeded(ByVal TemplatePath As String, ByVal DestinationPath As String, ByVal colParams As Specialized.NameValueCollection, ByVal colInfoNeeded As ArrayList, Optional ByRef WordApp As Word.Application = Nothing)
            Try
                Dim DocumentPath As String
                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty
                Dim strToday As String = String.Empty
                Dim strDocName As String = String.Empty
                Dim oPara As Word.Paragraph
                Dim i As Integer = 0

                'Instantiate the Word Object
                If IsNothing(WordApp) Then
                    WordApp = GetWordApp()
                End If

                If Not System.IO.File.Exists(TemplatePath) Then
                    Throw New Exception("File Not Found: " + TemplatePath)
                End If
                System.IO.File.Copy(TemplatePath, DestinationPath)

                If System.IO.File.Exists(DestinationPath) Then
                    With WordApp

                        DestDoc = .Documents.Open(DestinationPath)
                        DestDoc = WordApp.ActiveDocument

                        With DestDoc
                            .Activate()

                            ' Find and Replace the TAGs with Values.
                            For i = 0 To colParams.Count - 1
                                strKey = colParams.Keys(i).ToString
                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                            Next


                            ' Add Info Needed
                            .Tables.Item(1).Cell(1, 1).Range.Text = " - " + colInfoNeeded.Item(0).ToString
                            For i = 1 To colInfoNeeded.Count - 1
                                With .Tables.Item(1)
                                    .Rows.Add()
                                    .Cell(i + 1, 1).Range.Text = " - " + colInfoNeeded.Item(i).ToString
                                End With
                            Next
                            .Save()
                        End With
                        DestDoc = Nothing
                    End With
                Else
                    Throw New Exception("Unable to copy template " & TemplatePath & " to " & DestinationPath & " in pLetterGen object.")
                End If
            Catch ex As Exception
                SrcDoc = Nothing
                If Not WordApp Is Nothing Then
                    If Not WordApp.ActiveDocument Is Nothing Then WordApp.ActiveDocument.Close(False)
                End If
                System.IO.File.Delete(DestinationPath)
                Throw ex
            End Try
        End Function

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
                'Instantiate the Word Object
                If IsNothing(WordApp) Then
                    WordApp = GetWordApp()
                End If

                If Not System.IO.File.Exists(TemplatePath) Then
                    Throw New Exception("File Not Found: " + TemplatePath)
                End If

                ' check if user has rights to save inspection
                ' will be saving letter generated = true
                If Not oOwner.CheckWriteAccess(moduleID, staffID, SqlHelper.EntityTypes.Inspection) Then
                    returnVal = "You do not have rights to save Inspection."
                    Exit Function
                End If

                System.IO.File.Copy(TemplatePath, DestinationPath)

                If System.IO.File.Exists(DestinationPath) Then
                    With WordApp

                        DestDoc = .Documents.Open(DestinationPath)
                        DestDoc = WordApp.ActiveDocument

                        With DestDoc
                            .Activate()

                            ' Find and Replace the TAGs with Values.
                            For i = 0 To colParams.Count - 1
                                strKey = colParams.Keys(i).ToString
                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                strValue = strValue.Replace(vbCrLf, Chr(13))
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
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

                                InitVars(bolVars)
                                ProcessVars(bolVars, oOwner.Facility, facLastInspDate)

                                ' records
                                ProcessRecords(bolVars, WordApp, DestDoc)

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

                                ProcessComponents(bolVars, WordApp, DestDoc)

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

                            .Save()
                        End With
                        DestDoc = Nothing
                        '.ActiveDocument.Close(False)
                    End With
                Else
                    Throw New Exception("Unable to copy template " & TemplatePath & " to " & DestinationPath & " in pLetterGen object.")
                End If
            Catch ex As Exception
                bolDeleteFile = True
                Throw ex
            Finally
                If bolDeleteFile Then
                    SrcDoc = Nothing
                    If Not WordApp Is Nothing Then
                        If Not WordApp.ActiveDocument Is Nothing Then WordApp.ActiveDocument.Close(False)
                    End If
                    System.IO.File.Delete(DestinationPath)
                    RaiseEvent CloseCheckListProgress()
                End If
            End Try
        End Function
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
        Private Sub ProcessRecords(ByRef bolVars As InspPrepGuidelineVars, ByRef WordApp As Word.Application, ByRef DestDoc As Word.Document)
            Dim oPara As Word.Paragraph
            Try
                With DestDoc
                    ' 1
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
                    'oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    'oPara.Range.Text = "<REMOVE>"
                End With
            Catch ex As Exception
                Throw ex
            End Try
        End Sub
        Private Sub ProcessComponents(ByVal bolVars As InspPrepGuidelineVars, ByRef WordApp As Word.Application, ByRef DestDoc As Word.Document)
            Dim oPara As Word.Paragraph
            Try
                With DestDoc
                    ' 15
                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                    oPara.Range.Text = "The tank fill ports (ensure that keys are available for any locks that may be on the fill port caps)"
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

        ' CheckList
        Public Function CreateInspCheckList(ByVal facID As Integer, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, ByVal oInsp As MUSTER.BusinessLogic.pInspection, ByVal dsTank As DataSet, ByVal dsPipe As DataSet, ByVal dsTerm As DataSet, ByVal progress As Integer, Optional ByRef WordApp As Word.Application = Nothing)
            Try
                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty
                Dim oAddressInfo As MUSTER.Info.AddressInfo
                Dim strAddress As String
                Dim strDate, strTime, strTimes() As String
                Dim i, j, rowIndex, colIndex, colCount As Integer
                Dim dr, drows() As DataRow
                Dim dt, dtSub As DataTable
                Dim dv, dvSub As DataView
                Dim ds As DataSet

                Dim oPara As Word.Paragraph
                Dim oTable As Word.Table

                'Instantiate the Word Object
                If IsNothing(WordApp) Then
                    WordApp = GetWordApp()
                End If

                If Not System.IO.File.Exists(TemplatePath) Then
                    Throw New Exception("File Not Found: " + TemplatePath)
                End If
                System.IO.File.Copy(TemplatePath, DestinationPath)

                If System.IO.File.Exists(DestinationPath) Then
                    With WordApp

                        DestDoc = .Documents.Open(DestinationPath)
                        DestDoc = WordApp.ActiveDocument

                        With DestDoc
                            .Activate()

                            ' Find and Replace the TAGs with Values.
                            For i = 0 To colParams.Count - 1
                                strKey = colParams.Keys(i).ToString
                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                            Next

                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                            oPara.Range.Text = "<Space>"
                            oPara.Format.SpaceAfter = 1
                            oPara.Range.InsertParagraphAfter()

                            progress += 10
                            RaiseEvent CheckListProgress(progress)

                            ' Add Tanks
                            InspectionAddTanks(dsTank, WordApp, DestDoc)
                            progress += 10
                            RaiseEvent CheckListProgress(progress)

                            ' Add Pipes
                            InspectionAddPipes(dsPipe, WordApp, DestDoc)
                            progress += 10
                            RaiseEvent CheckListProgress(progress)

                            ' Add Terminations
                            InspectionAddTerms(dsTerm, WordApp, DestDoc)
                            progress += 10
                            RaiseEvent CheckListProgress(progress)

                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                            oPara.Range.Text = "<Space>"
                            oPara.Format.SpaceAfter = 1
                            oPara.Range.InsertParagraphAfter()

                            ' registration / testing / construction
                            dt = oInsp.CheckListMaster.RegTable()
                            dv = dt.DefaultView
                            dv.Sort = "CL_POSITION"

                            If dt.Rows.Count > 0 Then
                                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dt.Rows.Count, 4)
                                oTable.Range.Font.Name = "Arial"
                                oTable.Range.Font.Size = 8
                                oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(5.8)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                'oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)

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
                                            'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                                        Else
                                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                            oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                            oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                            'oTable.Cell(i + 1, 5).Range.Text = dv.Item(i)("CCAT")
                                        End If
                                    End With
                                Next

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
                                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dt.Rows.Count, 4)
                                oTable.Range.Font.Name = "Arial"
                                oTable.Range.Font.Size = 8
                                oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(5.8)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                'oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)

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
                                            'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                                        Else
                                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                            oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                            oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                            'oTable.Cell(i + 1, 5).Range.Text = dv.Item(i)("CCAT")
                                        End If
                                    End With
                                Next
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
                                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, drows.Length, 4)
                                oTable.Range.Font.Name = "Arial"
                                oTable.Range.Font.Size = 8
                                oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(5.8)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                'oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)

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
                                            'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                                        Else
                                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                            oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                            oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                            'oTable.Cell(i + 1, 5).Range.Text = dv.Item(i)("CCAT")
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
                                    oTable.Range.Font.Name = "Arial"
                                    oTable.Range.Font.Size = 8
                                    oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                    oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                    oTable.Columns.Item(2).Width = WordApp.InchesToPoints(1.7)
                                    oTable.Columns.Item(3).Width = WordApp.InchesToPoints(1.7)
                                    oTable.Columns.Item(4).Width = WordApp.InchesToPoints(1.7)
                                    oTable.Columns.Item(5).Width = WordApp.InchesToPoints(1.7)

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
                                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, drows.Length, 4)
                                oTable.Range.Font.Name = "Arial"
                                oTable.Range.Font.Size = 8
                                oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(5.8)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                'oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)

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
                                            Else
                                                ' add cp readings cp tank/pipe/term tested by inspector
                                                dtSub = ds.Tables("CPTankInspectorTested")
                                                If dtSub.Rows.Count > 0 Then
                                                    If dtSub.Rows(0)("Yes") Then
                                                        oTable.Cell(i + 1, 3).Range.Text = "X"
                                                        oTable.Cell(i + 1, 4).Range.Text = ""
                                                    ElseIf dtSub.Rows(0)("No") Then
                                                        oTable.Cell(i + 1, 3).Range.Text = ""
                                                        oTable.Cell(i + 1, 4).Range.Text = "X"
                                                    End If
                                                End If
                                            End If
                                            'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                                        Else
                                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                            oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                            oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                            'oTable.Cell(i + 1, 5).Range.Text = dv.Item(i)("CCAT")
                                        End If
                                    End With
                                Next

                                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                oPara.Range.Text = "<Space>"
                                oPara.Format.SpaceAfter = 1
                                oPara.Range.InsertParagraphAfter()

                                ' add cp readings galvanic / impressed current
                                'dtSub = ds.Tables("CPTankGalvanic")
                                'If dtSub.Rows.Count > 0 Then
                                '    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, 1, 1)
                                '    oTable.Range.Font.Name = "Arial"
                                '    oTable.Range.Font.Size = 8
                                '    oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                '    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                '    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                '    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                '    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                '    oTable.Columns.Item(1).Width = WordApp.InchesToPoints(7.5)

                                '    oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                '    If dtSub.Rows(0)("Galvanic") Then
                                '        oTable.Cell(1, 1).Range.Text = "GALVANIC: _X_  IMPRESSED CURRENT: ___"
                                '    ElseIf dtSub.Rows(0)("Impressed Current") Then
                                '        oTable.Cell(1, 1).Range.Text = "GALVANIC: ___  IMPRESSED CURRENT: _X_"
                                '    Else
                                '        oTable.Cell(1, 1).Range.Text = "GALVANIC: ___  IMPRESSED CURRENT: ___"
                                '    End If
                                '    With oTable.Rows.Item(1).Shading
                                '        .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                '        .Texture = Word.WdTextureIndex.wdTexture30Percent
                                '    End With

                                '    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                '    oPara.Range.Text = "<Space>"
                                '    oPara.Format.SpaceAfter = 1
                                '    oPara.Range.InsertParagraphAfter()
                                'End If

                                ' add cp readings description of remote reference cell placement row
                                dtSub = ds.Tables("CPTankRemote")
                                dvSub = dtSub.DefaultView
                                If dtSub.Rows.Count > 0 Then
                                    ' CP READINGS DESCRIPTION
                                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dtSub.Rows.Count + 1, 1)
                                    oTable.Range.Font.Name = "Arial"
                                    oTable.Range.Font.Size = 8
                                    oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                    oTable.Columns.Item(1).Width = WordApp.InchesToPoints(7.5)

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
                                    oTable.Range.Font.Name = "Arial"
                                    oTable.Range.Font.Size = 8
                                    oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                    oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                    oTable.Columns.Item(2).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.8)
                                    oTable.Columns.Item(4).Width = WordApp.InchesToPoints(1.0)
                                    oTable.Columns.Item(5).Width = WordApp.InchesToPoints(1.0)
                                    oTable.Columns.Item(6).Width = WordApp.InchesToPoints(1.0)
                                    oTable.Columns.Item(7).Width = WordApp.InchesToPoints(1.0)
                                    oTable.Columns.Item(8).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(9).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(10).Width = WordApp.InchesToPoints(0.5)

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
                                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, drows.Length, 4)
                                oTable.Range.Font.Name = "Arial"
                                oTable.Range.Font.Size = 8
                                oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(5.8)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                'oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)

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
                                            Else
                                                ' add cp readings cp tank/pipe/term tested by inspector
                                                dtSub = ds.Tables("CPPipeInspectorTested")
                                                If dtSub.Rows.Count > 0 Then
                                                    If dtSub.Rows(0)("Yes") Then
                                                        oTable.Cell(i + 1, 3).Range.Text = "X"
                                                        oTable.Cell(i + 1, 4).Range.Text = ""
                                                    ElseIf dtSub.Rows(0)("No") Then
                                                        oTable.Cell(i + 1, 3).Range.Text = ""
                                                        oTable.Cell(i + 1, 4).Range.Text = "X"
                                                    End If
                                                End If
                                            End If
                                            'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                                        Else
                                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                            oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                            oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                            'oTable.Cell(i + 1, 5).Range.Text = dv.Item(i)("CCAT")
                                        End If
                                    End With
                                Next

                                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                oPara.Range.Text = "<Space>"
                                oPara.Format.SpaceAfter = 1
                                oPara.Range.InsertParagraphAfter()

                                ' add cp readings galvanic / impressed current
                                'dtSub = ds.Tables("CPPipeGalvanic")
                                'If dtSub.Rows.Count > 0 Then
                                '    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, 1, 1)
                                '    oTable.Range.Font.Name = "Arial"
                                '    oTable.Range.Font.Size = 8
                                '    oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                '    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                '    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                '    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                '    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                '    oTable.Columns.Item(1).Width = WordApp.InchesToPoints(7.5)

                                '    oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                '    If dtSub.Rows(0)("Galvanic") Then
                                '        oTable.Cell(1, 1).Range.Text = "GALVANIC: _X_  IMPRESSED CURRENT: ___"
                                '    ElseIf dtSub.Rows(0)("Impressed Current") Then
                                '        oTable.Cell(1, 1).Range.Text = "GALVANIC: ___  IMPRESSED CURRENT: _X_"
                                '    Else
                                '        oTable.Cell(1, 1).Range.Text = "GALVANIC: ___  IMPRESSED CURRENT: ___"
                                '    End If
                                '    With oTable.Rows.Item(1).Shading
                                '        .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                '        .Texture = Word.WdTextureIndex.wdTexture30Percent
                                '    End With

                                '    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                '    oPara.Range.Text = "<Space>"
                                '    oPara.Format.SpaceAfter = 1
                                '    oPara.Range.InsertParagraphAfter()
                                'End If

                                ' add cp readings description of remote reference cell placement row
                                dtSub = ds.Tables("CPPipeRemote")
                                dvSub = dtSub.DefaultView
                                If dtSub.Rows.Count > 0 Then
                                    ' CP READINGS DESCRIPTION
                                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dtSub.Rows.Count + 1, 1)
                                    oTable.Range.Font.Name = "Arial"
                                    oTable.Range.Font.Size = 8
                                    oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                    oTable.Columns.Item(1).Width = WordApp.InchesToPoints(7.5)

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
                                    oTable.Range.Font.Name = "Arial"
                                    oTable.Range.Font.Size = 8
                                    oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                    oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                    oTable.Columns.Item(2).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.8)
                                    oTable.Columns.Item(4).Width = WordApp.InchesToPoints(1.0)
                                    oTable.Columns.Item(5).Width = WordApp.InchesToPoints(1.0)
                                    oTable.Columns.Item(6).Width = WordApp.InchesToPoints(1.0)
                                    oTable.Columns.Item(7).Width = WordApp.InchesToPoints(1.0)
                                    oTable.Columns.Item(8).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(9).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(10).Width = WordApp.InchesToPoints(0.5)

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
                                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, drows.Length, 4)
                                oTable.Range.Font.Name = "Arial"
                                oTable.Range.Font.Size = 8
                                oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(5.8)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                'oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)

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
                                            Else
                                                ' add cp readings cp tank/pipe/term tested by inspector
                                                dtSub = ds.Tables("CPTermInspectorTested")
                                                If dtSub.Rows.Count > 0 Then
                                                    If dtSub.Rows(0)("Yes") Then
                                                        oTable.Cell(i + 1, 3).Range.Text = "X"
                                                        oTable.Cell(i + 1, 4).Range.Text = ""
                                                    ElseIf dtSub.Rows(0)("No") Then
                                                        oTable.Cell(i + 1, 3).Range.Text = ""
                                                        oTable.Cell(i + 1, 4).Range.Text = "X"
                                                    End If
                                                End If
                                            End If
                                            'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                                        Else
                                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                            oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                            oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                            'oTable.Cell(i + 1, 5).Range.Text = dv.Item(i)("CCAT")
                                        End If
                                    End With
                                Next

                                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                oPara.Range.Text = "<Space>"
                                oPara.Format.SpaceAfter = 1
                                oPara.Range.InsertParagraphAfter()

                                ' add cp readings galvanic / impressed current
                                'dtSub = ds.Tables("CPTermGalvanic")
                                'If dtSub.Rows.Count > 0 Then
                                '    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, 1, 1)
                                '    oTable.Range.Font.Name = "Arial"
                                '    oTable.Range.Font.Size = 8
                                '    oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                '    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                '    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                '    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                '    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                '    oTable.Columns.Item(1).Width = WordApp.InchesToPoints(7.5)

                                '    oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                '    If dtSub.Rows(0)("Galvanic") Then
                                '        oTable.Cell(1, 1).Range.Text = "GALVANIC: _X_  IMPRESSED CURRENT: ___"
                                '    ElseIf dtSub.Rows(0)("Impressed Current") Then
                                '        oTable.Cell(1, 1).Range.Text = "GALVANIC: ___  IMPRESSED CURRENT: _X_"
                                '    Else
                                '        oTable.Cell(1, 1).Range.Text = "GALVANIC: ___  IMPRESSED CURRENT: ___"
                                '    End If
                                '    With oTable.Rows.Item(1).Shading
                                '        .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                '        .Texture = Word.WdTextureIndex.wdTexture30Percent
                                '    End With

                                '    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                '    oPara.Range.Text = "<Space>"
                                '    oPara.Format.SpaceAfter = 1
                                '    oPara.Range.InsertParagraphAfter()
                                'End If

                                ' add cp readings description of remote reference cell placement row
                                dtSub = ds.Tables("CPTermRemote")
                                dvSub = dtSub.DefaultView
                                If dtSub.Rows.Count > 0 Then
                                    ' CP READINGS DESCRIPTION
                                    oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dtSub.Rows.Count + 1, 1)
                                    oTable.Range.Font.Name = "Arial"
                                    oTable.Range.Font.Size = 8
                                    oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                    oTable.Columns.Item(1).Width = WordApp.InchesToPoints(7.5)

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
                                    oTable.Range.Font.Name = "Arial"
                                    oTable.Range.Font.Size = 8
                                    oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                    oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                    oTable.Columns.Item(2).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.8)
                                    oTable.Columns.Item(4).Width = WordApp.InchesToPoints(1.0)
                                    oTable.Columns.Item(5).Width = WordApp.InchesToPoints(1.0)
                                    oTable.Columns.Item(6).Width = WordApp.InchesToPoints(1.0)
                                    oTable.Columns.Item(7).Width = WordApp.InchesToPoints(1.0)
                                    oTable.Columns.Item(8).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(9).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(10).Width = WordApp.InchesToPoints(0.5)

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
                                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, drows.Length, 4)
                                oTable.Range.Font.Name = "Arial"
                                oTable.Range.Font.Size = 8
                                oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(5.8)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                'oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)

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
                                            End If
                                            'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                                        Else
                                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                            oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                            oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
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
                                    oTable.Range.Font.Name = "Arial"
                                    oTable.Range.Font.Size = 8
                                    oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                    oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                    oTable.Columns.Item(2).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(6).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(7).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(8).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(9).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(10).Width = WordApp.InchesToPoints(2.8)

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
                                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, drows.Length, 4)
                                oTable.Range.Font.Name = "Arial"
                                oTable.Range.Font.Size = 8
                                oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(5.8)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                'oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)

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
                                            'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                                        Else
                                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                            oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                            oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                            'oTable.Cell(i + 1, 5).Range.Text = dv.Item(i)("CCAT")
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
                                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, drows.Length, 4)
                                oTable.Range.Font.Name = "Arial"
                                oTable.Range.Font.Size = 8
                                oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(5.8)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                'oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)

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
                                            End If
                                            'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                                        Else
                                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                            oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                            oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
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
                                    oTable.Range.Font.Name = "Arial"
                                    oTable.Range.Font.Size = 8
                                    oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                    oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                    oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                    oTable.Columns.Item(2).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(6).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(7).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(8).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(9).Width = WordApp.InchesToPoints(0.5)
                                    oTable.Columns.Item(10).Width = WordApp.InchesToPoints(2.8)

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
                                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, drows.Length, 4)
                                oTable.Range.Font.Name = "Arial"
                                oTable.Range.Font.Size = 8
                                oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(5.8)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                'oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)

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
                                            'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                                        Else
                                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                            oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                            oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                            'oTable.Cell(i + 1, 5).Range.Text = dv.Item(i)("CCAT")
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
                                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dt.Rows.Count, 4)
                                oTable.Range.Font.Name = "Arial"
                                oTable.Range.Font.Size = 8
                                oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(5.8)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                'oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)

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
                                            'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                                        Else
                                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                            oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                            oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                            'oTable.Cell(i + 1, 5).Range.Text = dv.Item(i)("CCAT")
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
                                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dt.Rows.Count, 4)
                                oTable.Range.Font.Name = "Arial"
                                oTable.Range.Font.Size = 8
                                oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(5.8)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                'oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)

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
                                            'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                                        Else
                                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                            oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                            oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                            'oTable.Cell(i + 1, 5).Range.Text = dv.Item(i)("CCAT")
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
                                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, dt.Rows.Count, 4)
                                oTable.Range.Font.Name = "Arial"
                                oTable.Range.Font.Size = 8
                                oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(5.8)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                'oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)

                                For i = 0 To dt.Rows.Count - 1
                                    oTable.Rows.Item(i + 1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                    With oTable.Rows.Item(i + 1).Shading
                                        'oTable.Cell(i + 1, 1).Range.Text = dv.Item(i)("Line#")
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
                                            'oTable.Cell(i + 1, 5).Range.Text = "CCAT"
                                        Else
                                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                            oTable.Cell(i + 1, 3).Range.Text = IIf(dv.Item(i)("Yes"), "X", "")
                                            oTable.Cell(i + 1, 4).Range.Text = IIf(dv.Item(i)("No"), "X", "")
                                            'oTable.Cell(i + 1, 5).Range.Text = dv.Item(i)("CCAT")
                                        End If
                                    End With
                                Next

                                oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                oPara.Range.Text = "<Space>"
                                oPara.Format.SpaceAfter = 1
                                oPara.Range.InsertParagraphAfter()
                            End If

                            ' page break
                            WordApp.Selection.GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:=.Bookmarks.Item("\endofdoc").Name)
                            WordApp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)

                            ' Monitor Wells
                            ds = oInsp.CheckListMaster.MWellTable
                            dt = ds.Tables("TankPipeMW")
                            If dt.Rows.Count > 0 Then
                                dv = dt.DefaultView
                                dv.Sort = "CL_POSITION"

                                oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, 1, 4)
                                oTable.Range.Font.Name = "Arial"
                                oTable.Range.Font.Size = 8
                                oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(5.8)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                'oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)

                                ' Header Line
                                oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                With oTable.Rows.Item(1).Shading
                                    'oTable.Cell(1, 1).Range.Text = dv.Item(0)("Line#")
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
                                oTable.Range.Font.Name = "Arial"
                                oTable.Range.Font.Size = 8
                                oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                                oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                                oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                                oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(6).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(7).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(8).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(9).Width = WordApp.InchesToPoints(0.5)
                                oTable.Columns.Item(10).Width = WordApp.InchesToPoints(2.8)

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




                            ' Inspection Comments
                            'Dim strComments As String = oInsp.CheckListMaster.InspectionComments.InsComments
                            Dim nRows As Integer = 0
                            If oInsp.CheckListMaster.InspectionComments.InsComments = String.Empty Then
                                'nRows = 15
                                nRows = 38
                            Else
                                nRows = 2
                            End If

                            ' page break
                            WordApp.Selection.GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:=.Bookmarks.Item("\endofdoc").Name)
                            WordApp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)

                            oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, nRows, 1)
                            oTable.Range.Font.Name = "Arial"
                            oTable.Range.Font.Size = 8
                            oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                            oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                            oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                            oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                            oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                            oTable.Columns.Item(1).Width = WordApp.InchesToPoints(7.5)
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












                            ' Inspection CCAT
                            'Dim strComments As String = oInsp.CheckListMaster.InspectionComments.InsComments
                            nRows = 0
                            If oInsp.CheckListMaster.InspectionInfo.CCATsCollection.Count > 1 Then
                                'nRows = 15
                                nRows = 38
                            Else
                                nRows = 2
                            End If

                            ' page break
                            WordApp.Selection.GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:=.Bookmarks.Item("\endofdoc").Name)
                            WordApp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)

                            oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, nRows, 1)
                            oTable.Range.Font.Name = "Arial"
                            oTable.Range.Font.Size = 8
                            oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                            oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                            oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                            oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                            oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                            oTable.Columns.Item(1).Width = WordApp.InchesToPoints(7.5)
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

                            ' content
                            Dim g As Integer = 0
                            While g <= oInsp.CheckListMaster.InspectionInfo.CitationsCollection.Count - 1

                                Dim key As String = oInsp.CheckListMaster.InspectionInfo.CitationsCollection.GetKeys(g)


                                With oTable.Rows.Item(2 + g).Shading
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                End With

                                oTable.Rows.Item(2 + g).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                If oInsp.CheckListMaster.InspectionInfo.CitationsCollection.Item(key).CCAT <> String.Empty Then

                                    Dim CCAT As String = oInsp.CheckListMaster.InspectionInfo.CitationsCollection(key).CCAT

                                    For kk As Integer = 0 To dv.Count - 1

                                        If dv.Item(kk)("QUESTION_ID") = oInsp.CheckListMaster.InspectionInfo.CitationsCollection(key).QuestionID Then

                                            oTable.Rows.Item(2 + g).Range.Text = String.Format("Citation {0}:   {1}", dv.Item(kk)("Line#"), CCAT)

                                            Exit For

                                        End If

                                    Next

                                    oTable.Rows.Item(2 + g).Range.Font.Bold = False

                                End If

                                g = g + 1
                            End While


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












                            ' page break
                            WordApp.Selection.GoTo(What:=Word.WdGoToItem.wdGoToBookmark, Name:=.Bookmarks.Item("\endofdoc").Name)
                            WordApp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage)

                            oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, IIf(dt.Rows.Count + 2 > 33, dt.Rows.Count + 2, 33), 2)
                            oTable.Range.Font.Name = "Arial"
                            oTable.Range.Font.Size = 8
                            oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                            oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                            oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                            oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                            oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                            oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                            oTable.Columns.Item(2).Width = WordApp.InchesToPoints(6.8)

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

                            For i = 0 To dt.Rows.Count - 1


                                oTable.Rows.Item(i + 3).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                                With oTable.Rows.Item(i + 3).Shading
                                    .BackgroundPatternColor = Word.WdColor.wdColorWhite

                                    oTable.Cell(i + 3, 1).Range.Text = dv.Item(i)("Line#")
                                    oTable.Cell(i + 3, 2).Range.Text = dv.Item(i)("Description")


                                End With
                            Next

                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                            oPara.Range.Font.Name = "Arial"
                            oPara.Range.Font.Size = 8
                            oPara.Range.Text = "BY YOUR SIGNATURE, IT IS ACKNOWLEDGED THAT YOU UNDERSTAND AND" + vbCrLf + _
                                               "AGREE WITH ANY DISCREPANCIES THAT MAY BE LISTED ABOVE" + vbCrLf + vbCrLf + _
                                               "OWNER/OWNER'S REPRESENTATIVE SIGNATURE ________________________________________" + vbCrLf + vbCrLf + _
                                               "DATE ____________________"
                            oPara.Range.InsertParagraphAfter()

                            ' SOC

                            ' delete spacing between tables
                            oPara.Range.Text = "<Space>"
                            For Each para As Word.Paragraph In .Content.Paragraphs
                                If para.Range.Text = oPara.Range.Text Then
                                    para.Range.Delete()
                                End If
                            Next

                            .Save()

                        End With
                        DestDoc = Nothing
                        '.ActiveDocument.Close(False)
                    End With
                    RaiseEvent CheckListProgress(100)
                Else
                    Throw New Exception("Unable to copy template " & TemplatePath & " to " & DestinationPath & " in pLetterGen object.")
                End If
            Catch ex As Exception
                SrcDoc = Nothing
                If Not WordApp Is Nothing Then
                    If Not WordApp.ActiveDocument Is Nothing Then WordApp.ActiveDocument.Close(False)
                End If
                System.IO.File.Delete(DestinationPath)
                Throw ex
            End Try
        End Function
        Private Sub InspectionAddTanks(ByVal ds As DataSet, ByRef WordApp As Word.Application, ByRef DestDoc As Word.Document)
            Dim dt As DataTable
            Dim i, colCount1, colCount2, colIndex As Integer
            Dim oTable As Word.Table
            Dim oPara As Word.Paragraph
            Try
                With DestDoc
                    ' Add Tank Table
                    For i = 0 To ds.Tables.Count - 1
                        dt = ds.Tables(i)
                        ' First Row
                        ' to determine how many columns to display in first row of a given tank
                        colCount1 = 10
                        If Not dt.Columns.Contains("CP Type") Then
                            colCount1 -= 1
                        End If

                        oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, 3, colCount1)
                        oTable.Range.Font.Name = "Arial"
                        oTable.Range.Font.Size = 8
                        oTable.Range.Font.Bold = True
                        oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True

                        oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                        oTable.Columns.Item(2).Width = WordApp.InchesToPoints(0.7)

                        If colCount1 = 10 Then
                            oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.76)
                            oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.76)
                            oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.76)
                            oTable.Columns.Item(6).Width = WordApp.InchesToPoints(0.76)
                            oTable.Columns.Item(7).Width = WordApp.InchesToPoints(0.76)
                            oTable.Columns.Item(8).Width = WordApp.InchesToPoints(0.76)
                            oTable.Columns.Item(9).Width = WordApp.InchesToPoints(0.76)
                            oTable.Columns.Item(10).Width = WordApp.InchesToPoints(0.76)
                        Else
                            oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.87)
                            oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.87)
                            oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.87)
                            oTable.Columns.Item(6).Width = WordApp.InchesToPoints(0.87)
                            oTable.Columns.Item(7).Width = WordApp.InchesToPoints(0.87)
                            oTable.Columns.Item(8).Width = WordApp.InchesToPoints(0.87)
                            oTable.Columns.Item(9).Width = WordApp.InchesToPoints(0.87)
                        End If

                        ' row 1
                        oTable.Rows.Item(1).Cells.Merge()
                        oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        oTable.Rows.Item(1).Range.Text = "TANKS"
                        With oTable.Rows.Item(1).Shading
                            '.BackgroundPatternColor = Word.WdColor.wdColorBlack
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture30Percent
                        End With

                        ' row 2
                        oTable.Rows.Item(2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(2).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture30Percent
                        End With

                        ' row 3
                        oTable.Rows.Item(3).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(3).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                        End With

                        ' Fill Data
                        For colIndex = 1 To colCount1
                            ' row 2 - header
                            oTable.Rows.Item(2).Cells().Item(colIndex).Range.Text = dt.Columns(colIndex - 1).ColumnName

                            ' row 3 - data
                            oTable.Rows.Item(3).Cells().Item(colIndex).Range.Text = dt.Rows(0)(colIndex - 1)
                        Next

                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "<Space>"
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()

                        ' Second row
                        ' to determine how many cells to display in second row of a given tank
                        colCount2 = 8
                        If Not dt.Columns.Contains("Leak Detection") Then
                            colCount2 -= 1
                        End If
                        If Not dt.Columns.Contains("Overfill Type") Then
                            colCount2 -= 1
                        End If
                        If Not dt.Columns.Contains("Lined") Then
                            colCount2 -= 1
                        End If
                        If Not dt.Columns.Contains("Lining Inspected") Then
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

                        If colCount2 > 1 Then
                            oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, 3, colCount2)
                            oTable.Range.Font.Name = "Arial"
                            oTable.Range.Font.Size = 8
                            oTable.Range.Font.Bold = True
                            oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                            oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                            oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                            oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                            oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True


                            oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)

                            If colCount2 = 2 Then
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(6.8)
                            ElseIf colCount2 = 3 Then
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(3.4)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(3.4)
                            ElseIf colCount2 = 4 Then
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(3)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(3)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.8)
                            ElseIf colCount2 = 5 Then
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(2.6)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(2.6)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.8)
                                oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.8)
                            ElseIf colCount2 = 6 Then
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(2.2)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(2.2)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.8)
                                oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.8)
                                oTable.Columns.Item(6).Width = WordApp.InchesToPoints(0.8)
                            ElseIf colCount2 = 7 Then
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(1.8)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(1.8)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.8)
                                oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.8)
                                oTable.Columns.Item(6).Width = WordApp.InchesToPoints(0.8)
                                oTable.Columns.Item(7).Width = WordApp.InchesToPoints(0.8)
                            Else
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(1.4)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(1.4)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.8)
                                oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.8)
                                oTable.Columns.Item(6).Width = WordApp.InchesToPoints(0.8)
                                oTable.Columns.Item(7).Width = WordApp.InchesToPoints(0.8)
                                oTable.Columns.Item(8).Width = WordApp.InchesToPoints(0.8)
                            End If

                            ' row 1
                            oTable.Rows.Item(1).Cells.Merge()
                            oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            oTable.Rows.Item(1).Range.Text = "TANKS"
                            With oTable.Rows.Item(1).Shading
                                '.BackgroundPatternColor = Word.WdColor.wdColorBlack
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                .Texture = Word.WdTextureIndex.wdTexture30Percent
                            End With

                            ' row 2
                            oTable.Rows.Item(2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oTable.Rows.Item(2).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                .Texture = Word.WdTextureIndex.wdTexture30Percent
                            End With

                            ' row 3
                            oTable.Rows.Item(3).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oTable.Rows.Item(3).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            End With

                            ' Fill Data

                            ' row 2 - header
                            oTable.Rows.Item(2).Cells().Item(1).Range.Text = dt.Columns(0).ColumnName
                            ' row 3 - data
                            oTable.Rows.Item(3).Cells().Item(1).Range.Text = dt.Rows(0)(0)

                            For colIndex = 2 To colCount2
                                ' row 2 - header
                                oTable.Rows.Item(2).Cells().Item(colIndex).Range.Text = dt.Columns(colCount1 + colIndex - 2).ColumnName

                                ' row 3 - data
                                oTable.Rows.Item(3).Cells().Item(colIndex).Range.Text = dt.Rows(0)(colCount1 + colIndex - 2)
                            Next

                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                            oPara.Range.Text = "<Space>"
                            oPara.Format.SpaceAfter = 1
                            oPara.Range.InsertParagraphAfter()
                        End If
                    Next
                End With
            Catch ex As Exception
                Throw ex
            End Try
        End Sub
        Private Sub InspectionAddPipes(ByVal ds As DataSet, ByRef WordApp As Word.Application, ByRef DestDoc As Word.Document)
            Dim dt As DataTable
            Dim i, colCount1, colCount2, colIndex As Integer
            Dim oTable As Word.Table
            Dim oPara As Word.Paragraph
            Try
                With DestDoc
                    ' Add Pipe Table
                    For i = 0 To ds.Tables.Count - 1
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

                        oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, 3, colCount1)
                        oTable.Range.Font.Name = "Arial"
                        oTable.Range.Font.Size = 8
                        oTable.Range.Font.Bold = True
                        oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True


                        oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                        oTable.Columns.Item(2).Width = WordApp.InchesToPoints(0.7)
                        oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.8)

                        If colCount1 = 6 Then
                            oTable.Columns.Item(4).Width = WordApp.InchesToPoints(1.76)
                            oTable.Columns.Item(5).Width = WordApp.InchesToPoints(1.76)
                            oTable.Columns.Item(6).Width = WordApp.InchesToPoints(1.76)
                        ElseIf colCount1 = 7 Then
                            oTable.Columns.Item(4).Width = WordApp.InchesToPoints(1.325)
                            oTable.Columns.Item(5).Width = WordApp.InchesToPoints(1.325)
                            oTable.Columns.Item(6).Width = WordApp.InchesToPoints(1.325)
                            oTable.Columns.Item(7).Width = WordApp.InchesToPoints(1.325)
                        Else
                            oTable.Columns.Item(4).Width = WordApp.InchesToPoints(1.06)
                            oTable.Columns.Item(5).Width = WordApp.InchesToPoints(1.06)
                            oTable.Columns.Item(6).Width = WordApp.InchesToPoints(1.06)
                            oTable.Columns.Item(7).Width = WordApp.InchesToPoints(1.06)
                            oTable.Columns.Item(8).Width = WordApp.InchesToPoints(1.06)
                        End If

                        ' row 1
                        oTable.Rows.Item(1).Cells.Merge()
                        oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        oTable.Rows.Item(1).Range.Text = "PIPING"
                        With oTable.Rows.Item(1).Shading
                            '.BackgroundPatternColor = Word.WdColor.wdColorBlack
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture30Percent
                        End With

                        ' row 2
                        oTable.Rows.Item(2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(2).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture30Percent
                        End With

                        ' row 3
                        oTable.Rows.Item(3).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(3).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                        End With

                        ' Fill Data
                        For colIndex = 1 To colCount1
                            ' row 2 - header
                            oTable.Rows.Item(2).Cells().Item(colIndex).Range.Text = dt.Columns(colIndex - 1).ColumnName

                            ' row 3 - data
                            oTable.Rows.Item(3).Cells().Item(colIndex).Range.Text = dt.Rows(0)(colIndex - 1)
                        Next

                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "<Space>"
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()

                        ' Second row
                        ' to determine how many cells to display in second row of a given pipe
                        colCount2 = 7
                        If Not dt.Columns.Contains("Primary Pipe LD") Then
                            colCount2 -= 1
                        End If
                        If Not dt.Columns.Contains("Secondary Pipe LD") Then
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

                        If colCount2 > 1 Then
                            oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, 3, colCount2)
                            oTable.Range.Font.Name = "Arial"
                            oTable.Range.Font.Size = 8
                            oTable.Range.Font.Bold = True
                            oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                            oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                            oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                            oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                            oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True


                            oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)

                            If colCount2 = 2 Then
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(6.8)
                            ElseIf colCount2 = 3 Then
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(3.4)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(3.4)
                            ElseIf colCount2 = 4 Then
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(3)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(3)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.8)
                            ElseIf colCount2 = 5 Then
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(2.6)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(2.6)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.8)
                                oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.8)
                            ElseIf colCount2 = 6 Then
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(2.2)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(2.2)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.8)
                                oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.8)
                                oTable.Columns.Item(6).Width = WordApp.InchesToPoints(0.8)
                            Else
                                oTable.Columns.Item(2).Width = WordApp.InchesToPoints(1.8)
                                oTable.Columns.Item(3).Width = WordApp.InchesToPoints(1.8)
                                oTable.Columns.Item(4).Width = WordApp.InchesToPoints(0.8)
                                oTable.Columns.Item(5).Width = WordApp.InchesToPoints(0.8)
                                oTable.Columns.Item(6).Width = WordApp.InchesToPoints(0.8)
                                oTable.Columns.Item(7).Width = WordApp.InchesToPoints(0.8)
                            End If

                            ' row 1
                            oTable.Rows.Item(1).Cells.Merge()
                            oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            oTable.Rows.Item(1).Range.Text = "PIPING"
                            With oTable.Rows.Item(1).Shading
                                '.BackgroundPatternColor = Word.WdColor.wdColorBlack
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                .Texture = Word.WdTextureIndex.wdTexture30Percent
                            End With

                            ' row 2
                            oTable.Rows.Item(2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oTable.Rows.Item(2).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                                .Texture = Word.WdTextureIndex.wdTexture30Percent
                            End With

                            ' row 3
                            oTable.Rows.Item(3).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                            With oTable.Rows.Item(3).Shading
                                .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            End With

                            ' Fill Data

                            ' row 2 - header
                            oTable.Rows.Item(2).Cells().Item(1).Range.Text = dt.Columns(0).ColumnName
                            ' row 3 - data
                            oTable.Rows.Item(3).Cells().Item(1).Range.Text = dt.Rows(0)(0)

                            For colIndex = 2 To colCount2
                                ' row 2 - header
                                oTable.Rows.Item(2).Cells().Item(colIndex).Range.Text = dt.Columns(colCount1 + colIndex - 2).ColumnName

                                ' row 3 - data
                                oTable.Rows.Item(3).Cells().Item(colIndex).Range.Text = dt.Rows(0)(colCount1 + colIndex - 2)
                            Next

                            oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                            oPara.Range.Text = "<Space>"
                            oPara.Format.SpaceAfter = 1
                            oPara.Range.InsertParagraphAfter()
                        End If
                    Next
                End With
            Catch ex As Exception
                Throw ex
            End Try
        End Sub
        Private Sub InspectionAddTerms(ByVal ds As DataSet, ByRef WordApp As Word.Application, ByRef DestDoc As Word.Document)
            Dim dt As DataTable
            Dim i, colCount, colIndex As Integer
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

                        oTable = .Tables.Add(.Bookmarks.Item("\endofdoc").Range, 3, colCount)
                        oTable.Range.Font.Name = "Arial"
                        oTable.Range.Font.Size = 8
                        oTable.Range.Font.Bold = True
                        oTable.Rows.Height = WordApp.InchesToPoints(0.25)
                        oTable.Borders.Item(Word.WdBorderType.wdBorderBottom).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderTop).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Visible = True
                        oTable.Borders.Item(Word.WdBorderType.wdBorderVertical).Visible = True


                        oTable.Columns.Item(1).Width = WordApp.InchesToPoints(0.7)
                        oTable.Columns.Item(2).Width = WordApp.InchesToPoints(0.7)
                        oTable.Columns.Item(3).Width = WordApp.InchesToPoints(0.7)

                        If colCount = 5 Then
                            oTable.Columns.Item(4).Width = WordApp.InchesToPoints(2.7)
                            oTable.Columns.Item(5).Width = WordApp.InchesToPoints(2.7)
                        ElseIf colCount = 6 Then
                            oTable.Columns.Item(4).Width = WordApp.InchesToPoints(1.8)
                            oTable.Columns.Item(5).Width = WordApp.InchesToPoints(1.8)
                            oTable.Columns.Item(6).Width = WordApp.InchesToPoints(1.8)
                        ElseIf colCount = 7 Then
                            oTable.Columns.Item(4).Width = WordApp.InchesToPoints(1.7)
                            oTable.Columns.Item(5).Width = WordApp.InchesToPoints(1.7)
                            oTable.Columns.Item(6).Width = WordApp.InchesToPoints(1.0)
                            oTable.Columns.Item(7).Width = WordApp.InchesToPoints(1.0)
                        ElseIf colCount = 8 Then
                            oTable.Columns.Item(4).Width = WordApp.InchesToPoints(1.3)
                            oTable.Columns.Item(5).Width = WordApp.InchesToPoints(1.3)
                            oTable.Columns.Item(6).Width = WordApp.InchesToPoints(1.0)
                            oTable.Columns.Item(7).Width = WordApp.InchesToPoints(1.0)
                            oTable.Columns.Item(8).Width = WordApp.InchesToPoints(0.8)
                        End If

                        ' row 1
                        oTable.Rows.Item(1).Cells.Merge()
                        oTable.Rows.Item(1).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        oTable.Rows.Item(1).Range.Text = "PIPING TERMINATIONS"
                        With oTable.Rows.Item(1).Shading
                            '.BackgroundPatternColor = Word.WdColor.wdColorBlack
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture30Percent
                        End With

                        ' row 2
                        oTable.Rows.Item(2).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(2).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                            .Texture = Word.WdTextureIndex.wdTexture30Percent
                        End With

                        ' row 3
                        oTable.Rows.Item(3).Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter
                        With oTable.Rows.Item(3).Shading
                            .BackgroundPatternColor = Word.WdColor.wdColorWhite
                        End With

                        ' Fill Data
                        For colIndex = 1 To colCount
                            ' row 2 - header
                            oTable.Rows.Item(2).Cells().Item(colIndex).Range.Text = dt.Columns(colIndex - 1).ColumnName

                            ' row 3 - data
                            oTable.Rows.Item(3).Cells().Item(colIndex).Range.Text = dt.Rows(0)(colIndex - 1)
                        Next

                        oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                        oPara.Range.Text = "<Space>"
                        oPara.Format.SpaceAfter = 1
                        oPara.Range.InsertParagraphAfter()
                    Next
                End With
            Catch ex As Exception
                Throw ex
            End Try
        End Sub

        ' C & E
        Public Function GenerateCAELetter(ByVal colParams As Specialized.NameValueCollection, ByVal strFacsForAgreedOrder() As String, _
                ByVal TemplatePath As String, ByVal COCTemplatePath As String, ByVal COITemplatePath As String, ByVal DestinationPath As String, _
                ByVal dtFacs As DataTable, ByVal dtCitations As DataTable, ByVal dtDiscreps As DataTable, ByVal dtCOIFacs As DataTable, _
                ByVal alFacsIndex As ArrayList, ByVal alCitationsIndex As ArrayList, _
                ByVal alCorrectiveActionIndex As ArrayList, ByVal alCorrectiveActionWithDueDateIndex As ArrayList, _
                ByVal alCorrectiveActionAddOnIndex As ArrayList, ByVal alDiscrepsIndex As ArrayList, ByVal alDiscrepsCorrectiveActionIndex As ArrayList, _
                Optional ByVal bolCOCRequired As Boolean = False, Optional ByRef WordApp As Word.Application = Nothing)
            Dim strKey As String = String.Empty
            Dim strValue As String = String.Empty
            Dim prevFac, i, j, k, nStrKeyEndIndex, lineNumBegin, lineNumEnd As Integer
            Dim dv As DataView

            Try
                'Instantiate the Word Object
                If IsNothing(WordApp) Then
                    WordApp = GetWordApp()
                End If

                If Not System.IO.File.Exists(TemplatePath) Then
                    Throw New Exception("File Not Found: " + TemplatePath)
                End If
                System.IO.File.Copy(TemplatePath, DestinationPath)

                If System.IO.File.Exists(DestinationPath) Then
                    With WordApp

                        DestDoc = .Documents.Open(DestinationPath)
                        DestDoc = WordApp.ActiveDocument

                        .Visible = True

                        With DestDoc
                            .Activate()

                            If COITemplatePath <> String.Empty Then
                                If System.IO.File.Exists(COITemplatePath) Then
                                    If Not dtCOIFacs Is Nothing Then
                                        If dtCOIFacs.Rows.Count > 0 Then
                                            colParams.Add("<enclosed>", "Certification of Inspection")
                                            colParams.Add("<CommaEnclosed>", ", Certification of Inspection")
                                            'If Not colParams.Get("<CommaEnclosed>") Is Nothing Then
                                            'End If
                                        End If
                                    End If
                                End If
                            End If

                            ' Find and Replace the TAGs with Values.
                            For i = 0 To colParams.Count - 1
                                strKey = colParams.Keys(i).ToString
                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                            Next


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
                                            WordApp.Selection.TypeText(strFacsForAgreedOrder(i) + ", ")
                                        Else
                                            WordApp.Selection.TypeText(strFacsForAgreedOrder(i))
                                        End If
                                    Next
                                    'End If
                                End If
                            End If

                            Dim wrdGlobal As New Word.Global
                            '' 1. 2. 3.
                            'With wrdGlobal.ListGalleries.Item(Word.WdListGalleryType.wdNumberGallery).ListTemplates.Item(1).ListLevels.Item(1)
                            '    .NumberFormat = "%1."
                            '    .TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab
                            '    .NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic
                            '    .NumberPosition = WordApp.InchesToPoints(0.25)
                            '    .Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft
                            '    .TextPosition = WordApp.InchesToPoints(0.5)
                            '    .TabPosition = WordApp.InchesToPoints(0.5)
                            '    .ResetOnHigher = 0
                            '    .StartAt = 1
                            '    With .Font
                            '        .Bold = False
                            '    End With
                            'End With

                            ' a. b. c.
                            With wrdGlobal.ListGalleries.Item(Word.WdListGalleryType.wdNumberGallery).ListTemplates.Item(4).ListLevels.Item(1)
                                .NumberFormat = "%1."
                                .TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab
                                .NumberStyle = Word.WdListNumberStyle.wdListNumberStyleUppercaseLetter
                                .NumberPosition = WordApp.InchesToPoints(0.25)
                                .Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft
                                .TextPosition = WordApp.InchesToPoints(0.5)
                                .TabPosition = WordApp.InchesToPoints(0.5)
                                .ResetOnHigher = 0
                                .StartAt = 1

                                With .Font
                                    .Bold = False
                                End With
                            End With

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

                                        'fill rows
                                        'Dim hasInspectedOnCell As Boolean = False
                                        'If .Tables.Item(facIndex).Columns.Count > 3 Then
                                        '    hasInspectedOnCell = True
                                        'End If
                                        dv = dtFacs.DefaultView
                                        dv.Sort = "FACILITY_ID"
                                        For i = 0 To dtFacs.Rows.Count - 1
                                            With .Tables.Item(facIndex)
                                                .Cell(i + 2, 1).Range.Text = "Facility ID #" + dv.Item(i)("FACILITY_ID").ToString + ", " + _
                                                                            dv.Item(i)("FACILITY").ToString + ", " + _
                                                                            dv.Item(i)("ADDRESS").ToString
                                                '.Cell(i + 2, 2).Range.Text = dtFacs.Rows(i)("ADDRESS").ToString
                                                '.Cell(i + 2, 3).Range.Text = dtFacs.Rows(i)("FACILITY_ID").ToString
                                                'If hasInspectedOnCell Then
                                                '    .Cell(i + 2, 4).Range.Text = dtFacs.Rows(i)("INSPECTEDON").ToString
                                                'End If

                                                .Cell(i + 2, 1).Range.Font.Bold = 0
                                                '.Cell(i + 2, 2).Range.Font.Bold = 0
                                                '.Cell(i + 2, 3).Range.Font.Bold = 0
                                                'If hasInspectedOnCell Then
                                                '    .Cell(i + 2, 4).Range.Font.Bold = 0
                                                'End If
                                            End With
                                        Next ' For i = 0 To dtFacs.Rows.Count - 1
                                    Next ' For Each facIndex As Integer In dtFacsIndex.ToArray
                                End If ' If dtFacsIndex.Count > 0 Then
                            End If ' If Not dtFacs Is Nothing Then
                            .Content.Find.Execute(FindText:="<Facility>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)

                            If Not dtCitations Is Nothing Then
                                ' Citation Table
                                If alCitationsIndex.Count > 0 Then
                                    For Each citIndex As Integer In alCitationsIndex.ToArray
                                        prevFac = 0
                                        ' Fill Table with Text
                                        With .Tables.Item(citindex)
                                            If WordApp.Selection.Range.InStory(.Cell(1, 1).Range) Then

                                                If citindex > 0 Then
                                                    WordApp.Selection.TypeBackspace()
                                                    WordApp.Selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=1)
                                                End If

                                                WordApp.Selection.Document.Content.Tables.Item(citindex).Cell(1, 1).Range.Select()
                                            End If
                                            dv = dtCitations.DefaultView
                                            dv.Sort = "FACILITY_ID, CITATION_INDEX"
                                            For i = 0 To dtCitations.Rows.Count - 1
                                                Threading.Thread.Sleep(200)
                                                If prevFac <> dv.Item(i)("FACILITY_ID") Then
                                                    'If prevFac <> 0 Then
                                                    'WordApp.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                                                    'lineNumEnd = WordApp.Selection.Information(Word.WdInformation.wdFirstCharacterLineNumber)
                                                    'WordApp.Selection.HomeKey(Unit:=Word.WdUnits.wdLine, Extend:=Word.WdMovementType.wdExtend)
                                                    'If lineNumEnd - lineNumBegin - 1 > 0 Then
                                                    'WordApp.Selection.MoveUp(Unit:=Word.WdUnits.wdLine, Count:=(lineNumEnd - lineNumBegin - 1), Extend:=Word.WdMovementType.wdExtend)
                                                    'End If

                                                    'WordApp.Selection.Range.ListFormat.ApplyListTemplate(ListTemplate:=wrdGlobal.ListGalleries.Item(Word.WdListGalleryType.wdNumberGallery).ListTemplates.Item(4), _
                                                    'ContinuePreviousList:=False, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, _
                                                    'DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                                                    'WordApp.Selection.Font.Bold = False

                                                    'WordApp.Selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=1)
                                                    'WordApp.Selection.TypeParagraph()
                                                    ' End If
                                                    'lineNumBegin = WordApp.Selection.Information(Word.WdInformation.wdFirstCharacterLineNumber)
                                                    prevFac = dv.Item(i)("FACILITY_ID")
                                                    If dtFacs.Rows.Count > 1 Then

                                                        If i > 0 Then
                                                            WordApp.Selection.TypeBackspace()
                                                            WordApp.Selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=1)
                                                            WordApp.Selection.TypeText(Chr(13))

                                                        End If


                                                        WordApp.Selection.Range.InsertParagraph()

                                                        WordApp.Selection.TypeText("Facility ID #(" + dv.Item(i)("FACILITY_ID").ToString + ")")
                                                        WordApp.Selection.TypeParagraph()

                                                    Else
                                                        WordApp.Selection.Range.InsertParagraph()


                                                    End If
                                                    WordApp.Selection.Range.ListFormat.ApplyListTemplate(ListTemplate:=wrdGlobal.ListGalleries.Item(Word.WdListGalleryType.wdNumberGallery).ListTemplates.Item(4), _
                                                    ContinuePreviousList:=False, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, _
                                                    DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                                                    WordApp.Selection.Font.Bold = False

                                                    WordApp.Selection.Range.ListFormat.ListLevelNumber = 1

                                                    If WordApp.Selection.Range.ListFormat.ListValue <> 1 Then
                                                        WordApp.Selection.Range.ListFormat.ApplyListTemplate(ListTemplate:=wrdGlobal.ListGalleries.Item(Word.WdListGalleryType.wdNumberGallery).ListTemplates.Item(4), _
                                                        ContinuePreviousList:=False, ApplyTo:=Word.WdListApplyTo.wdListApplyToThisPointForward, _
                                                        DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                                                        WordApp.Selection.Font.Bold = False
                                                    End If

                                                Else



                                                    WordApp.Selection.Range.ListFormat.ApplyListTemplate(ListTemplate:=wrdGlobal.ListGalleries.Item(Word.WdListGalleryType.wdNumberGallery).ListTemplates.Item(4), _
                                                    ContinuePreviousList:=True, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, _
                                                    DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                                                    WordApp.Selection.Font.Bold = False

                                                    WordApp.Selection.Range.ListFormat.ListLevelNumber = 1

                                                End If







                                                WordApp.Selection.TypeText(dv.Item(i)("CITATIONTEXT").ToString + Chr(13))

                                                Dim CCAT As New Text.StringBuilder
                                                Dim CCATStr As String = String.Empty

                                                If Not TypeOf dtCitations.Rows(i).Item("CCAT_COMMENTS") Is DBNull Then
                                                    CCATStr = dtCitations.Rows(i).Item("CCAT_COMMENTS")
                                                Else
                                                    CCATStr = dtCitations.Rows(i).Item("CCAT")
                                                End If

                                                If CCATStr.Length > 0 Then

                                                    WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlineDemote()

                                                    For Each item As String In CCATStr.Trim.Split(","c)
                                                        item = item.Trim

                                                        If item.Length > 2 AndAlso item.StartsWith("PT") Then
                                                            CCAT.Append("Term #").Append(item.Substring(2)).Append(Chr(13))
                                                        ElseIf item.Length > 1 AndAlso item.StartsWith("P") Then
                                                            CCAT.Append("Pipe #").Append(item.Substring(1)).Append(Chr(13))
                                                        ElseIf item.Length > 1 AndAlso item.StartsWith("T") Then
                                                            CCAT.Append("Tank #").Append(item.Substring(1)).Append(Chr(13))
                                                        End If

                                                    Next

                                                    WordApp.Selection.Range.ListFormat.ListLevelNumber = 2

                                                    WordApp.Selection.TypeText(CCAT.ToString)
                                                    WordApp.Selection.Range.ListParagraphs.Item(WordApp.Selection.Range.ListParagraphs.Count).OutlinePromote()

                                                End If


                                                CCAT.Length = 0

                                            Next

                                        End With
                                    Next ' For Each citIndex As Integer In dtCitationsIndex.ToArray
                                    WordApp.Selection.TypeBackspace()
                                    WordApp.Selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=1)
                                    WordApp.Selection.GoToNext(Word.WdGoToItem.wdGoToTable)
                                    WordApp.Selection.Select()

                                End If ' If dtCitationsIndex.Count > 0 Then

                                ' Corrective Action Table
                                If alCorrectiveActionIndex.Count > 0 Then
                                    For Each citIndex As Integer In alCorrectiveActionIndex.ToArray
                                        prevFac = 0
                                        ' Fill Table with Text
                                        With .Tables.Item(citindex)
                                            If WordApp.Selection.Range.InStory(.Cell(1, 1).Range) Then
                                                WordApp.Selection.Document.Content.Tables.Item(citindex).Cell(1, 1).Range.Select()
                                            End If
                                            dv = dtCitations.DefaultView
                                            dv.Sort = "FACILITY_ID, CITATION_INDEX"
                                            For i = 0 To dtCitations.Rows.Count - 1
                                                Threading.Thread.Sleep(200)

                                                If dv.Item(i)("CorrectiveAction").ToString <> String.Empty Then
                                                    If prevFac <> dv.Item(i)("FACILITY_ID") Then

                                                        prevFac = dv.Item(i)("FACILITY_ID")
                                                        If dtFacs.Rows.Count > 1 Then

                                                            If i > 0 Then
                                                                WordApp.Selection.TypeBackspace()
                                                                WordApp.Selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=1)
                                                                WordApp.Selection.TypeText(Chr(13))

                                                            End If


                                                            WordApp.Selection.Range.InsertParagraph()

                                                            WordApp.Selection.TypeText("Facility ID #(" + dv.Item(i)("FACILITY_ID").ToString + ")")
                                                            WordApp.Selection.TypeParagraph()

                                                            WordApp.Selection.Range.ListFormat.ApplyListTemplate(ListTemplate:=wrdGlobal.ListGalleries.Item(Word.WdListGalleryType.wdNumberGallery).ListTemplates.Item(4), _
                                                                 ContinuePreviousList:=False, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, _
                                                                 DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)

                                                            WordApp.Selection.Font.Bold = False

                                                            WordApp.Selection.Range.ListFormat.ListLevelNumber = 1

                                                            If WordApp.Selection.Range.ListFormat.ListValue <> 1 Then
                                                                WordApp.Selection.Range.ListFormat.ApplyListTemplate(ListTemplate:=wrdGlobal.ListGalleries.Item(Word.WdListGalleryType.wdNumberGallery).ListTemplates.Item(4), _
                                                                ContinuePreviousList:=False, ApplyTo:=Word.WdListApplyTo.wdListApplyToThisPointForward, _
                                                                DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                                                                WordApp.Selection.Font.Bold = False
                                                            End If
                                                        End If
                                                    End If

                                                    WordApp.Selection.TypeText(dv.Item(i)("CorrectiveAction").ToString)
                                                    If i < dtCitations.Rows.Count - 1 Then WordApp.Selection.TypeParagraph()

                                                End If

                                            Next

                                        End With
                                    Next ' For Each citIndex As Integer In dtCitationsIndex.ToArray
                                    WordApp.Selection.TypeBackspace()
                                    WordApp.Selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=1)
                                    WordApp.Selection.GoToNext(Word.WdGoToItem.wdGoToTable)
                                    WordApp.Selection.Select()

                                End If ' If dtCitationsIndex.Count > 0 Then

                                ' Corrective Action With Due Date Table
                                If alCorrectiveActionWithDueDateIndex.Count > 0 Then
                                    For Each citIndex As Integer In alCorrectiveActionWithDueDateIndex.ToArray
                                        prevFac = 0
                                        ' Fill Table with Text
                                        With .Tables.Item(citindex)
                                            If WordApp.Selection.Range.InStory(.Cell(1, 1).Range) Then
                                                WordApp.Selection.Document.Content.Tables.Item(citindex).Cell(1, 1).Range.Select()
                                            End If
                                            dv = dtCitations.DefaultView
                                            dv.Sort = "FACILITY_ID, CITATION_INDEX"
                                            For i = 0 To dtCitations.Rows.Count - 1
                                                Threading.Thread.Sleep(100)

                                                If dv.Item(i)("CorrectiveAction").ToString <> String.Empty Then
                                                    If prevFac <> dv.Item(i)("FACILITY_ID") Then
                                                        If prevFac <> 0 Then

                                                            ' if there is a table below the corrective action in template to get extra text from
                                                            If alCorrectiveActionAddOnIndex.Contains(citindex + 1) Then
                                                                For j = 1 To WordApp.ActiveDocument.Tables.Item(citindex + 1).Rows.Count
                                                                    strKey = WordApp.ActiveDocument.Tables.Item(citIndex + 1).Cell(j, 1).Range.Text
                                                                    nStrKeyEndIndex = strKey.Length
                                                                    For k = strKey.Length - 1 To 0 Step -1
                                                                        If Char.IsWhiteSpace(strKey.Chars(k)) Or Char.GetUnicodeCategory(strKey.Chars(k)) = Globalization.UnicodeCategory.Control Then
                                                                            nStrKeyEndIndex -= 1
                                                                        Else
                                                                            Exit For
                                                                        End If
                                                                    Next
                                                                    strKey = strKey.Substring(0, nStrKeyEndIndex)
                                                                    WordApp.Selection.TypeText(strKey)
                                                                    WordApp.Selection.TypeParagraph()
                                                                Next
                                                            End If

                                                            WordApp.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                                                            lineNumEnd = WordApp.Selection.Information(Word.WdInformation.wdFirstCharacterLineNumber)
                                                            WordApp.Selection.HomeKey(Unit:=Word.WdUnits.wdLine, Extend:=Word.WdMovementType.wdExtend)
                                                            If lineNumEnd - lineNumBegin - 1 > 0 Then
                                                                WordApp.Selection.MoveUp(Unit:=Word.WdUnits.wdLine, Count:=(lineNumEnd - lineNumBegin - 1), Extend:=Word.WdMovementType.wdExtend)
                                                            End If

                                                            WordApp.Selection.Range.ListFormat.ApplyListTemplate(ListTemplate:=wrdGlobal.ListGalleries.Item(Word.WdListGalleryType.wdNumberGallery).ListTemplates.Item(4), _
                                                               ContinuePreviousList:=False, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, _
                                                              DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                                                            WordApp.Selection.Font.Bold = False

                                                            WordApp.Selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=1)
                                                            WordApp.Selection.TypeParagraph()
                                                        End If
                                                        lineNumBegin = WordApp.Selection.Information(Word.WdInformation.wdFirstCharacterLineNumber)
                                                        prevFac = dv.Item(i)("FACILITY_ID")
                                                        If dtFacs.Rows.Count > 1 Then
                                                            WordApp.Selection.TypeText("Facility ID #(" + dv.Item(i)("FACILITY_ID").ToString + ")")
                                                            WordApp.Selection.TypeParagraph()
                                                        End If
                                                    End If

                                                    WordApp.Selection.TypeText(dv.Item(i)("CorrectiveAction").ToString + " by " + dv.Item(i)("DUE").ToString)
                                                    If i < dtCitations.Rows.Count - 1 Then WordApp.Selection.TypeParagraph()

                                                    If i >= dtCitations.Rows.Count - 1 Then
                                                        'WordApp.Selection.TypeParagraph()
                                                        ' if there is a table below the corrective action in template to get extra text from
                                                        If alCorrectiveActionAddOnIndex.Contains(citindex + 1) Then
                                                            WordApp.Selection.TypeParagraph()
                                                            For j = 1 To WordApp.ActiveDocument.Tables.Item(citindex + 1).Rows.Count
                                                                strKey = WordApp.ActiveDocument.Tables.Item(citIndex + 1).Cell(j, 1).Range.Text
                                                                nStrKeyEndIndex = strKey.Length
                                                                For k = strKey.Length - 1 To 0 Step -1
                                                                    If Char.IsWhiteSpace(strKey.Chars(k)) Or Char.GetUnicodeCategory(strKey.Chars(k)) = Globalization.UnicodeCategory.Control Then
                                                                        nStrKeyEndIndex -= 1
                                                                    Else
                                                                        Exit For
                                                                    End If
                                                                Next
                                                                strKey = strKey.Substring(0, nStrKeyEndIndex)
                                                                WordApp.Selection.TypeText(strKey)
                                                                If j < WordApp.ActiveDocument.Tables.Item(citindex + 1).Rows.Count Then WordApp.Selection.TypeParagraph()
                                                            Next
                                                        End If

                                                        'WordApp.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=2)
                                                        WordApp.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                                                        lineNumEnd = WordApp.Selection.Information(Word.WdInformation.wdFirstCharacterLineNumber)
                                                        WordApp.Selection.HomeKey(Unit:=Word.WdUnits.wdLine, Extend:=Word.WdMovementType.wdExtend)
                                                        If lineNumEnd - lineNumBegin - 1 > 0 Then
                                                            WordApp.Selection.MoveUp(Unit:=Word.WdUnits.wdLine, Count:=(lineNumEnd - lineNumBegin - 1), Extend:=Word.WdMovementType.wdExtend)
                                                        End If

                                                        WordApp.Selection.Range.ListFormat.ApplyListTemplate(ListTemplate:=wrdGlobal.ListGalleries.Item(Word.WdListGalleryType.wdNumberGallery).ListTemplates.Item(4), _
                                                            ContinuePreviousList:=False, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, _
                                                            DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                                                        WordApp.Selection.Font.Bold = False

                                                        WordApp.Selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=1)
                                                        If alCorrectiveActionAddOnIndex.Contains(citindex + 1) Then
                                                            WordApp.Selection.TypeParagraph()
                                                        End If
                                                    End If
                                                End If
                                            Next
                                        End With
                                    Next ' For Each citIndex As Integer In dtCitationsIndex.ToArray
                                End If ' If dtCitationsIndex.Count > 0 Then

                            End If ' If Not dtCitations Is Nothing Then
                            .Content.Find.Execute(FindText:="<Citation>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)
                            .Content.Find.Execute(FindText:=".                            ", ReplaceWith:=Chr(127), Replace:=Word.WdReplace.wdReplaceAll)
                            .Content.Find.Execute(FindText:="<CorrectiveAction>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)
                            .Content.Find.Execute(FindText:="<CorrectiveActionWithDueDate>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)
                            .Content.Find.Execute(FindText:="<CorrectiveActionWithDueDate", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)

                            If Not dtDiscreps Is Nothing Then
                                ' Discrepancy Table
                                If alDiscrepsIndex.Count > 0 Then
                                    For Each discrepIndex As Integer In alDiscrepsIndex.ToArray
                                        prevFac = 0
                                        ' Fill Table with Text
                                        With .Tables.Item(discrepIndex)
                                            If WordApp.Selection.Range.InStory(.Cell(1, 1).Range) Then
                                                WordApp.Selection.Document.Content.Tables.Item(discrepIndex).Cell(1, 1).Range.Select()
                                            End If
                                            dv = dtDiscreps.DefaultView
                                            dv.Sort = "FACILITY_ID, DISCREP_INDEX"
                                            For i = 0 To dtDiscreps.Rows.Count - 1
                                                Threading.Thread.Sleep(100)

                                                If dv.Item(i)("DISCREP TEXT").ToString <> String.Empty Then
                                                    If prevFac <> dv.Item(i)("FACILITY_ID") Then
                                                        If prevFac <> 0 Then
                                                            WordApp.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                                                            lineNumEnd = WordApp.Selection.Information(Word.WdInformation.wdFirstCharacterLineNumber)
                                                            WordApp.Selection.HomeKey(Unit:=Word.WdUnits.wdLine, Extend:=Word.WdMovementType.wdExtend)
                                                            If lineNumEnd - lineNumBegin - 1 > 0 Then
                                                                WordApp.Selection.MoveUp(Unit:=Word.WdUnits.wdLine, Count:=(lineNumEnd - lineNumBegin - 1), Extend:=Word.WdMovementType.wdExtend)
                                                            End If

                                                            WordApp.Selection.Range.ListFormat.ApplyListTemplate(ListTemplate:=wrdGlobal.ListGalleries.Item(Word.WdListGalleryType.wdNumberGallery).ListTemplates.Item(4), _
                                                               ContinuePreviousList:=False, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, _
                                                              DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                                                            WordApp.Selection.Font.Bold = False

                                                            WordApp.Selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=1)
                                                            WordApp.Selection.TypeParagraph()
                                                        End If
                                                        lineNumBegin = WordApp.Selection.Information(Word.WdInformation.wdFirstCharacterLineNumber)
                                                        prevFac = dv.Item(i)("FACILITY_ID")
                                                        If dtFacs.Rows.Count > 1 Then
                                                            WordApp.Selection.TypeText("Facility ID #(" + dv.Item(i)("FACILITY_ID").ToString + ")")
                                                            WordApp.Selection.TypeParagraph()
                                                        End If
                                                    End If

                                                    WordApp.Selection.TypeText(dv.Item(i)("DISCREP TEXT").ToString)
                                                    If i < dtDiscreps.Rows.Count - 1 Then WordApp.Selection.TypeParagraph()

                                                    If i >= dtDiscreps.Rows.Count - 1 Then
                                                        'WordApp.Selection.TypeParagraph()
                                                        'WordApp.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=2)
                                                        'WordApp.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                                                        lineNumEnd = WordApp.Selection.Information(Word.WdInformation.wdFirstCharacterLineNumber)
                                                        WordApp.Selection.HomeKey(Unit:=Word.WdUnits.wdLine, Extend:=Word.WdMovementType.wdExtend)
                                                        If lineNumEnd - lineNumBegin - 1 > 0 Then
                                                            WordApp.Selection.MoveUp(Unit:=Word.WdUnits.wdLine, Count:=(lineNumEnd - lineNumBegin - 1), Extend:=Word.WdMovementType.wdExtend)
                                                        End If

                                                        WordApp.Selection.Range.ListFormat.ApplyListTemplate(ListTemplate:=wrdGlobal.ListGalleries.Item(Word.WdListGalleryType.wdNumberGallery).ListTemplates.Item(4), _
                                                            ContinuePreviousList:=False, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, _
                                                            DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                                                        WordApp.Selection.Font.Bold = False
                                                    End If
                                                End If
                                            Next
                                        End With
                                    Next ' For Each facIndex As Integer In dtDiscrepsIndex.ToArray
                                End If ' If dtDiscrepsIndex.Count > 0 Then

                                ' Discrep Corrective Action Table
                                If alDiscrepsCorrectiveActionIndex.Count > 0 Then
                                    For Each discrepCAIndex As Integer In alDiscrepsCorrectiveActionIndex.ToArray
                                        prevFac = 0
                                        ' Fill Table with Text
                                        With .Tables.Item(discrepCAIndex)
                                            If WordApp.Selection.Range.InStory(.Cell(1, 1).Range) Then
                                                WordApp.Selection.Document.Content.Tables.Item(discrepCAIndex).Cell(1, 1).Range.Select()
                                            End If
                                            dv = dtDiscreps.DefaultView
                                            dv.Sort = "FACILITY_ID, DISCREP_INDEX"
                                            For i = 0 To dtDiscreps.Rows.Count - 1
                                                Threading.Thread.Sleep(100)

                                                If dv.Item(i)("CorrectiveAction").ToString <> String.Empty Then
                                                    If prevFac <> dv.Item(i)("FACILITY_ID") Then
                                                        If prevFac <> 0 Then
                                                            WordApp.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                                                            lineNumEnd = WordApp.Selection.Information(Word.WdInformation.wdFirstCharacterLineNumber)
                                                            WordApp.Selection.HomeKey(Unit:=Word.WdUnits.wdLine, Extend:=Word.WdMovementType.wdExtend)
                                                            If lineNumEnd - lineNumBegin - 1 > 0 Then
                                                                WordApp.Selection.MoveUp(Unit:=Word.WdUnits.wdLine, Count:=(lineNumEnd - lineNumBegin - 1), Extend:=Word.WdMovementType.wdExtend)
                                                            End If

                                                            WordApp.Selection.Range.ListFormat.ApplyListTemplate(ListTemplate:=wrdGlobal.ListGalleries.Item(Word.WdListGalleryType.wdNumberGallery).ListTemplates.Item(4), _
                                                               ContinuePreviousList:=False, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, _
                                                              DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                                                            WordApp.Selection.Font.Bold = False

                                                            WordApp.Selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=1)
                                                            WordApp.Selection.TypeParagraph()
                                                        End If
                                                        lineNumBegin = WordApp.Selection.Information(Word.WdInformation.wdFirstCharacterLineNumber)
                                                        prevFac = dv.Item(i)("FACILITY_ID")
                                                        If dtFacs.Rows.Count > 1 Then
                                                            WordApp.Selection.TypeText("Facility ID #(" + dv.Item(i)("FACILITY_ID").ToString + ")")
                                                            WordApp.Selection.TypeParagraph()
                                                        End If
                                                    End If

                                                    WordApp.Selection.TypeText(dv.Item(i)("CorrectiveAction").ToString)
                                                    If i < dtDiscreps.Rows.Count - 1 Then WordApp.Selection.TypeParagraph()

                                                    If i >= dtDiscreps.Rows.Count - 1 Then
                                                        'WordApp.Selection.TypeParagraph()
                                                        'WordApp.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=2)
                                                        WordApp.Selection.MoveLeft(Unit:=Word.WdUnits.wdCharacter, Count:=1)
                                                        lineNumEnd = WordApp.Selection.Information(Word.WdInformation.wdFirstCharacterLineNumber)
                                                        WordApp.Selection.HomeKey(Unit:=Word.WdUnits.wdLine, Extend:=Word.WdMovementType.wdExtend)
                                                        If lineNumEnd - lineNumBegin - 1 > 0 Then
                                                            WordApp.Selection.MoveUp(Unit:=Word.WdUnits.wdLine, Count:=(lineNumEnd - lineNumBegin - 1), Extend:=Word.WdMovementType.wdExtend)
                                                        End If

                                                        WordApp.Selection.Range.ListFormat.ApplyListTemplate(ListTemplate:=wrdGlobal.ListGalleries.Item(Word.WdListGalleryType.wdNumberGallery).ListTemplates.Item(4), _
                                                            ContinuePreviousList:=False, ApplyTo:=Word.WdListApplyTo.wdListApplyToWholeList, _
                                                            DefaultListBehavior:=Word.WdDefaultListBehavior.wdWord10ListBehavior)
                                                        WordApp.Selection.Font.Bold = False

                                                        WordApp.Selection.MoveDown(Unit:=Word.WdUnits.wdLine, Count:=1)
                                                        If alDiscrepsCorrectiveActionIndex.Contains(discrepCAIndex + 1) Then
                                                            WordApp.Selection.TypeParagraph()
                                                        End If
                                                    End If
                                                End If
                                            Next
                                            ' if there is a table below the corrective action in template to get extra text from
                                            'If alDiscrepsCorrectiveActionIndex.Contains(discrepCAIndex + 1) Then
                                            '    ' TODO - need to handle for multiple facilities
                                            '    WordApp.Selection.TypeParagraph()
                                            '    For i = 1 To WordApp.ActiveDocument.Tables.Item(discrepCAIndex + 1).Rows.Count
                                            '        WordApp.Selection.TypeText(WordApp.ActiveDocument.Tables.Item(discrepCAIndex + 1).Cell(i, 1).Range.Text)
                                            '        'If i < WordApp.ActiveDocument.Tables.Item(citindex + 1).Rows.Count Then
                                            '        '    WordApp.Selection.TypeParagraph()
                                            '        'End If
                                            '    Next
                                            'End If
                                        End With
                                    Next ' For Each discrepCAIndex As Integer In alDiscrepsCorrectiveActionIndex.ToArray
                                End If ' If alDiscrepsCorrectiveActionIndex.Count > 0 Then

                            End If ' If Not dtDiscreps Is Nothing Then
                            .Content.Find.Execute(FindText:="<Discrepancy>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)
                            .Content.Find.Execute(FindText:="<DiscrepCorrectiveAction>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)

                            ' remove the extra tables
                            Dim deltblCount As Integer = 0
                            If alCorrectiveActionAddOnIndex.Count > 0 Then
                                For Each i In alCorrectiveActionAddOnIndex.ToArray
                                    .Tables.Item(i - deltblCount).Delete()
                                    deltblCount += 1
                                Next
                            End If

                            .Save()

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
                                                    DestDoc.Application.Selection.InsertFile(FILENAME:=COCTemplatePath, ConfirmConversions:=False, Link:=False, Attachment:=False)

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

                                                    ' <Date>
                                                    strKey = "<Date>"
                                                    strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                                    .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                                    ' <Due Date>
                                                    strKey = "<Due Date>"
                                                    strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                                    .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                                Next ' For i = 0 To dtFacs.Rows.Count - 1
                                            End If ' If dtFacs.Rows.Count > 0 Then
                                        End If ' If Not dtFacs Is Nothing Then
                                    End If ' If System.IO.File.Exists(COCTemplatePath) Then
                                End If ' If COCTemplatePath <> String.Empty Then
                            End If

                            ' append coi
                            If COITemplatePath <> String.Empty Then
                                If System.IO.File.Exists(COITemplatePath) Then
                                    If Not dtCOIFacs Is Nothing Then
                                        If dtCOIFacs.Rows.Count > 0 Then
                                            dv = dtCOIFacs.DefaultView
                                            dv.Sort = "FACILITY_ID"
                                            Dim dr As DataRow
                                            For i = 0 To dtCOIFacs.Rows.Count - 1
                                                Threading.Thread.Sleep(100)

                                                dr = dtFacs.Select("FACILITY_ID = " + dv.Item(i)("FACILITY_ID").ToString)(0)
                                                ' insert page break
                                                DestDoc.Application.Selection.EndKey(Unit:=Word.WdUnits.wdStory)
                                                DestDoc.Application.Selection.InsertBreak(Word.WdBreakType.wdPageBreak)

                                                ' insert file
                                                DestDoc.Application.Selection.InsertFile(FILENAME:=COITemplatePath, ConfirmConversions:=False, Link:=False, Attachment:=False)

                                                ' set tags
                                                ' <COI Facility Name>
                                                strKey = "<COI Facility Name>"
                                                'strValue = dv.Item(i)("FACILITY")
                                                strValue = dr("FACILITY")
                                                'strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                                ' <COI Facility Address 1>
                                                strKey = "<COI Facility Address 1>"
                                                strValue = dr("ADDRESS_LINE_ONE")
                                                'strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                                ' <COI Facility Address 2>
                                                'strKey = "<COI Facility Address 2>"
                                                'strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                                '.Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                                ' <COI Facility City/State/Zip>
                                                strKey = "<COI Facility City/State/Zip>"
                                                strValue = dr("CITY") + ", " + _
                                                            dr("STATE") + " " + _
                                                            dr("ZIP")
                                                'strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                                ' <COI Facility ID>
                                                strKey = "<COI Facility ID>"
                                                strValue = dr("FACILITY_ID").ToString
                                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                                ' <Due Date>
                                                strKey = "<Due Date>"
                                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                                .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)

                                            Next ' For i = 0 To dtCOIFacs.Rows.Count - 1
                                        End If ' If dtCOIFacs.Rows.Count > 0 Then
                                    End If ' If Not dtCOIFacs Is Nothing Then
                                End If ' If System.IO.File.Exists(COITemplatePath) Then
                            End If ' If COITemplatePath <> String.Empty Then

                            Threading.Thread.Sleep(500)
                            .Save()

                        End With  'dest doc

                    End With ' word app

                Else
                    Throw New Exception("Unable to copy template " & TemplatePath & " to " & DestinationPath & " in pLetterGen object.")
                End If
            Catch ex As Exception
                SrcDoc = Nothing
                If Not WordApp Is Nothing Then
                    If Not WordApp.ActiveDocument Is Nothing Then WordApp.ActiveDocument.Close(False)
                End If
                System.IO.File.Delete(DestinationPath)
                Throw ex
            End Try
        End Function

        ' Financial
        Public Function CreateFinancialLetter(ByVal strModuleID As String, ByVal strLetter_To_Print As String, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal strfile As String = "", Optional ByVal strSignature As String = "") ', Optional ByVal strFiles As String = "")
            Try
                Dim DocumentPath As String
                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty
                Dim strToday As String = String.Empty
                Dim strDocName As String = String.Empty
                Dim bolKeyDeductionReasons As Boolean = False
                Dim bolKeyReimbursementConditions As Boolean = False
                Dim strDeductionReasonsValue As String = String.Empty
                Dim oPara As Word.Paragraph
                Dim i As Integer = 0
                Dim strReimbursementConditionsValue As String = String.Empty
                'Dim strEnvelopes() As String = strFiles.Split(",")

                'Instantiate the Word Object
                If IsNothing(WordApp) Then
                    WordApp = GetWordApp()
                End If

                If Not System.IO.File.Exists(TemplatePath) Then
                    Throw New Exception("File Not Found: " + TemplatePath)
                End If
                System.IO.File.Copy(TemplatePath, DestinationPath)

                If System.IO.File.Exists(DestinationPath) Then
                    With WordApp

                        DestDoc = .Documents.Open(DestinationPath)
                        DestDoc = WordApp.ActiveDocument

                        With DestDoc
                            .Activate()

                            ' Find and Replace the TAGs with Values.
                            For i = 0 To colParams.Count - 1
                                strKey = colParams.Keys(i).ToString
                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                If strKey <> "<Reimbursement Conditions>" And strKey <> "<DEDUCTION REASONS>" Then
                                    .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                                ElseIf strKey = "<DEDUCTION REASONS>" Then
                                    bolKeyDeductionReasons = True
                                    strDeductionReasonsValue = strValue
                                ElseIf strKey = "<Reimbursement Conditions>" Then
                                    strReimbursementConditionsValue = strValue
                                    bolKeyReimbursementConditions = True
                                End If
                            Next
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
                            'oPara = .Content.Paragraphs.Add()
                            'oPara.Range.Font.Bold = 0
                            'oPara.Range.Text = 
                            'oPara.ID = "BULLET"
                            'oPara.Format.SpaceAfter = 1
                            'oPara.Range.InsertParagraphAfter()
                            'CreateEnvelope(WordApp, DestDoc, "ABC", "1234", "Jackson", "1234", "")
                            'DestDoc.Envelope.Insert(ExtractAddress:=True, OmitReturnAddress:= _
                            'False, PrintBarCode:=False, PrintFIMA:=False, Height:=WordApp.InchesToPoints(4.13 _
                            '), Width:=WordApp.InchesToPoints(9.5), Address:="Mr James ", AutoText:= _
                            '    "ToolsCreateLabels", ReturnAddress:="", ReturnAutoText:= _
                            '"ToolsCreateLabels", AddressFromLeft:=Word.WdConstants.wdAutoPosition, AddressFromTop:= _
                            'Word.WdConstants.wdAutoPosition, ReturnAddressFromLeft:=Word.WdConstants.wdAutoPosition, _
                            'ReturnAddressFromTop:=Word.WdConstants.wdAutoPosition, DefaultOrientation:= _
                            'Word.WdEnvelopeOrientation.wdCenterLandscape, DefaultFaceUp:=True, PrintEPostage:=False)


                            .Save()

                        End With

                        .Visible = True

                    End With
                Else
                    Throw New Exception("Unable to copy template " & TemplatePath & " to " & DestinationPath & " in pLetterGen object.")
                End If
            Catch ex As Exception
                SrcDoc = Nothing
                If Not WordApp Is Nothing Then
                    If Not WordApp.ActiveDocument Is Nothing Then WordApp.ActiveDocument.Close(False)
                End If
                System.IO.File.Delete(DestinationPath)
                Throw ex
            End Try
        End Function
        Public Function CreateFinancialGenericLetter(ByVal strModuleID As String, ByVal strLetter_To_Print As String, ByVal colParams As Specialized.NameValueCollection, ByVal cols As Int16, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal TableFormat As Int16 = 35, Optional ByVal bApplyBorders As Boolean = False, Optional ByVal bApplyShading As Boolean = False, Optional ByVal bApplyFont As Boolean = True, Optional ByVal bApplyColor As Boolean = False, Optional ByVal bApplyHeader As Boolean = True, Optional ByVal bApplyLastRow As Boolean = False, Optional ByVal bApplyFirstCol As Boolean = False, Optional ByVal bApplyLastCol As Boolean = False, Optional ByVal bApplyAutofit As Boolean = True)
            Try
                Dim DocumentPath As String
                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty
                Dim strToday As String = String.Empty
                Dim strDocName As String = String.Empty
                Dim i As Integer = 0
                Dim oPara As Word.Paragraph

                'Instantiate the Word Object
                If IsNothing(WordApp) Then
                    WordApp = GetWordApp()
                End If

                If Not System.IO.File.Exists(TemplatePath) Then
                    Throw New Exception("File Not Found: " + TemplatePath)
                End If
                System.IO.File.Copy(TemplatePath, DestinationPath)

                If System.IO.File.Exists(DestinationPath) Then
                    With WordApp

                        DestDoc = .Documents.Open(DestinationPath)
                        DestDoc = WordApp.ActiveDocument

                        With DestDoc
                            .Activate()

                            ' Find and Replace the TAGs with Values.
                            For i = 0 To colParams.Count - 1

                                strKey = colParams.Keys(i).ToString
                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                If strKey = "<DATA>" Then
                                    oPara = .Content.Paragraphs.Add(.Bookmarks.Item("\endofdoc").Range)
                                    oPara.Range.Text = strValue
                                    'Word.WdTableFormat.wdTableFormatContemporary
                                    oPara.Range.ConvertToTable("|", , cols, , TableFormat, bApplyBorders, bApplyShading, bApplyFont, bApplyColor, bApplyHeader, bApplyLastRow, bApplyFirstCol, bApplyLastCol, bApplyAutofit)
                                    oPara.Format.SpaceAfter = 1
                                    oPara.Range.InsertParagraphAfter()

                                Else
                                    .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                                End If
                            Next


                            With DestDoc.PageSetup
                                .LineNumbering.Active = False
                                '.Orientation = Word.WdOrientation.wdOrientLandscape
                                .Orientation = Word.WdOrientation.wdOrientPortrait
                                .TopMargin = WordApp.InchesToPoints(0.88)
                                .BottomMargin = WordApp.InchesToPoints(1.25)
                                .LeftMargin = WordApp.InchesToPoints(1)
                                .RightMargin = WordApp.InchesToPoints(1)
                                .Gutter = WordApp.InchesToPoints(0)
                                .HeaderDistance = WordApp.InchesToPoints(0.5)
                                .FooterDistance = WordApp.InchesToPoints(0.5)
                                '.PageWidth = WordApp.InchesToPoints(11)
                                .PageHeight = WordApp.InchesToPoints(8.5)
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
                            'WordApp.ActiveWindow.ActivePane.SmallScroll(Down:=23)
                            'WordApp.ActiveWindow.ActivePane.VerticalPercentScrolled = 0
                            'WordApp.Selection.Tables.Item(1).Columns.Item(8).SetWidth(ColumnWidth:=180.3, RulerStyle:= _
                            '    Word.WdRulerStyle.wdAdjustNone)
                            'WordApp.Selection.Tables.Item(1).Columns.Item(7).SetWidth(ColumnWidth:=106.25, RulerStyle:= _
                            '    Word.WdRulerStyle.wdAdjustNone)
                            'WordApp.Selection.Tables.Item(1).Columns.Item(6).SetWidth(ColumnWidth:=106.05, RulerStyle:= _
                            '    Word.WdRulerStyle.wdAdjustNone)
                            'WordApp.Selection.Tables.Item(1).Columns.Item(5).SetWidth(ColumnWidth:=67.05, RulerStyle:= _
                            '    Word.WdRulerStyle.wdAdjustNone)
                            'WordApp.Selection.Tables.Item(1).Columns.Item(4).SetWidth(ColumnWidth:=131.65, RulerStyle:= _
                            '    Word.WdRulerStyle.wdAdjustNone)
                            'WordApp.ActiveWindow.ActivePane.LargeScroll(ToRight:=1)
                            'WordApp.ActiveWindow.ActivePane.HorizontalPercentScrolled = 0



                            .Save()
                        End With

                        'per Danny do not close the Word App when creating Letters.
                        'DestDoc = Nothing
                        '.Quit(False)
                    End With
                Else
                    Throw New Exception("Unable to copy template " & TemplatePath & " to " & DestinationPath & " in pLetterGen object.")
                End If

            Catch ex As Exception
                SrcDoc = Nothing
                If Not WordApp Is Nothing Then WordApp.Quit(False)
                Throw ex
            End Try

        End Function

        Public Function CreateCompanyInfoLetter(ByVal strModuleID As String, ByVal strLetter_To_Print As String, ByVal colParams As Specialized.NameValueCollection, ByVal TemplatePath As String, ByVal DestinationPath As String, Optional ByRef WordApp As Word.Application = Nothing, Optional ByVal strfile As String = "", Optional ByVal strSignature As String = "") ', Optional ByVal strFiles As String = "")
            Try
                Dim DocumentPath As String
                Dim strKey As String = String.Empty
                Dim strValue As String = String.Empty
                Dim strToday As String = String.Empty
                Dim strDocName As String = String.Empty
                Dim oPara As Word.Paragraph
                Dim i As Integer = 0
                Dim strInfoNeededValue As String = String.Empty
                'Dim strEnvelopes() As String = strFiles.Split(",")

                'Instantiate the Word Object
                If IsNothing(WordApp) Then
                    WordApp = GetWordApp()
                End If

                If Not System.IO.File.Exists(TemplatePath) Then
                    Throw New Exception("File Not Found: " + TemplatePath)
                End If
                System.IO.File.Copy(TemplatePath, DestinationPath)

                If System.IO.File.Exists(DestinationPath) Then
                    With WordApp

                        DestDoc = .Documents.Open(DestinationPath)
                        DestDoc = WordApp.ActiveDocument

                        With DestDoc
                            .Activate()

                            ' Find and Replace the TAGs with Values.
                            For i = 0 To colParams.Count - 1
                                strKey = colParams.Keys(i).ToString
                                strValue = IIf(IsNothing(colParams.Get(strKey).ToString), "", colParams.Get(strKey).ToString)
                                strValue = strValue.Replace(vbCrLf, Chr(13))

                                If strKey <> "<InfoNeeded>" Then
                                    .Content.Find.Execute(FindText:=strKey, ReplaceWith:=IIf(strValue = String.Empty, "", strValue), Replace:=Word.WdReplace.wdReplaceAll)
                                Else
                                    strInfoNeededValue = strValue
                                End If
                            Next

                            .Content.Find.Execute(FindText:="<InfoNeeded>", ReplaceWith:=String.Empty, Replace:=Word.WdReplace.wdReplaceAll)
                            If .Tables.Count > 0 Then
                                .Tables.Item(1).Cell(1, 1).Range.Text = strInfoNeededValue
                            End If
                            'oPara = .Content.Paragraphs.Add()
                            'oPara.Range.Font.Bold = 0
                            'oPara.Range.Text = 
                            'oPara.ID = "BULLET"
                            'oPara.Format.SpaceAfter = 1
                            'oPara.Range.InsertParagraphAfter()


                            .Save()

                        End With

                        .Visible = True

                    End With
                Else
                    Throw New Exception("Unable to copy template " & TemplatePath & " to " & DestinationPath & " in pLetterGen object.")
                End If
            Catch ex As Exception
                SrcDoc = Nothing
                If Not WordApp Is Nothing Then
                    If Not WordApp.ActiveDocument Is Nothing Then WordApp.ActiveDocument.Close(False)
                End If
                System.IO.File.Delete(DestinationPath)
                Throw ex
            End Try
        End Function

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
#Region "General Operations"
        Private Function GetWordApp() As Word.Application
            Dim WordApp As Word.Application
            Try
                If IsNothing(WordApp) Then
                    WordApp = GetObject(, "Word.Application")
                End If

                ' WordApp.Visible = True
            Catch ex As Exception
                If ex.Message.ToUpper = "Cannot Create ActiveX Component.".ToUpper Then
                    WordApp = New Word.Application
                ElseIf ex.Message.ToUpper = "The RPC server is unavailable.".ToUpper Then
                    WordApp = New Word.Application
                Else
                    Throw ex
                End If
            End Try
            Return WordApp

        End Function
#End Region

#End Region
    End Class
End Namespace
