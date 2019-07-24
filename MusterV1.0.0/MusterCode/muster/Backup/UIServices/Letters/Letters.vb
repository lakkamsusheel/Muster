Imports System.IO
Imports MUSTER.UIUtilsGen
Imports System.Text
Imports System.Windows.Forms
'Imports System.Web.Mail
'-------------------------------------------------------------------------------
' MUSTER.MUSTER.Letters
'   Letters Form
'
' Copyright (C) 2004 CIBER, Inc.
' All rights reserved.
'
' Release   Initials    Date      Description
'  1.0        ??      ??/??/??    Original class definition.
'  1.1        JVC2    02/08/2005  Integrated calls to new ProfileData
'-------------------------------------------------------------------------------
'
' TODO - Integrate with application 2/9/05
'
Public Class Letters
    Inherits System.Windows.Forms.Form




#Region "User Defined Variables"
    Private WithEvents WordApp As Word.Application
    Private strLastModule As String
    Private strLastYear As String

    Dim result As DialogResult
    'Dim LetterCons As LettersandReportsConsumer
    Dim nRegisterFlag As Integer = 0
    Dim rp As New Remove_Pencil
    Dim refreshCount As Integer = 0
    Dim curUser As String = String.Empty
    Dim bolLoading As Boolean = False
    Dim dsDocuments As DataSet
#End Region
#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub
    Public Sub New(ByVal RegisterFlag As Integer)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()
        nRegisterFlag = RegisterFlag
        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    Friend WithEvents rtxtContent As System.Windows.Forms.RichTextBox
    Friend WithEvents btnOpen As System.Windows.Forms.Button
    Friend WithEvents lblDocumentList As System.Windows.Forms.Label
    Friend WithEvents grpDocType As System.Windows.Forms.GroupBox
    Friend WithEvents rdUnprinted As System.Windows.Forms.RadioButton
    Friend WithEvents rdPrinted As System.Windows.Forms.RadioButton
    Friend WithEvents btnPrinted As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents ugDocumentList As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents cmbModule As System.Windows.Forms.ComboBox
    Friend WithEvents lblModule As System.Windows.Forms.Label
    Friend WithEvents cmbYear As System.Windows.Forms.ComboBox
    Friend WithEvents lblYear As System.Windows.Forms.Label
    Friend WithEvents pnlBottom As System.Windows.Forms.Panel
    Friend WithEvents pnlTop As System.Windows.Forms.Panel
    Friend WithEvents pnlDetails As System.Windows.Forms.Panel
    Friend WithEvents btnrefresh As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.rtxtContent = New System.Windows.Forms.RichTextBox
        Me.btnOpen = New System.Windows.Forms.Button
        Me.btnClose = New System.Windows.Forms.Button
        Me.lblDocumentList = New System.Windows.Forms.Label
        Me.grpDocType = New System.Windows.Forms.GroupBox
        Me.rdPrinted = New System.Windows.Forms.RadioButton
        Me.rdUnprinted = New System.Windows.Forms.RadioButton
        Me.btnPrinted = New System.Windows.Forms.Button
        Me.btnDelete = New System.Windows.Forms.Button
        Me.ugDocumentList = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.cmbModule = New System.Windows.Forms.ComboBox
        Me.lblModule = New System.Windows.Forms.Label
        Me.cmbYear = New System.Windows.Forms.ComboBox
        Me.lblYear = New System.Windows.Forms.Label
        Me.pnlBottom = New System.Windows.Forms.Panel
        Me.pnlTop = New System.Windows.Forms.Panel
        Me.pnlDetails = New System.Windows.Forms.Panel
        Me.btnrefresh = New System.Windows.Forms.Button
        Me.grpDocType.SuspendLayout()
        CType(Me.ugDocumentList, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlBottom.SuspendLayout()
        Me.pnlTop.SuspendLayout()
        Me.pnlDetails.SuspendLayout()
        Me.SuspendLayout()
        '
        'rtxtContent
        '
        Me.rtxtContent.Location = New System.Drawing.Point(776, 8)
        Me.rtxtContent.Name = "rtxtContent"
        Me.rtxtContent.Size = New System.Drawing.Size(16, 24)
        Me.rtxtContent.TabIndex = 3
        Me.rtxtContent.Text = ""
        Me.rtxtContent.Visible = False
        '
        'btnOpen
        '
        Me.btnOpen.BackColor = System.Drawing.SystemColors.Control
        Me.btnOpen.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnOpen.Location = New System.Drawing.Point(216, 8)
        Me.btnOpen.Name = "btnOpen"
        Me.btnOpen.Size = New System.Drawing.Size(96, 23)
        Me.btnOpen.TabIndex = 7
        Me.btnOpen.Text = "Open Letter"
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.SystemColors.Control
        Me.btnClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(440, 8)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(96, 23)
        Me.btnClose.TabIndex = 8
        Me.btnClose.Text = "Close"
        '
        'lblDocumentList
        '
        Me.lblDocumentList.Location = New System.Drawing.Point(16, 32)
        Me.lblDocumentList.Name = "lblDocumentList"
        Me.lblDocumentList.TabIndex = 130
        Me.lblDocumentList.Text = "Document List"
        '
        'grpDocType
        '
        Me.grpDocType.Controls.Add(Me.rdPrinted)
        Me.grpDocType.Controls.Add(Me.rdUnprinted)
        Me.grpDocType.Location = New System.Drawing.Point(16, 72)
        Me.grpDocType.Name = "grpDocType"
        Me.grpDocType.Size = New System.Drawing.Size(384, 48)
        Me.grpDocType.TabIndex = 131
        Me.grpDocType.TabStop = False
        Me.grpDocType.Text = "Document Type"
        '
        'rdPrinted
        '
        Me.rdPrinted.Location = New System.Drawing.Point(168, 24)
        Me.rdPrinted.Name = "rdPrinted"
        Me.rdPrinted.Size = New System.Drawing.Size(144, 16)
        Me.rdPrinted.TabIndex = 127
        Me.rdPrinted.Text = "Show Archived Letters"
        '
        'rdUnprinted
        '
        Me.rdUnprinted.Location = New System.Drawing.Point(8, 24)
        Me.rdUnprinted.Name = "rdUnprinted"
        Me.rdUnprinted.Size = New System.Drawing.Size(152, 16)
        Me.rdUnprinted.TabIndex = 126
        Me.rdUnprinted.Text = "Show UnArchived Letters"
        '
        'btnPrinted
        '
        Me.btnPrinted.BackColor = System.Drawing.SystemColors.Control
        Me.btnPrinted.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrinted.Location = New System.Drawing.Point(320, 8)
        Me.btnPrinted.Name = "btnPrinted"
        Me.btnPrinted.Size = New System.Drawing.Size(112, 23)
        Me.btnPrinted.TabIndex = 132
        Me.btnPrinted.Text = "Archive Letter(s)"
        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.SystemColors.Control
        Me.btnDelete.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnDelete.Location = New System.Drawing.Point(544, 8)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(96, 23)
        Me.btnDelete.TabIndex = 133
        Me.btnDelete.Text = "Delete"
        Me.btnDelete.Visible = False
        '
        'ugDocumentList
        '
        Me.ugDocumentList.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugDocumentList.DisplayLayout.Override.HeaderClickAction = Infragistics.Win.UltraWinGrid.HeaderClickAction.SortMulti
        Me.ugDocumentList.DisplayLayout.Scrollbars = Infragistics.Win.UltraWinGrid.Scrollbars.Both
        Me.ugDocumentList.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugDocumentList.Location = New System.Drawing.Point(0, 0)
        Me.ugDocumentList.Name = "ugDocumentList"
        Me.ugDocumentList.Size = New System.Drawing.Size(840, 310)
        Me.ugDocumentList.TabIndex = 134
        '
        'cmbModule
        '
        Me.cmbModule.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbModule.DropDownWidth = 160
        Me.cmbModule.ItemHeight = 13
        Me.cmbModule.Location = New System.Drawing.Point(432, 96)
        Me.cmbModule.Name = "cmbModule"
        Me.cmbModule.Size = New System.Drawing.Size(144, 21)
        Me.cmbModule.TabIndex = 135
        '
        'lblModule
        '
        Me.lblModule.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblModule.Location = New System.Drawing.Point(432, 72)
        Me.lblModule.Name = "lblModule"
        Me.lblModule.Size = New System.Drawing.Size(56, 23)
        Me.lblModule.TabIndex = 136
        Me.lblModule.Text = "Module"
        '
        'cmbYear
        '
        Me.cmbYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbYear.DropDownWidth = 60
        Me.cmbYear.ItemHeight = 13
        Me.cmbYear.Location = New System.Drawing.Point(600, 96)
        Me.cmbYear.Name = "cmbYear"
        Me.cmbYear.Size = New System.Drawing.Size(56, 21)
        Me.cmbYear.TabIndex = 137
        '
        'lblYear
        '
        Me.lblYear.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblYear.Location = New System.Drawing.Point(600, 72)
        Me.lblYear.Name = "lblYear"
        Me.lblYear.Size = New System.Drawing.Size(56, 23)
        Me.lblYear.TabIndex = 138
        Me.lblYear.Text = "Year"
        '
        'pnlBottom
        '
        Me.pnlBottom.Controls.Add(Me.btnrefresh)
        Me.pnlBottom.Controls.Add(Me.btnPrinted)
        Me.pnlBottom.Controls.Add(Me.btnDelete)
        Me.pnlBottom.Controls.Add(Me.btnOpen)
        Me.pnlBottom.Controls.Add(Me.btnClose)
        Me.pnlBottom.Controls.Add(Me.rtxtContent)
        Me.pnlBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlBottom.Location = New System.Drawing.Point(0, 438)
        Me.pnlBottom.Name = "pnlBottom"
        Me.pnlBottom.Size = New System.Drawing.Size(840, 120)
        Me.pnlBottom.TabIndex = 139
        '
        'pnlTop
        '
        Me.pnlTop.Controls.Add(Me.lblDocumentList)
        Me.pnlTop.Controls.Add(Me.grpDocType)
        Me.pnlTop.Controls.Add(Me.cmbModule)
        Me.pnlTop.Controls.Add(Me.lblModule)
        Me.pnlTop.Controls.Add(Me.cmbYear)
        Me.pnlTop.Controls.Add(Me.lblYear)
        Me.pnlTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTop.Name = "pnlTop"
        Me.pnlTop.Size = New System.Drawing.Size(840, 128)
        Me.pnlTop.TabIndex = 140
        '
        'pnlDetails
        '
        Me.pnlDetails.Controls.Add(Me.ugDocumentList)
        Me.pnlDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlDetails.Location = New System.Drawing.Point(0, 128)
        Me.pnlDetails.Name = "pnlDetails"
        Me.pnlDetails.Size = New System.Drawing.Size(840, 310)
        Me.pnlDetails.TabIndex = 141
        '
        'btnrefresh
        '
        Me.btnrefresh.BackColor = System.Drawing.SystemColors.Control
        Me.btnrefresh.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnrefresh.Location = New System.Drawing.Point(8, 8)
        Me.btnrefresh.Name = "btnrefresh"
        Me.btnrefresh.Size = New System.Drawing.Size(184, 24)
        Me.btnrefresh.TabIndex = 134
        Me.btnrefresh.Text = "Refresh/Audit Archives"
        '
        'Letters
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.AutoScroll = True
        Me.ClientSize = New System.Drawing.Size(840, 558)
        Me.Controls.Add(Me.pnlDetails)
        Me.Controls.Add(Me.pnlTop)
        Me.Controls.Add(Me.pnlBottom)
        Me.Location = New System.Drawing.Point(60, 120)
        Me.Name = "Letters"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "Letters"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.grpDocType.ResumeLayout(False)
        CType(Me.ugDocumentList, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlBottom.ResumeLayout(False)
        Me.pnlTop.ResumeLayout(False)
        Me.pnlDetails.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "Form Events"
    Private Sub Letters_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Try
            bolLoading = True

            If _container.GoToUser.Length > 0 Then
                Text = String.Format("Letters For User: {0}", _container.GoToUser)
                curUser = _container.GoToUser
            Else
                curUser = MusterContainer.AppUser.ID
            End If

            LoadPrimaryModules()
            LoadCalendarYear()
            'If nRegisterFlag = 1 Then
            rdUnprinted.Checked = True
            bolLoading = False
            cmbModule.SelectedValue = MusterContainer.AppUser.DefaultModule
            rdUnprinted_CheckedChanged(sender, e)
            'End If

            ' cmbModule.SelectedIndex = -1

            _container.GoToUser = String.Empty
            _container.GotoYear = String.Empty

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub Letters_Activated(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.Activated
        ugDocumentList.Focus()
        If ugDocumentList.Rows.Count > 0 Then
            If ugDocumentList.ActiveRow Is Nothing Then
                ugDocumentList.ActiveRow = ugDocumentList.Rows(0)
            End If
        End If
        'lstDocumentList.Focus()
        'If lstDocumentList.Items.Count > 0 And lstDocumentList.SelectedItems.Count > 0 Then
        '    lstDocumentList.SelectedItems(0).Selected = True
        'End If
    End Sub
#End Region
#Region "Change Events"
    Private Sub rdPrinted_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdPrinted.CheckedChanged
        Try
            If rdUnprinted.Checked Then Exit Sub
            LoadCalendarYear(1)
            PopulateLetters()
            btnDelete.Visible = False
            btnrefresh.Visible = True
            btnPrinted.Enabled = False
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub rdUnprinted_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rdUnprinted.CheckedChanged
        Try
            If rdPrinted.Checked Then Exit Sub
            LoadCalendarYear()
            PopulateLetters()
            btnDelete.Visible = True
            btnrefresh.Visible = False
            btnPrinted.Enabled = True
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbModule_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbModule.SelectedIndexChanged
        Try
            If bolLoading Then Exit Sub
            If dsDocuments.Tables.Count <= 0 Then Exit Sub
            ugDocumentList.DataSource = Nothing

            'If rdUnprinted.Checked Then
            '    LoadCalendarYear(False)
            'ElseIf rdPrinted.Checked Then
            '    LoadCalendarYear(True)
            'End If

            If cmbModule.SelectedIndex > -1 Then
                strLastModule = cmbModule.Text


                If Not cmbModule.Text = "ALL" Then
                    dsDocuments.Tables(0).DefaultView.RowFilter = "Module_ID = " + UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString + _
                                                                                " AND CALENDAR_YEAR = " + IIf(cmbYear.SelectedIndex <> -1, UIUtilsGen.GetComboBoxValueString(cmbYear).ToString, " CALENDAR_YEAR")
                Else
                    dsDocuments.Tables(0).DefaultView.RowFilter = "CALENDAR_YEAR = " + IIf(cmbYear.SelectedIndex <> -1, UIUtilsGen.GetComboBoxValueString(cmbYear).ToString, " CALENDAR_YEAR")
                End If
                LoadDocumentGrid(dsDocuments.Tables(0))
            End If


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub cmbYear_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmbYear.SelectedIndexChanged
        Try
            If bolLoading Then Exit Sub

            If dsDocuments.Tables.Count <= 0 Then Exit Sub

            If Me.cmbYear.SelectedIndex > -1 Then
                ugDocumentList.DataSource = Nothing
                If rdUnprinted.Checked Or rdPrinted.Checked Then



                    Me.strLastYear = Me.cmbYear.Text
                    If cmbModule.Text = "ALL" Then
                        dsDocuments.Tables(0).DefaultView.RowFilter = "CALENDAR_YEAR = " + IIf(cmbYear.SelectedIndex <> -1, UIUtilsGen.GetComboBoxValueString(cmbYear).ToString, " CALENDAR_YEAR")
                    Else
                        dsDocuments.Tables(0).DefaultView.RowFilter = "Module_ID = " + IIf(cmbModule.SelectedIndex <> -1, UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString, "Module_ID") + " AND CALENDAR_YEAR = " + IIf(cmbYear.SelectedIndex <> -1, UIUtilsGen.GetComboBoxValueString(cmbYear).ToString, " CALENDAR_YEAR")
                    End If

                End If

                LoadDocumentGrid(dsDocuments.Tables(0))
            End If


        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "Button Events"

    Private Sub btnRefresh_Clicked(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnrefresh.Click


        If Me.ugDocumentList.Selected.Rows Is Nothing OrElse Me.ugDocumentList.Selected.Rows.Count = 0 Then
            MsgBox("Please Select all the rows you want to refresh")
            Exit Sub
        End If
        Cursor.Current = Cursors.WaitCursor

        For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In Me.ugDocumentList.Selected.Rows
            Me.ugDocumentList.ActiveRow = row

            'refreshes without opening
            btnOpen_Click(sender, Nothing)
        Next

        Cursor.Current = Cursors.Arrow

        MsgBox(String.Format("You have updated {0} document(s) from temp letter to archive to match filename is database records", refreshCount))

        ugDocumentList.Selected.Rows.Clear()

        refreshCount = 0

    End Sub

    Private Sub btnOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpen.Click

        Dim DOC_PATH As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_SystemGenerated).ProfileValue & "\"
        Dim SrcDoc As Word.Document
        Try
            If rdPrinted.Checked = False And rdUnprinted.Checked = False Then
                MessageBox.Show("Please select any one of the Document Type.")
                Exit Sub
            End If

            If ugDocumentList.ActiveRow Is Nothing Then
                MessageBox.Show("Please Select a Document to Open.")
                Exit Sub
            End If

            Dim textfile As String = ugDocumentList.ActiveRow.Cells("Document Name").Text


            If Not System.IO.File.Exists(ugDocumentList.ActiveRow.Cells("Document Location").Text + ugDocumentList.ActiveRow.Cells("Document Name").Text) AndAlso Me.rdUnprinted.Checked Then
                textfile = textfile.ToUpper.Replace(".DOC", "_TEMPLATE.DOC")
            End If


            If Not System.IO.File.Exists(ugDocumentList.ActiveRow.Cells("Document Location").Text + textfile) AndAlso Me.rdUnprinted.Checked Then
                MessageBox.Show("Invalid path. Please make sure the network path is accessible. " + vbCrLf + ugDocumentList.ActiveRow.Cells("Document Location").Text + ugDocumentList.ActiveRow.Cells("Document Name").Text + vbCrLf + "This document might have been manually archived outside the system. Try archiving the record")
                Exit Sub
            ElseIf Not System.IO.File.Exists(ugDocumentList.ActiveRow.Cells("Document Location").Text + textfile) AndAlso Me.rdPrinted.Checked Then

                Dim found As Boolean = False
                Dim mode As Integer = 0
                Dim p As String = textfile

                ' Loop through ann finds the file in temp
                For mode = 0 To 2

                    Dim list As ArrayList

                    If mode >= 1 Then p = p.Substring(0, p.LastIndexOf("_")) + "*." + p.Substring(p.LastIndexOf(".") + 1)

                    list = GetFilesRecursive(DOC_PATH, p)

                    For Each Path As String In list

                        If File.Exists(Path) Then
                            found = True

                            If e Is Nothing Then
                                refreshCount += 1
                            End If

                            System.IO.File.Move(Path, ugDocumentList.ActiveRow.Cells("Document Location").Text + ugDocumentList.ActiveRow.Cells("Document Name").Text)
                            Exit For
                        End If
                    Next
                    If found Then Exit For
                Next


                If Not found Then
                    If (MessageBox.Show("Invalid path" + vbCrLf + ugDocumentList.ActiveRow.Cells("Document Location").Text + textfile + "This document may not have been properly archived. Would you like to delete this file?", "No archive document", MessageBoxButtons.YesNo) = DialogResult.Yes) Then
                        btnDelete_Click(sender, e)
                    End If


                    If Not e Is Nothing Then
                        Exit Sub
                    Else
                        refreshCount += 1
                    End If

                End If

            End If

            If Not e Is Nothing Then

                If textfile.ToUpper.IndexOf(".DOC") > -1 Then



                    WordApp = UIUtilsGen.GetWordApp

                    If Not WordApp Is Nothing Then
                        WordApp.Visible = True

                        If Not ugDocumentList.Rows.Count <= 0 Then
                            ' SrcDoc = WordApp.Documents.Open(lstDocumentList.SelectedItems(0).SubItems(6).Text + CStr(Trim(lstDocumentList.SelectedItems(0).SubItems(1).Text)))
                            SrcDoc = WordApp.Documents.Open(ugDocumentList.ActiveRow.Cells("Document Location").Text + textfile)
                        End If
                    End If

                Else
                    UIUtilsGen.OpenInPDFFile(ugDocumentList.ActiveRow.Cells("Document Location").Text + textfile)
                End If

            End If



        Catch ex As Exception
            'Delay()
            If Not e Is Nothing Then
                UIUtilsGen.Delay(, 2)
                If Not WordApp Is Nothing Then
                    If Not SrcDoc Is Nothing Then
                        SrcDoc.Close(False)
                    End If
                End If
                Dim MyErr As ErrorReport
                MyErr = New ErrorReport(New Exception("Cannot open the file in Word: " + ex.Message, ex))
                MyErr.ShowDialog()

            End If

        Finally
            'Delay()
            If Not e Is Nothing Then
                UIUtilsGen.Delay(, 2)
                SrcDoc = Nothing
                WordApp = Nothing
                ugDocumentList.Focus()
                If ugDocumentList.Rows.Count > 0 Then
                    If ugDocumentList.ActiveRow Is Nothing Then
                        ugDocumentList.ActiveRow = ugDocumentList.Rows(0)
                    End If
                End If
            End If

        End Try
    End Sub
    Private Sub btnPrinted_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrinted.Click
        Dim DOC_PATH As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_SystemArchive).ProfileValue & "\"
        Dim strPrintedPath As String = DOC_PATH
        Dim strUnPrintedPath As String = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_SystemGenerated).ProfileValue & "\"
        Dim doc As Word.Document
        Dim unfoundDocs As String = String.Empty
        Dim strModuleName As String = String.Empty
        Try

            If DOC_PATH = "\" Then
                MsgBox("Unspecified System Archive Path. Please contact the Administrator to update the path.")
                Exit Sub
            End If

            If strUnPrintedPath = "\" Then
                MsgBox("Unspecified System Generated Document Path. Please contact the Administrator to update the path.")
                Exit Sub
            End If

            If ugDocumentList.Rows.Count <= 0 Then
                MessageBox.Show("No Records to Modify.")
                Exit Sub
            End If

            If ugDocumentList.Selected.Rows.Count <= 0 Then
                If Not ugDocumentList.ActiveRow Is Nothing Then
                    ugDocumentList.ActiveRow.Selected = True
                Else
                    MsgBox("Select a Document to Print.")
                    Exit Sub
                End If
            End If

            If ugDocumentList.Selected.Rows.Count > 0 Then
                result = MessageBox.Show("Are you Sure you want to Archive the Selected Document(s)?", "UST Letters", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                If result = DialogResult.No Then
                    Exit Sub
                End If
            End If


         
            Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow
            Dim ownerID As String
            Dim emailQuery As String
            Dim emailDs As DataSet

            Dim emailCount As Int16 = 1
            Dim objShell As Object
            Dim emailObj As New DataObject
            For Each ugRow In ugDocumentList.Rows


                If ugRow.Selected Then

                    Dim letter As String = ugRow.Cells("Document Name").Text
                    Dim emails As String = String.Empty
                    strModuleName = UIUtilsGen.GetModuleNameByID(CInt(Trim(ugRow.Cells("Module_ID").Text)))
                    strPrintedPath = DOC_PATH & strModuleName & "\" & CStr(Format(Now, "yyyy")) & "\"
                    If strModuleName = "CAE" Or strModuleName = "Inspection" Then
                        ownerID = ugRow.Cells("Entity ID").Text()
                        emailQuery = "Select contact_name, C.email_address as cEmail, C.email_address_personal as cEmail2, " + _
                                     "Owner_ID,O.email_address as oEmail, O.email_address_personal as oEmail2 " + _
                                     "From vCONTACT_LIST C left join tblReg_Owner O on C.entityID = O.Owner_ID " + _
                                     "Where Owner_ID = '" + ownerID + "'"
                        emailDs = MusterContainer.pLetter.GetDataSet(emailQuery)
                        If Not emailDs Is Nothing And Not emailDs.Tables Is Nothing And emailDs.Tables.Count > 0 And emailDs.Tables(0).Rows.Count > 0 Then

                            For Each row As DataRow In emailDs.Tables(0).Rows
                                If Not System.IO.File.Exists(strUnPrintedPath + letter) Then
                                    letter = letter.ToUpper.Replace(".DOC", "_TEMPLATE.DOC")
                                End If
                                If row("cEmail") <> String.Empty Then
                                    emails = emails + row("cEmail")
                                End If
                                If row("cEmail2") <> String.Empty Then
                                    emails = emails + ", " + row("cEmail2")
                                End If
                                If row("oEmail") <> String.Empty Then
                                    emails = emails + ", " + row("oEmail")
                                End If
                                If row("oEmail2") <> String.Empty Then
                                    emails = emails + ", " + row("oEmail2")
                                End If

                            Next
                            If MsgBox("Press Ctrl + C " + vbCrLf + emails + vbCrLf + strPrintedPath + letter + vbCrLf + _
                                      "Would you like to send emails first and then continue archiving.", MsgBoxStyle.YesNo) = vbYes Then

                                '   emailObj.SetData(DataFormats.Text, emails)
                                'emailObj.SetData(DataFormats.Text, strUnPrintedPath + letter)
                                Clipboard.SetDataObject(emails + vbCrLf + strPrintedPath + letter, True)
                                objShell = CreateObject("Wscript.Shell")
                                objShell.Run("mailto: " + emails)

                                'Shell("mailto: " + emails)
                            End If
                        End If
                    End If




                    If Not System.IO.Directory.Exists(strPrintedPath) Then
                        System.IO.Directory.CreateDirectory(strPrintedPath)
                    End If

                    'To Mark Printed the Selected Document and Move it to System Archive Path.
                    'If (LetterCons.MakePrinted(CInt(Trim(lstDocumentList.SelectedItems(0).SubItems(0).Text)), DOC_PATH)) Then
                    'MusterContainer.pLetter.Retrieve(CInt(Trim(ugDocumentList.ActiveRow.Cells("ID").Text)))
                    'MusterContainer.pLetter.DatePrinted = Today.ToShortDateString
                    'MusterContainer.pLetter.DocumentLocation = DOC_PATH
                    'MusterContainer.pLetter.Save()

                    'try archiving physical record from templetters
                    Try

                        If Not System.IO.File.Exists(strUnPrintedPath + letter) Then
                            letter = letter.ToUpper.Replace(".DOC", "_TEMPLATE.DOC")
                        End If

                        System.IO.File.Move(strUnPrintedPath + letter, strPrintedPath + ugRow.Cells("Document Name").Text)


                    Catch ex As FileNotFoundException

                        'If not archived already, look for previous years
                        If Not File.Exists(strPrintedPath + letter) Then

                            Dim list As ArrayList = UIUtilsGen.GetFilesRecursive(DOC_PATH, letter)
                            Dim found As Boolean = False

                            ' Loop through and display each path.
                            For Each Path As String In list

                                If File.Exists(Path) Then
                                    found = True
                                    System.IO.File.Move(Path, strPrintedPath + ugRow.Cells("Document Name").Text)

                                    Exit For

                                End If
                            Next

                            If Not found Then
                                unfoundDocs = String.Format("{0}{1}", IIf(unfoundDocs Is Nothing, String.Empty, String.Format("{0}{1}", unfoundDocs, vbCrLf)), strUnPrintedPath + letter)
                            End If
                        End If


                    Catch Ex As Exception
                        Throw Ex
                    End Try

                    'complete archiving of database record
                    MusterContainer.pLetter.UpdatePrintedStatus(CInt(Trim(ugRow.Cells("Document_ID").Text)), strPrintedPath, Today.ToShortDateString)

                End If
            Next

            'if any unfound temp docs
            If Not unfoundDocs Is Nothing AndAlso unfoundDocs.Length > 0 Then
                MsgBox(String.Format("Archiving completed, but there are unarchived documents that were not found neither in the temp folder nor the archive directory.{0}{1}", vbCrLf, unfoundDocs))

            End If

            'LetterCons = Nothing
            PopulateLetters()
            'End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot open the file in Word: " + ex.Message, ex))
            MyErr.ShowDialog()

        Finally
            unfoundDocs = Nothing
            WordApp = Nothing
        End Try

    End Sub
    Private Sub btnClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClose.Click
        Try
            Me.Close()
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Try

            'If lstDocumentList.SelectedItems.Count <= 0 Then
            '    MessageBox.Show("No Records to Delete.")
            '    Exit Sub
            'End If

            If ugDocumentList.Rows.Count <= 0 Then
                MessageBox.Show("No Records to Delete.")
                Exit Sub
            End If

            If ugDocumentList.Selected.Rows.Count <= 0 Then
                If Not ugDocumentList.ActiveRow Is Nothing Then
                    ugDocumentList.ActiveRow.Selected = True
                Else
                    MsgBox("Select a Document to Delete.")
                    Exit Sub
                End If
            End If

            Dim ugRow As Infragistics.Win.UltraWinGrid.UltraGridRow

            result = MessageBox.Show("Are you Sure you want to Delete this Record?", "UST Letters", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If result = DialogResult.No Then
                Exit Sub
            End If

            ' check if the doc is open
            WordApp = UIUtilsGen.GetWordApp

            If Not WordApp Is Nothing Then
                For Each ugRow In ugDocumentList.Selected.Rows
                    For Each doc As Word.Document In WordApp.Documents
                        If ugRow.Cells("Document Name").Text = doc.Name Then
                            MsgBox("Document is currently open." + vbCrLf + _
                                    "Document has to be closed to be deleted.", , "UST Letters")
                            Exit Sub
                        End If
                    Next
                Next
            End If



            'If Not ugDocumentList.Rows.Count <= 0 Then
            '    If Not ugDocumentList.ActiveRow Is Nothing Then
            '        MusterContainer.pLetter.Retrieve(CInt(Trim(ugDocumentList.ActiveRow.Cells("Document_ID").Text)))
            '        MusterContainer.pLetter.Deleted = True
            '        MusterContainer.pLetter.Save()

            '        ' delete physical file
            '        Try
            '            If System.IO.File.Exists(ugDocumentList.ActiveRow.Cells("Document Location").Text + ugDocumentList.ActiveRow.Cells("Document Name").Text) Then
            '                System.IO.File.Delete(ugDocumentList.ActiveRow.Cells("Document Location").Text + ugDocumentList.ActiveRow.Cells("Document Name").Text)
            '            End If
            '        Catch ex As Exception
            '            ' do nothing
            '        End Try

            '        PopulateLetters()
            '    End If
            'End If

            For Each ugRow In ugDocumentList.Selected.Rows
                MusterContainer.pLetter.Retrieve(CInt(Trim(ugRow.Cells("Document_ID").Text)))
                MusterContainer.pLetter.Deleted = True
                MusterContainer.pLetter.Save()
                Dim letter As String


                ' delete physical file
                Try
                    letter = ugRow.Cells("Document Name").Text

                    If Not System.IO.File.Exists(ugRow.Cells("Document Location").Text + letter) Then
                        letter = letter.ToUpper.Replace(".DOC", "_TEMPLATE.DOC")
                    End If

                    If System.IO.File.Exists(ugRow.Cells("Document Location").Text + letter) Then
                        System.IO.File.Delete(ugRow.Cells("Document Location").Text + letter)
                    End If
                Catch ex As Exception
                    ' do nothing
                End Try

            Next

            PopulateLetters()

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try

    End Sub
    Private Sub ugDocumentList_DoubleClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles ugDocumentList.DoubleClick

        If Not UIUtilsInfragistics.WinGridRowDblClicked(sender, e) Then
            Me.Cursor = Cursors.Default
            Exit Sub
        End If

        btnOpen.PerformClick()
    End Sub
#End Region
#Region "Miscellaneous Functions"
    Private Sub PopulateLetters()
        Try


            'LetterCons = New LettersandReportsConsumer
            'Dim Ltrs As New InfoRepository.Letters
            'Dim Ltr As InfoRepository.Letter

            'If rdUnprinted.Checked = True Then
            '    Ltrs = LetterCons.GetLetters(0)
            'ElseIf rdPrinted.Checked = True Then
            '    Ltrs = LetterCons.GetLetters(1)
            'Else
            '    If Not nRegisterFlag = 1 Then
            '        MessageBox.Show("Select atleast one Document Type.")
            '        Exit Sub
            '    End If
            'End If
            'If Ltrs.Count <= 0 Then
            '    If Not nRegisterFlag = 1 Then
            '        MessageBox.Show("No Letters Found.")
            '        Exit Sub
            '    End If
            'End If

            'For Each Ltr In Ltrs

            '    If rdUnprinted.Checked = True Then
            '        lstDocumentList.Items.Add(New ListViewItem(New String() {Ltr.DOCUMENT_ID, Ltr.DOCUMENT_NAME, Ltr.TYPE_OF_DOCUMENT, Ltr.DOCUMENT_DESCRIPTION, "", Ltr.DATE_EDITED.ToShortDateString, Ltr.DOCUMENT_LOCATION}))
            '    Else
            '        lstDocumentList.Items.Add(New ListViewItem(New String() {Ltr.DOCUMENT_ID, Ltr.DOCUMENT_NAME, Ltr.TYPE_OF_DOCUMENT, Ltr.DOCUMENT_DESCRIPTION, Ltr.DATE_PRINTED, Ltr.DATE_EDITED.ToShortDateString, Ltr.DOCUMENT_LOCATION}))
            '    End If
            'Next


            Dim dtTable As DataTable
            Dim drow As DataRow



            If rdUnprinted.Checked Then
                dsDocuments = MusterContainer.pLetter.GetDocumentsList(curUser, False)
                'dtTable = MusterContainer.pLetter.UnPrintedLetterTable(MusterContainer.AppUser.ID)
            ElseIf rdPrinted.Checked Then
                dsDocuments = MusterContainer.pLetter.GetDocumentsList(curUser, True)
                'dtTable = MusterContainer.pLetter.PrintedLetterTable(MusterContainer.AppUser.ID)
            Else
                If Not nRegisterFlag = 1 Then
                    MessageBox.Show("Select atleast one Document Type.")
                    Exit Sub
                End If
            End If


            'For Each drow In dtTable.Rows
            '    If rdUnprinted.Checked Then
            '        lstDocumentList.Items.Add(New ListViewItem(New String() {drow("ID"), drow("Name"), drow("DocumentType"), drow("DocumentDescription"), "", drow("date_created"), drow("DocumentLocation")}))
            '    Else
            '        lstDocumentList.Items.Add(New ListViewItem(New String() {drow("ID"), drow("Name"), drow("DocumentType"), drow("DocumentDescription"), drow("Date Printed"), drow("date_created"), drow("DocumentLocation")}))
            '    End If
            'Next

            cmbModule_SelectedIndexChanged(Me, Nothing)

            'cmbModule.SelectedText = Me.strLastYear

            'If cmbYear.Items.Count > 0 Then
            'cmbYear.SelectedItem = Me.strLastYear
            'End If

            'dtTable = dsDocuments.Tables(0)
            'ugDocumentList.DataSource = Nothing
            'dtTable.DefaultView.Sort = "[Date Created] ASC"

            'If Not cmbModule.SelectedIndex = -1 And Not cmbYear.SelectedIndex = -1 Then
            '    dtTable.DefaultView.RowFilter = "Module_ID = " + UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString + " AND CALENDAR_YEAR = " + UIUtilsGen.GetComboBoxValueString(cmbYear).ToString
            'ElseIf Not cmbModule.SelectedIndex = -1 Then
            '    dtTable.DefaultView.RowFilter = "Module_ID = " + UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString
            'End If



            'dsDocuments.Tables(0).DefaultView.RowFilter = "Module_ID = " + IIf(cmbModule.SelectedIndex <> -1, UIUtilsGen.GetComboBoxValueInt(cmbModule).ToString, "Module_ID") + _
            '                                                  " AND CALENDAR_YEAR = " + IIf(cmbYear.SelectedIndex <> -1, UIUtilsGen.GetComboBoxValueString(cmbYear).ToString, " CALENDAR_YEAR")



            'LoadDocumentGrid(dtTable)

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Public Sub boldToHTML(ByVal doc As Word.Document)
        Try
            With doc.Content.Find

                .ClearFormatting()
                .Font.Bold = 1
                .Font.Color = Word.WdColor.wdColorDarkRed
                .Text = "Dear"
                .Replacement.ClearFormatting()
                .Replacement.Font.Bold = 0
                .Font.Color = Word.WdColor.wdColorDarkRed
                '.Text = "*"
                .Replacement.Text = "Hii"
                .Execute(findtext:="TEMPLATES HERE", _
                            ReplaceWith:="MohanraJ", _
                   Format:=True, _
                   Replace:=Word.WdReplace.wdReplaceAll)
                'ReplaceWith:="<b>^&</b>", _
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Public Sub italicToHTML(ByVal doc As Word.Document)
        Try
            With doc.Content.Find
                .ClearFormatting()
                .Font.Italic = 1
                .Replacement.ClearFormatting()
                .Replacement.Font.Bold = 0
                .Text = "*"
                .Execute(findtext:="", _
                   ReplaceWith:="<i>^&</i>", _
                   Format:=True, _
                   Replace:=Word.WdReplace.wdReplaceAll)
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
    Private Sub LoadPrimaryModules()
        bolLoading = True
        'Dim dtTable As DataTable = MusterContainer.AppUser.ListModulesUserHasAccessTo(MusterContainer.AppUser.UserKey)
        Dim dtTable As DataTable = MusterContainer.AppUser.ListModulesUserHasAccessTo(MusterContainer.AppUser.UserKey)
        Dim drow As DataRow
        drow = dtTable.NewRow
        drow("PROPERTY_NAME") = "ALL"
        drow("PROPERTY_ID") = "0"
        dtTable.Rows.Add(drow)
        dtTable.DefaultView.Sort = "PROPERTY_NAME"
        'dtTable.DefaultView.RowFilter = "PROPERTY_ID NOT IN (894,1303,1311,1312)"
        cmbModule.DataSource = dtTable.DefaultView 'MusterContainer.AppUser.ListPrimaryModules
        cmbModule.DisplayMember = "PROPERTY_NAME"
        cmbModule.ValueMember = "PROPERTY_ID"
        bolLoading = False
    End Sub
    Private Sub LoadCalendarYear(Optional ByVal PrintedFlag As Integer = 0)
        bolLoading = True
        Dim dsYear As DataSet
        dsYear = MusterContainer.pLetter.GetCalendarYear(curUser, PrintedFlag)
        If dsYear.Tables.Count <= 0 Then Exit Sub
        cmbYear.DataSource = dsYear.Tables(0).DefaultView
        cmbYear.DisplayMember = "DATE_CREATED"
        cmbYear.ValueMember = "DATE_CREATED"

        If _container.GotoYear <> String.Empty AndAlso IsNumeric(_container.GotoYear) Then
            cmbYear.SelectedValue = Convert.ToInt16(_container.GotoYear)
        Else
            cmbYear.SelectedValue = Year(Today)
        End If


        'If cmbYear.Items.Count > 0 Then
        '    cmbYear.SelectedIndex = -1
        'End If
        bolLoading = False
    End Sub
    Private Sub LoadDocumentGrid(ByVal dtTable As DataTable)
        ugDocumentList.DataSource = dtTable
        ugDocumentList.DrawFilter = rp
        ugDocumentList.DisplayLayout.Bands(0).Columns("Entity Type").Width = 100
        ugDocumentList.DisplayLayout.Bands(0).Columns("Entity ID").Width = 75
        ugDocumentList.DisplayLayout.Bands(0).Columns("Document Type").Width = 200
        ugDocumentList.DisplayLayout.Bands(0).Columns("Document Name").Width = 350
        ugDocumentList.DisplayLayout.Bands(0).Columns("Date Printed").Width = 75
        ugDocumentList.DisplayLayout.Bands(0).Columns("Date Created").Width = 75

        ugDocumentList.DisplayLayout.Bands(0).Columns("Description").Hidden = True
        ugDocumentList.DisplayLayout.Bands(0).Columns("Created_By").Hidden = True
        ugDocumentList.DisplayLayout.Bands(0).Columns("Document Location").Hidden = True
        ugDocumentList.DisplayLayout.Bands(0).Columns("Entity Type ID").Hidden = True
        ugDocumentList.DisplayLayout.Bands(0).Columns("last_edited_by").Hidden = True
        ugDocumentList.DisplayLayout.Bands(0).Columns("date_last_edited").Hidden = True
        If curUser <> MusterContainer.AppUser.ID OrElse MusterContainer.AppUser.ID = "ADMIN" Then
            ugDocumentList.DisplayLayout.Bands(0).Columns("Owning_User").Header.Caption = "Owner User"
        Else
            ugDocumentList.DisplayLayout.Bands(0).Columns("Owning_User").Hidden = True
        End If

        ugDocumentList.DisplayLayout.Bands(0).Columns("MODULE_ID").Hidden = True
        ugDocumentList.DisplayLayout.Bands(0).Columns("DOCUMENT_ID").Hidden = True
        ugDocumentList.DisplayLayout.Bands(0).Columns("CALENDAR_YEAR").Hidden = True
        ugDocumentList.DisplayLayout.Bands(0).Columns("EVENT_ID").Hidden = True

        ugDocumentList.DisplayLayout.Bands(0).Columns("EVENT_SEQUENCE").Header.Caption = "Event #"
        ugDocumentList.DisplayLayout.Bands(0).Columns("EVENT_TYPE").Header.Caption = "Event Type"

        If rdUnprinted.Checked Then
            ugDocumentList.DisplayLayout.Bands(0).Columns("Date Printed").Hidden = True
        End If

        ugDocumentList.DisplayLayout.Override.AllowUpdate = Infragistics.Win.DefaultableBoolean.False
        ugDocumentList.DisplayLayout.Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.Free
        ugDocumentList.DisplayLayout.Override.RowSizingArea = Infragistics.Win.UltraWinGrid.RowSizingArea.EntireRow
        ugDocumentList.DisplayLayout.Bands(0).Override.RowSizing = Infragistics.Win.UltraWinGrid.RowSizing.AutoFree
        ugDocumentList.DisplayLayout.Bands(0).Override.RowSizingAutoMaxLines = 5
        ugDocumentList.DisplayLayout.Bands(0).Override.CellClickAction = Infragistics.Win.UltraWinGrid.CellClickAction.RowSelect


    End Sub
    'Function to Delay after closing a word document to resolve RPC Server is Unavailable Issue
    'Private Sub Delay()
    '    Dim MyTime As DateTime
    '    MyTime = Now
    '    Do Until Now > MyTime.AddSeconds(2)
    '    Loop
    'End Sub
#End Region
End Class
