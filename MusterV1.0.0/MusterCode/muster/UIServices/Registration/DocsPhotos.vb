Imports System.IO
Imports System.Text
Public Class DocsPhotos
    '-------------------------------------------------------------------------------
    ' MUSTER.MUSTER.DocsPhotos
    '   Documents and Photos Form
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
    Inherits System.Windows.Forms.Form
    Private dir As DirectoryInfo
    Private dirInsp As DirectoryInfo
    Private docDir As DirectoryInfo

    Private CurrentPage As Integer = 1
    Private nTotalPages As Integer
    Private objDrRow() As DataRow
    Private docFileinfo() As FileInfo
    Dim fileinfo() As FileInfo
    Friend WithEvents mstContainer As MusterContainer
    Private myPicture As PictureBox
    Private nNoOfImages As Integer
    Private nStartImage As Integer
    Private nEndImage As Integer
    Friend boolselectImages As Boolean
    Private nFacilityId As Integer = 0
    Private no As Integer = 1
    Private pnlImage As Panel
    Private strFileNames() As String
    'Private DOCPATH As String = String.Empty
    Private IMGPATH As String = String.Empty
    Private pnlDesc As Panel
    Private txtDesc As TextBox
    Dim strFileName As String = String.Empty
    'Dim i As Integer = 0

#Region " Windows Form Designer generated code "

    Public Sub New(ByRef mContainer As MusterContainer)
        MyBase.New()
        mstContainer = mContainer

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        If Not mstContainer.ProfileData Is Nothing Then
            mstContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_FacImages).Reset()
            IMGPATH = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_FacImages).ProfileValue
            'DOCPATH = MusterContainer.ProfileData.Retrieve(UIUtilsGen.FilePathKey_SystemArchive).ProfileValue
        End If
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
    Friend WithEvents tbCntrlDocsImages As System.Windows.Forms.TabControl
    Friend WithEvents tbPageImages As System.Windows.Forms.TabPage
    Friend WithEvents pnlImages As System.Windows.Forms.Panel
    Friend WithEvents tbPageImage As System.Windows.Forms.TabPage
    Friend WithEvents pnlImageView As System.Windows.Forms.Panel
    Friend WithEvents pnlImageBottom As System.Windows.Forms.Panel
    Friend WithEvents btnImageClose As System.Windows.Forms.Button
    Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
    Friend WithEvents pnlCountBar As System.Windows.Forms.Panel
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents btnPrevious As System.Windows.Forms.Button
    Friend WithEvents lblThroughValue As System.Windows.Forms.Label
    Friend WithEvents lblThrough As System.Windows.Forms.Label
    Friend WithEvents lblImagesValue As System.Windows.Forms.Label
    Friend WithEvents lblTotalImages As System.Windows.Forms.Label
    Friend WithEvents lblTotalImagesValue As System.Windows.Forms.Label
    Friend WithEvents lblImages As System.Windows.Forms.Label
    Friend WithEvents pnlImgTop As System.Windows.Forms.Panel
    Friend WithEvents pnlImgMiddle As System.Windows.Forms.Panel
    Friend WithEvents pnlImgBottom As System.Windows.Forms.Panel
    Friend WithEvents txtLargeComments As System.Windows.Forms.TextBox
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.tbCntrlDocsImages = New System.Windows.Forms.TabControl
        Me.tbPageImages = New System.Windows.Forms.TabPage
        Me.pnlImages = New System.Windows.Forms.Panel
        Me.pnlImgBottom = New System.Windows.Forms.Panel
        Me.pnlImgMiddle = New System.Windows.Forms.Panel
        Me.pnlImgTop = New System.Windows.Forms.Panel
        Me.pnlCountBar = New System.Windows.Forms.Panel
        Me.btnNext = New System.Windows.Forms.Button
        Me.btnPrevious = New System.Windows.Forms.Button
        Me.lblThroughValue = New System.Windows.Forms.Label
        Me.lblThrough = New System.Windows.Forms.Label
        Me.lblImagesValue = New System.Windows.Forms.Label
        Me.lblTotalImages = New System.Windows.Forms.Label
        Me.lblTotalImagesValue = New System.Windows.Forms.Label
        Me.lblImages = New System.Windows.Forms.Label
        Me.tbPageImage = New System.Windows.Forms.TabPage
        Me.pnlImageBottom = New System.Windows.Forms.Panel
        Me.btnImageClose = New System.Windows.Forms.Button
        Me.pnlImageView = New System.Windows.Forms.Panel
        Me.PictureBox1 = New System.Windows.Forms.PictureBox
        Me.txtLargeComments = New System.Windows.Forms.TextBox
        Me.tbCntrlDocsImages.SuspendLayout()
        Me.tbPageImages.SuspendLayout()
        Me.pnlImages.SuspendLayout()
        Me.pnlCountBar.SuspendLayout()
        Me.tbPageImage.SuspendLayout()
        Me.pnlImageBottom.SuspendLayout()
        Me.pnlImageView.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbCntrlDocsImages
        '
        Me.tbCntrlDocsImages.Appearance = System.Windows.Forms.TabAppearance.FlatButtons
        Me.tbCntrlDocsImages.Controls.Add(Me.tbPageImages)
        Me.tbCntrlDocsImages.Controls.Add(Me.tbPageImage)
        Me.tbCntrlDocsImages.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tbCntrlDocsImages.Location = New System.Drawing.Point(0, 0)
        Me.tbCntrlDocsImages.Name = "tbCntrlDocsImages"
        Me.tbCntrlDocsImages.SelectedIndex = 0
        Me.tbCntrlDocsImages.Size = New System.Drawing.Size(776, 525)
        Me.tbCntrlDocsImages.TabIndex = 0
        '
        'tbPageImages
        '
        Me.tbPageImages.Controls.Add(Me.pnlImages)
        Me.tbPageImages.Controls.Add(Me.pnlCountBar)
        Me.tbPageImages.Location = New System.Drawing.Point(4, 25)
        Me.tbPageImages.Name = "tbPageImages"
        Me.tbPageImages.Size = New System.Drawing.Size(768, 496)
        Me.tbPageImages.TabIndex = 1
        Me.tbPageImages.Text = "Images"
        '
        'pnlImages
        '
        Me.pnlImages.Controls.Add(Me.pnlImgBottom)
        Me.pnlImages.Controls.Add(Me.pnlImgMiddle)
        Me.pnlImages.Controls.Add(Me.pnlImgTop)
        Me.pnlImages.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlImages.DockPadding.All = 4
        Me.pnlImages.Location = New System.Drawing.Point(0, 0)
        Me.pnlImages.Name = "pnlImages"
        Me.pnlImages.Size = New System.Drawing.Size(768, 478)
        Me.pnlImages.TabIndex = 0
        '
        'pnlImgBottom
        '
        Me.pnlImgBottom.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlImgBottom.Location = New System.Drawing.Point(4, 316)
        Me.pnlImgBottom.Name = "pnlImgBottom"
        Me.pnlImgBottom.Size = New System.Drawing.Size(760, 156)
        Me.pnlImgBottom.TabIndex = 6
        '
        'pnlImgMiddle
        '
        Me.pnlImgMiddle.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlImgMiddle.Location = New System.Drawing.Point(4, 160)
        Me.pnlImgMiddle.Name = "pnlImgMiddle"
        Me.pnlImgMiddle.Size = New System.Drawing.Size(760, 156)
        Me.pnlImgMiddle.TabIndex = 5
        '
        'pnlImgTop
        '
        Me.pnlImgTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlImgTop.Location = New System.Drawing.Point(4, 4)
        Me.pnlImgTop.Name = "pnlImgTop"
        Me.pnlImgTop.Size = New System.Drawing.Size(760, 156)
        Me.pnlImgTop.TabIndex = 4
        '
        'pnlCountBar
        '
        Me.pnlCountBar.Controls.Add(Me.btnNext)
        Me.pnlCountBar.Controls.Add(Me.btnPrevious)
        Me.pnlCountBar.Controls.Add(Me.lblThroughValue)
        Me.pnlCountBar.Controls.Add(Me.lblThrough)
        Me.pnlCountBar.Controls.Add(Me.lblImagesValue)
        Me.pnlCountBar.Controls.Add(Me.lblTotalImages)
        Me.pnlCountBar.Controls.Add(Me.lblTotalImagesValue)
        Me.pnlCountBar.Controls.Add(Me.lblImages)
        Me.pnlCountBar.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlCountBar.Location = New System.Drawing.Point(0, 478)
        Me.pnlCountBar.Name = "pnlCountBar"
        Me.pnlCountBar.Size = New System.Drawing.Size(768, 18)
        Me.pnlCountBar.TabIndex = 6
        '
        'btnNext
        '
        Me.btnNext.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnNext.Location = New System.Drawing.Point(360, 0)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(32, 18)
        Me.btnNext.TabIndex = 7
        Me.btnNext.Text = ">"
        '
        'btnPrevious
        '
        Me.btnPrevious.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnPrevious.Location = New System.Drawing.Point(328, 0)
        Me.btnPrevious.Name = "btnPrevious"
        Me.btnPrevious.Size = New System.Drawing.Size(32, 18)
        Me.btnPrevious.TabIndex = 6
        Me.btnPrevious.Text = "<"
        '
        'lblThroughValue
        '
        Me.lblThroughValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblThroughValue.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblThroughValue.Location = New System.Drawing.Point(136, 0)
        Me.lblThroughValue.Name = "lblThroughValue"
        Me.lblThroughValue.Size = New System.Drawing.Size(32, 18)
        Me.lblThroughValue.TabIndex = 5
        Me.lblThroughValue.Text = "0"
        '
        'lblThrough
        '
        Me.lblThrough.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblThrough.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblThrough.Location = New System.Drawing.Point(80, 0)
        Me.lblThrough.Name = "lblThrough"
        Me.lblThrough.Size = New System.Drawing.Size(56, 18)
        Me.lblThrough.TabIndex = 4
        Me.lblThrough.Text = "Through"
        '
        'lblImagesValue
        '
        Me.lblImagesValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblImagesValue.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblImagesValue.Location = New System.Drawing.Point(48, 0)
        Me.lblImagesValue.Name = "lblImagesValue"
        Me.lblImagesValue.Size = New System.Drawing.Size(32, 18)
        Me.lblImagesValue.TabIndex = 3
        Me.lblImagesValue.Text = "0"
        '
        'lblTotalImages
        '
        Me.lblTotalImages.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalImages.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblTotalImages.Location = New System.Drawing.Point(636, 0)
        Me.lblTotalImages.Name = "lblTotalImages"
        Me.lblTotalImages.Size = New System.Drawing.Size(100, 18)
        Me.lblTotalImages.TabIndex = 2
        Me.lblTotalImages.Text = "Total Images"
        '
        'lblTotalImagesValue
        '
        Me.lblTotalImagesValue.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblTotalImagesValue.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblTotalImagesValue.Location = New System.Drawing.Point(736, 0)
        Me.lblTotalImagesValue.Name = "lblTotalImagesValue"
        Me.lblTotalImagesValue.Size = New System.Drawing.Size(32, 18)
        Me.lblTotalImagesValue.TabIndex = 1
        Me.lblTotalImagesValue.Text = "0"
        '
        'lblImages
        '
        Me.lblImages.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.lblImages.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblImages.Location = New System.Drawing.Point(0, 0)
        Me.lblImages.Name = "lblImages"
        Me.lblImages.Size = New System.Drawing.Size(48, 18)
        Me.lblImages.TabIndex = 0
        Me.lblImages.Text = "Images"
        '
        'tbPageImage
        '
        Me.tbPageImage.Controls.Add(Me.pnlImageBottom)
        Me.tbPageImage.Controls.Add(Me.pnlImageView)
        Me.tbPageImage.Location = New System.Drawing.Point(4, 25)
        Me.tbPageImage.Name = "tbPageImage"
        Me.tbPageImage.Size = New System.Drawing.Size(768, 496)
        Me.tbPageImage.TabIndex = 2
        '
        'pnlImageBottom
        '
        Me.pnlImageBottom.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlImageBottom.Controls.Add(Me.txtLargeComments)
        Me.pnlImageBottom.Controls.Add(Me.btnImageClose)
        Me.pnlImageBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlImageBottom.Location = New System.Drawing.Point(0, 456)
        Me.pnlImageBottom.Name = "pnlImageBottom"
        Me.pnlImageBottom.Size = New System.Drawing.Size(768, 40)
        Me.pnlImageBottom.TabIndex = 1
        '
        'btnImageClose
        '
        Me.btnImageClose.Location = New System.Drawing.Point(8, 8)
        Me.btnImageClose.Name = "btnImageClose"
        Me.btnImageClose.TabIndex = 0
        Me.btnImageClose.Text = "Close"
        '
        'pnlImageView
        '
        Me.pnlImageView.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlImageView.Controls.Add(Me.PictureBox1)
        Me.pnlImageView.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlImageView.Location = New System.Drawing.Point(0, 0)
        Me.pnlImageView.Name = "pnlImageView"
        Me.pnlImageView.Size = New System.Drawing.Size(768, 496)
        Me.pnlImageView.TabIndex = 0
        '
        'PictureBox1
        '
        Me.PictureBox1.BackColor = System.Drawing.SystemColors.AppWorkspace
        Me.PictureBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.PictureBox1.Location = New System.Drawing.Point(0, 0)
        Me.PictureBox1.Name = "PictureBox1"
        Me.PictureBox1.Size = New System.Drawing.Size(764, 492)
        Me.PictureBox1.TabIndex = 0
        Me.PictureBox1.TabStop = False
        '
        'txtLargeComments
        '
        Me.txtLargeComments.Location = New System.Drawing.Point(96, 8)
        Me.txtLargeComments.Name = "txtLargeComments"
        Me.txtLargeComments.Size = New System.Drawing.Size(664, 20)
        Me.txtLargeComments.TabIndex = 1
        Me.txtLargeComments.Text = ""
        '
        'DocsPhotos
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(776, 525)
        Me.Controls.Add(Me.tbCntrlDocsImages)
        Me.Location = New System.Drawing.Point(60, 120)
        Me.Name = "DocsPhotos"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "DocsPhotos"
        Me.tbCntrlDocsImages.ResumeLayout(False)
        Me.tbPageImages.ResumeLayout(False)
        Me.pnlImages.ResumeLayout(False)
        Me.pnlCountBar.ResumeLayout(False)
        Me.tbPageImage.ResumeLayout(False)
        Me.pnlImageBottom.ResumeLayout(False)
        Me.pnlImageView.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region


    Public Function getfiles(ByVal nFacilityId As Integer) As String
        Dim arr As New ArrayList
        Dim d As Integer
        Dim str As String = ""
        Dim oFileInfo() As FileInfo
        Dim oFileInfoInsp() As FileInfo

        If IMGPATH = String.Empty Then
            Throw New Exception("Document Path Unspecified. Please enter the Facilities path in Admin File Paths.")
        End If
        dir = New DirectoryInfo(IMGPATH & "\")
        dirInsp = New DirectoryInfo(IMGPATH & "\")

        oFileInfo = dir.GetFiles("F" & nFacilityId.ToString + "_*.JPG")
        '  oFileInfoInsp = dirInsp.GetFiles("Inspection_F_" & nFacilityId.ToString + "_*.JPG")
        nNoOfImages += oFileInfo.Length
        ' nNoOfImages += oFileInfoInsp.Length
        For d = 0 To oFileInfo.Length - 1
            str += oFileInfo(d).Name + ","
        Next
        oFileInfoInsp = dirInsp.GetFiles("Inspection_F_" & nFacilityId.ToString + "_*.JPG")
        nNoOfImages += oFileInfoInsp.Length
        For d = 0 To oFileInfoInsp.Length - 1
            str += oFileInfoInsp(d).Name + ","
        Next
        Return str
    End Function

    Public Function images(Optional ByVal strFiles As String = Nothing) As Integer
        'Dim endPage As Integer = (CurrentPage - 1) * 10
        'Dim startPage As Integer = endPage + 9
        Dim endPage As Integer = (CurrentPage - 1) * 15
        Dim startPage As Integer = endPage + 14
        Dim j As Integer
        Dim ttToolTip As New ToolTip

        If IMGPATH = String.Empty Then
            Throw New Exception("Document Path Unspecified. Please enter the Facilities path in Admin File Paths.")
        End If

        dir = New DirectoryInfo(IMGPATH & "\")
        If Not strFiles = Nothing Then
            If strFiles.EndsWith(",") Then
                strFiles.Remove(strFiles.Length - 1, 1)
            End If
        End If
        If IsNothing(strFileNames) Then
            strFileNames = strFiles.Split(",")
        End If

        tbCntrlDocsImages.SelectedTab = Me.tbPageImages
        'nTotalPages = System.Math.Ceiling(nNoOfImages / 10)
        nTotalPages = System.Math.Ceiling(nNoOfImages / 15)

        Try
            For j = startPage To endPage Step -1
                If j < nNoOfImages And nNoOfImages <> -1 Then
                    pnlImage = New Panel
                    myPicture = New PictureBox
                    myPicture.Text = strFileNames(j)
                    ttToolTip.SetToolTip(myPicture, strFileNames(j))
                    myPicture.Dock = DockStyle.Top
                    myPicture.Size = New System.Drawing.Size(112, 120)
                    myPicture.Cursor = Cursors.Hand
                    myPicture.Tag = no
                    myPicture.Image = Image.FromFile(dir.FullName + "\" + myPicture.Text)
                    myPicture.SizeMode = PictureBoxSizeMode.StretchImage

                    pnlImage.Width = 150
                    pnlImage.DockPadding.All = 4
                    pnlImage.Dock = DockStyle.Left
                    pnlImage.BorderStyle = BorderStyle.Fixed3D
                    AddHandler myPicture.Click, AddressOf ClickMyPicture
                    pnlImage.Controls.Add(myPicture)

                    Dim strTxtFileName As String = String.Empty
                    strTxtFileName = strFileNames(j).Substring(0, strFileNames(j).Length - 4).ToString
                    txtDesc = New TextBox
                    txtDesc.Name = strTxtFileName
                    txtDesc.Dock = DockStyle.Bottom
                    txtDesc.Tag = strTxtFileName + ".txt"
                    If System.IO.File.Exists(dir.FullName + strTxtFileName + ".txt") Then
                        txtDesc.Text = ReadFile(dir.FullName + strTxtFileName + ".txt")
                    Else
                        txtDesc.Text = ""
                    End If

                    AddHandler txtDesc.Leave, AddressOf DescriptionChanged
                    pnlImage.Controls.Add(txtDesc)
                    If no <= 15 Then
                        If no > 5 Then
                            If no > 10 Then
                                pnlImgBottom.Controls.Add(pnlImage)
                            Else
                                pnlImgMiddle.Controls.Add(pnlImage)
                            End If
                        Else
                            pnlImgTop.Controls.Add(pnlImage)
                        End If
                    End If
                    'If no <= 10 Then
                    '    If no > 5 Then
                    '        pnlImgBottom.Controls.Add(pnlImage)
                    '    Else
                    '        pnlImgTop.Controls.Add(pnlImage)
                    '    End If
                    'End If
                    no += 1
                    nStartImage = endPage
                    nEndImage = startPage
                End If
            Next
            Return nNoOfImages
        Catch ex As Exception
            Throw New Exception("No Images Found")
        End Try

    End Function

    Public Sub ImagesOnPage()
        lblImagesValue.Text = nStartImage + 1
        If nNoOfImages > nEndImage Then
            lblThroughValue.Text = nEndImage + 1
        Else
            lblThroughValue.Text = nNoOfImages
        End If
        lblTotalImagesValue.Text = nNoOfImages
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            boolselectImages = True
            tbCntrlDocsImages.SelectedTab = tbPageImages
            ImagesOnPage()
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot open Images:  " + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Public Sub ClickMyPicture(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try

            Dim clickedPicture As PictureBox = CType(sender, PictureBox)
            tbCntrlDocsImages.SelectedTab = Me.tbPageImage
            PictureBox1.Image = Image.FromFile(dir.FullName + "\" + clickedPicture.Text)
            PictureBox1.SizeMode = PictureBoxSizeMode.StretchImage
            Dim clickedComments As String = clickedPicture.Text.ToLower.Replace("jpg", "txt")
            'Dim clickedComments As String = clickedPicture.Text.Replace("JPG", "txt")
            If System.IO.File.Exists(dir.FullName + "\" + clickedComments) Then
                txtLargeComments.Text = ReadFile(dir.FullName + "\" + clickedComments)
            Else
                txtLargeComments.Text = ""
            End If

        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot open Picture:  " + ex.Message, ex))
            MyErr.ShowDialog()
        End Try

    End Sub

    Private Sub DescriptionChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Try
            Dim tempDesc As TextBox = CType(sender, TextBox)
            CreateFile(CStr(tempDesc.Tag), tempDesc.Text)
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot create Description File:  " + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Function ReadFile(ByVal FilePath As String) As String
        Dim returnDesc As String = String.Empty
        Dim oFile As System.IO.File
        Dim oRead As System.IO.StreamReader
        oRead = oFile.OpenText(FilePath)
        returnDesc = oRead.ReadLine()
        'EntireFile = oRead.ReadToEnd()
        oRead.Close()
        Return returnDesc
    End Function

    Private Sub CreateFile(ByVal FileName As String, ByVal Description As String)
        Dim oFile As System.IO.File
        Dim oWrite As System.IO.StreamWriter
        oWrite = oFile.CreateText(dir.FullName + FileName)
        oWrite.Write(Description)
        oWrite.Close()
    End Sub

    Private Sub btnImageClose_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImageClose.Click
        Try
            tbCntrlDocsImages.SelectedTab = Me.tbPageImages
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    'Friend Function populateListView(ByVal struserId As String, Optional ByVal drrow() As DataRow = Nothing, Optional ByVal nFileEntityId As Integer = 0) As Integer
    '    Dim dr As DataRow
    '    Dim strStatus As String
    '    Dim docStatusFileInfo As FileInfo
    '    Dim strDirectoryName As String = struserId
    '    Dim strYear As String = Today.Year
    '    nFacilityId = nFileEntityId
    '    Dim k As Integer
    '    Dim i As Integer
    '    Dim fileexists As Boolean = False
    '    Dim strfileName As String
    '    Try

    '        If DOCPATH = String.Empty Then
    '            Throw New Exception("Document Path Unspecified. Please enter the Facilities path in Admin File Paths.")
    '        End If

    '        '
    '        ' Get anything based on current pathing
    '        '
    '        'docDir = New DirectoryInfo(DOCPATH & "\")

    '        'Try
    '        '    docFileinfo = docDir.GetFiles("*" + nFacilityId.ToString + "_*")
    '        'Catch ex As Exception
    '        'End Try

    '        'objDrRow = drrow

    '        'If drrow.Length > 0 Then
    '        '    For Each dr In drrow
    '        '        If Not IsNothing(dr("DOCUMENT_LOCATION")) Then
    '        '            strfileName = dr("DOCUMENT_LOCATION")
    '        '            docStatusFileInfo = New FileInfo(strfileName)
    '        '            If docStatusFileInfo.Exists Then
    '        '                If docFileinfo.Length > 0 Then
    '        '                    strStatus = "Available"
    '        '                Else
    '        '                    strStatus = "Unavailable"
    '        '                End If
    '        '            Else
    '        '                strStatus = "Unavailable"
    '        '            End If
    '        '        End If

    '        '        If Not docFileinfo Is Nothing Then
    '        '            If docFileinfo.Length > 0 And i < docFileinfo.Length Then
    '        '                For k = 0 To ListView1.Items.Count - 1
    '        '                    strfileName = ListView1.Items(k).Text.ToString
    '        '                    If strfileName = docFileinfo(i).Name Then
    '        '                        fileexists = True
    '        '                    End If
    '        '                Next

    '        '                If dr("DOCUMENT_NAME") <> docFileinfo(i).Name Then
    '        '                    ListView1.Items.Add(New ListViewItem(New String() {dr("DOCUMENT_NAME"), dr("TYPE_OF_DOCUMENT"), IIf(IsDBNull(dr("DATE_LAST_EDITED")), "", dr("DATE_LAST_EDITED")), strStatus, dr("DOCUMENT_LOCATION") + dr("DOCUMENT_NAME")}))
    '        '                    If fileexists <> True Then
    '        '                    End If
    '        '                Else
    '        '                    ListView1.Items.Add(New ListViewItem(New String() {dr("DOCUMENT_NAME"), dr("TYPE_OF_DOCUMENT"), IIf(IsDBNull(dr("DATE_LAST_EDITED")), "", dr("DATE_LAST_EDITED")), strStatus, dr("DOCUMENT_LOCATION") + dr("DOCUMENT_NAME")}))
    '        '                End If
    '        '                i = i + 1

    '        '            Else
    '        '                ListView1.Items.Add(New ListViewItem(New String() {dr("DOCUMENT_NAME"), dr("TYPE_OF_DOCUMENT"), IIf(IsDBNull(dr("DATE_LAST_EDITED")), "", dr("DATE_LAST_EDITED")), strStatus, dr("DOCUMENT_LOCATION") + dr("DOCUMENT_NAME")}))
    '        '            End If
    '        '        End If
    '        '    Next
    '        'End If

    '        ''*******************************************************************************************
    '        ''
    '        '' Get anything based on previously filed documents
    '        ''
    '        'If drrow.Length > 0 Then
    '        '    For Each dr In drrow
    '        '        If Not IsNothing(dr("DOCUMENT_LOCATION")) Then
    '        '            docDir = New DirectoryInfo(dr("DOCUMENT_LOCATION") & "\")
    '        '            Exit For
    '        '        End If
    '        '    Next
    '        'End If
    '        'Try
    '        '    docFileinfo = docDir.GetFiles("*" + nFacilityId.ToString + "_*")
    '        'Catch ex As Exception
    '        'End Try

    '        'objDrRow = drrow

    '        'If drrow.Length > 0 Then
    '        '    For Each dr In drrow
    '        '        If Not IsNothing(dr("DOCUMENT_LOCATION")) Then
    '        '            strfileName = dr("DOCUMENT_LOCATION")
    '        '            docStatusFileInfo = New FileInfo(strfileName)
    '        '            If docStatusFileInfo.Exists Then
    '        '                If docFileinfo.Length > 0 Then
    '        '                    strStatus = "Available"
    '        '                Else
    '        '                    strStatus = "Unavailable"
    '        '                End If
    '        '            Else
    '        '                strStatus = "Unavailable"
    '        '            End If
    '        '        End If

    '        '        If Not docFileinfo Is Nothing Then
    '        '            If docFileinfo.Length > 0 And i < docFileinfo.Length Then
    '        '                For k = 0 To ListView1.Items.Count - 1
    '        '                    strfileName = ListView1.Items(k).Text.ToString
    '        '                    If strfileName = docFileinfo(i).Name Then
    '        '                        fileexists = True
    '        '                    End If
    '        '                Next

    '        '                If dr("DOCUMENT_NAME") <> docFileinfo(i).Name Then
    '        '                    ListView1.Items.Add(New ListViewItem(New String() {dr("DOCUMENT_NAME"), dr("TYPE_OF_DOCUMENT"), IIf(IsDBNull(dr("DATE_LAST_EDITED")), "", dr("DATE_LAST_EDITED")), strStatus, dr("DOCUMENT_LOCATION") + dr("DOCUMENT_NAME")}))
    '        '                    If fileexists <> True Then
    '        '                    End If
    '        '                Else
    '        '                    ListView1.Items.Add(New ListViewItem(New String() {dr("DOCUMENT_NAME"), dr("TYPE_OF_DOCUMENT"), IIf(IsDBNull(dr("DATE_LAST_EDITED")), "", dr("DATE_LAST_EDITED")), strStatus, dr("DOCUMENT_LOCATION") + dr("DOCUMENT_NAME")}))
    '        '                End If
    '        '                i = i + 1

    '        '            Else
    '        '                ListView1.Items.Add(New ListViewItem(New String() {dr("DOCUMENT_NAME"), dr("TYPE_OF_DOCUMENT"), IIf(IsDBNull(dr("DATE_LAST_EDITED")), "", dr("DATE_LAST_EDITED")), strStatus, dr("DOCUMENT_LOCATION") + dr("DOCUMENT_NAME")}))
    '        '            End If
    '        '        End If
    '        '    Next
    '        'End If

    '        'Return ListView1.Items.Count

    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Function

    Private Sub btnPrevious_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrevious.Click
        Try
            no = 1
            boolselectImages = True
            CurrentPage -= 1
            If CurrentPage < 1 Then
                CurrentPage = 1
            Else
                pnlImgTop.Controls.Clear()
                pnlImgMiddle.Controls.Clear()
                pnlImgBottom.Controls.Clear()
                images()
                ImagesOnPage()
            End If

        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        Try
            no = 1
            boolselectImages = True
            CurrentPage += 1
            If CurrentPage > nTotalPages Then
                CurrentPage = nTotalPages
            Else
                pnlImgTop.Controls.Clear()
                pnlImgMiddle.Controls.Clear()
                pnlImgBottom.Controls.Clear()
                images()
                ImagesOnPage()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    'Private Sub ListView1_DoubleClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim dr As DataRow
    '    Dim strFileStatus As String
    '    Dim fileName As Object
    '    Dim oWord As Word.Application
    '    Dim oDoc As Word.Document
    '    Dim missing As Object = System.Reflection.Missing.Value
    '    Dim areadOnly As Object = True
    '    Dim confirmConversions As Object = False
    '    Dim isVisible As Object = False
    '    Try

    '        'If ListView1.Items.Count > 0 Then
    '        '    Dim lstItem As ListViewItem = Me.ListView1.SelectedItems(0)
    '        '    strFileStatus = lstItem.SubItems(3).Text.ToUpper
    '        '    If strFileStatus.IndexOf(UCase("Unavailable")) >= 0 Then
    '        '        Me.RchTextDoc.Clear()
    '        '        btnDescription.Enabled = False
    '        '        Throw New Exception("File not found")

    '        '    Else
    '        '        fileName = lstItem.SubItems(4).Text
    '        '        btnDescription.Enabled = True
    '        '    End If
    '        'End If
    '        'Me.RchTextDoc.Clear()
    '        'oWord = New Word.Application
    '        'oWord.Visible = False
    '        'oDoc = oWord.Documents.Open(fileName, confirmConversions, [areadOnly])

    '        'Me.RchTextDoc.Text = oDoc.Content.Text
    '        'RchTextDoc.Refresh()

    '    Catch ex As Exception
    '        Dim MyErr As ErrorReport
    '        MyErr = New ErrorReport(New Exception("Cannot open File:  " + ex.Message, ex))
    '        MyErr.ShowDialog()
    '    Finally
    '        If Not oDoc Is Nothing Then oDoc.Close()
    '        oDoc = Nothing
    '        oWord = Nothing
    '    End Try

    '    '==============================================================================

    'End Sub
    'Private Sub btnOpenWord_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim strFileStatus As String
    '    Dim dr As DataRow
    '    Dim WordApp As New Word.ApplicationClass
    '    Dim fileName As Object
    '    Dim aDoc As Word.Document
    '    Dim areadOnly As Object = True
    '    Dim isVisible As Object = True
    '    Dim confirmConversions As Object = False
    '    Dim addToRecentFiles As Object = False
    '    Dim revert As Object = False
    '    ' Here is the way to handle parameters you don't care about in .NET
    '    Dim missing As Object = System.Reflection.Missing.Value
    '    Try

    '        'If ListView1.Items.Count > 0 Then
    '        '    If Me.ListView1.SelectedItems.Count <= 0 Then
    '        '        MsgBox("Please Select a Document First.")
    '        '        Exit Sub
    '        '    End If
    '        '    Dim lstItem As ListViewItem = Me.ListView1.SelectedItems(0)
    '        '    strFileStatus = lstItem.SubItems(3).Text.ToUpper
    '        '    If strFileStatus.IndexOf(UCase("Unavailable")) >= 0 Then
    '        '        'Throw New Exception("File not found")
    '        '        MsgBox("File not found")
    '        '        Exit Sub
    '        '    Else
    '        '        fileName = lstItem.SubItems(4).Text
    '        '    End If
    '        'End If

    '        '' Make word visible
    '        'WordApp.Visible = True
    '        '' Open the document that was chosen by the dialog
    '        'aDoc = WordApp.Documents.Open(fileName, confirmConversions, areadOnly, addToRecentFiles, missing, missing, revert, missing, missing, missing, missing, isVisible)

    '    Catch ex As Exception
    '        Dim MyErr As ErrorReport
    '        MyErr = New ErrorReport(New Exception("Cannot open the file in Word: " + ex.Message, ex))
    '        MyErr.ShowDialog()
    '    Finally
    '        aDoc = Nothing
    '        WordApp = Nothing
    '    End Try

    '    '==============================================================================
    'End Sub

    'Private Sub tbPageDocs_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    'End Sub

    Private Sub btnDocsPhotosCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim msgResult As MsgBoxResult
        Try
            msgResult = MsgBox("Do you want to Close Docs&Images?", MsgBoxStyle.YesNo, "DocsImages")
            If msgResult = MsgBoxResult.Yes Then
                Me.Close()
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub

    'Private Sub btnDescription_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
    '    Dim filename As Object
    '    Dim strFileStatus As String
    '    Try

    '        'If ListView1.Items.Count > 0 Then
    '        '    Dim lstItem As ListViewItem = Me.ListView1.SelectedItems(0)
    '        '    strFileStatus = lstItem.SubItems(3).Text.ToUpper
    '        '    If strFileStatus.IndexOf(UCase("Unavailable")) >= 0 Then
    '        '        Throw New Exception("File not found to update")

    '        '    Else
    '        '        filename = lstItem.SubItems(4).Text
    '        '    End If
    '        'End If
    '        '' Save the contents of the RichTextBox into the file.
    '        'If RchTextDoc.Text <> "" And Not IsNothing(filename) Then
    '        '    RchTextDoc.SaveFile(filename, RichTextBoxStreamType.PlainText)
    '        '    MsgBox("File updated Succesfully")
    '        'End If

    '    Catch ex As Exception
    '        Dim MyErr As ErrorReport
    '        MyErr = New ErrorReport(New Exception("Cannot update Description:   " + ex.Message, ex))
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub

    Private Sub tbCntrlDocsImages_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles tbCntrlDocsImages.Click
        Try
            Select Case tbCntrlDocsImages.SelectedTab.Name
                Case tbPageImages.Name
                    'images(nFacilityId)
                    If nNoOfImages <= 0 Then
                        MsgBox("No Images Found")
                    End If
                    'Case "TBPAGEDOCS"
                    'If Not ListView1.Items.Count > 0 Then
                    '    MsgBox("No Documents Found")
                    'End If
            End Select
        Catch ex As Exception
            Dim MyErr As ErrorReport
            MyErr = New ErrorReport(New Exception("Cannot open Images:  " + ex.Message, ex))
            MyErr.ShowDialog()
        End Try
    End Sub

End Class
