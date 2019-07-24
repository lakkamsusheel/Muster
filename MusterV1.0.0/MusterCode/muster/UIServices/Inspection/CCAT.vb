Public Class CCAT
    Inherits System.Windows.Forms.Form
#Region "User Defined Variables"
    Private WithEvents oInspection As MUSTER.BusinessLogic.pInspection
    Dim dtTank As DataTable
    Dim dtPipe As DataSet

    Dim dtTerm As DataSet
    Dim bolReadOnly, bolAllowClose, bolModifyingCCAT As Boolean
    Dim drow As Infragistics.Win.UltraWinGrid.UltraGridRow
    Dim grid, cell As String
    Dim rp As New Remove_Pencil
#End Region
#Region "Windows Form Designer generated code "

    Public Sub New(ByRef oInsp As MUSTER.BusinessLogic.pInspection, ByRef row As Infragistics.Win.UltraWinGrid.UltraGridRow, Optional ByVal [readOnly] As Boolean = False, Optional ByVal gridName As String = "", Optional ByVal ugCell As String = "", Optional ByVal modifyingCCAT As Boolean = False)
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Cursor.Current = Cursors.AppStarting
        oInspection = oInsp
        dtTank = New DataTable
        dtPipe = New DataSet
        dtTerm = New DataSet
        bolReadOnly = [readOnly]
        drow = row
        grid = gridName
        cell = ugCell
        chkTankMarkAll.Enabled = Not bolReadOnly
        chkPipeMarkAll.Enabled = Not bolReadOnly
        chkTermMarkAll.Enabled = Not bolReadOnly
        SetupGrid(False)
        bolAllowClose = False
        bolModifyingCCAT = modifyingCCAT
        Cursor.Current = Cursors.Default
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
    Friend WithEvents pnlCCATBottom As System.Windows.Forms.Panel
    Friend WithEvents btnCCATCancel As System.Windows.Forms.Button
    Friend WithEvents btnCCATOk As System.Windows.Forms.Button
    Friend WithEvents pnlCCATDetails As System.Windows.Forms.Panel
    Friend WithEvents lblCaption As System.Windows.Forms.Label
    Friend WithEvents btnComments As System.Windows.Forms.Button
    Friend WithEvents lblCaptionNum As System.Windows.Forms.Label
    Friend WithEvents pnlCCATHeader As System.Windows.Forms.Panel
    Friend WithEvents pnlCCATGrid As System.Windows.Forms.Panel
    Friend WithEvents pnlTerm As System.Windows.Forms.Panel
    Friend WithEvents pnlTank As System.Windows.Forms.Panel
    Friend WithEvents pnlPipe As System.Windows.Forms.Panel
    Friend WithEvents ugCCATTank As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugCCATTerm As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents ugCCATPipe As Infragistics.Win.UltraWinGrid.UltraGrid
    Friend WithEvents pnlTermTop As System.Windows.Forms.Panel
    Friend WithEvents chkTankMarkAll As System.Windows.Forms.CheckBox
    Friend WithEvents chkPipeMarkAll As System.Windows.Forms.CheckBox
    Friend WithEvents chkTermMarkAll As System.Windows.Forms.CheckBox
    Friend WithEvents Button1 As System.Windows.Forms.Button
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.pnlCCATBottom = New System.Windows.Forms.Panel
        Me.Button1 = New System.Windows.Forms.Button
        Me.btnCCATCancel = New System.Windows.Forms.Button
        Me.btnCCATOk = New System.Windows.Forms.Button
        Me.btnComments = New System.Windows.Forms.Button
        Me.pnlCCATDetails = New System.Windows.Forms.Panel
        Me.pnlCCATGrid = New System.Windows.Forms.Panel
        Me.pnlTerm = New System.Windows.Forms.Panel
        Me.ugCCATTerm = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlPipe = New System.Windows.Forms.Panel
        Me.ugCCATPipe = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlTank = New System.Windows.Forms.Panel
        Me.ugCCATTank = New Infragistics.Win.UltraWinGrid.UltraGrid
        Me.pnlTermTop = New System.Windows.Forms.Panel
        Me.chkTankMarkAll = New System.Windows.Forms.CheckBox
        Me.chkPipeMarkAll = New System.Windows.Forms.CheckBox
        Me.chkTermMarkAll = New System.Windows.Forms.CheckBox
        Me.pnlCCATHeader = New System.Windows.Forms.Panel
        Me.lblCaption = New System.Windows.Forms.Label
        Me.lblCaptionNum = New System.Windows.Forms.Label
        Me.pnlCCATBottom.SuspendLayout()
        Me.pnlCCATDetails.SuspendLayout()
        Me.pnlCCATGrid.SuspendLayout()
        Me.pnlTerm.SuspendLayout()
        CType(Me.ugCCATTerm, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlPipe.SuspendLayout()
        CType(Me.ugCCATPipe, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlTank.SuspendLayout()
        CType(Me.ugCCATTank, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlTermTop.SuspendLayout()
        Me.pnlCCATHeader.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlCCATBottom
        '
        Me.pnlCCATBottom.Controls.Add(Me.Button1)
        Me.pnlCCATBottom.Controls.Add(Me.btnCCATCancel)
        Me.pnlCCATBottom.Controls.Add(Me.btnCCATOk)
        Me.pnlCCATBottom.Controls.Add(Me.btnComments)
        Me.pnlCCATBottom.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlCCATBottom.Location = New System.Drawing.Point(0, 440)
        Me.pnlCCATBottom.Name = "pnlCCATBottom"
        Me.pnlCCATBottom.Size = New System.Drawing.Size(612, 40)
        Me.pnlCCATBottom.TabIndex = 0
        '
        'Button1
        '
        Me.Button1.Enabled = False
        Me.Button1.Location = New System.Drawing.Point(8, 8)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(104, 23)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "Add CCAT term"
        '
        'btnCCATCancel
        '
        Me.btnCCATCancel.Location = New System.Drawing.Point(328, 8)
        Me.btnCCATCancel.Name = "btnCCATCancel"
        Me.btnCCATCancel.TabIndex = 1
        Me.btnCCATCancel.Text = "Cancel"
        '
        'btnCCATOk
        '
        Me.btnCCATOk.Location = New System.Drawing.Point(248, 8)
        Me.btnCCATOk.Name = "btnCCATOk"
        Me.btnCCATOk.TabIndex = 0
        Me.btnCCATOk.Text = "Ok"
        '
        'btnComments
        '
        Me.btnComments.Location = New System.Drawing.Point(528, 8)
        Me.btnComments.Name = "btnComments"
        Me.btnComments.TabIndex = 2
        Me.btnComments.Text = "Comments"
        Me.btnComments.Visible = False
        '
        'pnlCCATDetails
        '
        Me.pnlCCATDetails.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.pnlCCATDetails.Controls.Add(Me.pnlCCATGrid)
        Me.pnlCCATDetails.Controls.Add(Me.pnlCCATHeader)
        Me.pnlCCATDetails.Controls.Add(Me.pnlCCATBottom)
        Me.pnlCCATDetails.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlCCATDetails.Location = New System.Drawing.Point(0, 0)
        Me.pnlCCATDetails.Name = "pnlCCATDetails"
        Me.pnlCCATDetails.Size = New System.Drawing.Size(616, 484)
        Me.pnlCCATDetails.TabIndex = 0
        '
        'pnlCCATGrid
        '
        Me.pnlCCATGrid.Controls.Add(Me.pnlTerm)
        Me.pnlCCATGrid.Controls.Add(Me.pnlPipe)
        Me.pnlCCATGrid.Controls.Add(Me.pnlTank)
        Me.pnlCCATGrid.Controls.Add(Me.pnlTermTop)
        Me.pnlCCATGrid.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlCCATGrid.Location = New System.Drawing.Point(0, 48)
        Me.pnlCCATGrid.Name = "pnlCCATGrid"
        Me.pnlCCATGrid.Size = New System.Drawing.Size(612, 392)
        Me.pnlCCATGrid.TabIndex = 47
        '
        'pnlTerm
        '
        Me.pnlTerm.Controls.Add(Me.ugCCATTerm)
        Me.pnlTerm.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTerm.Location = New System.Drawing.Point(0, 272)
        Me.pnlTerm.Name = "pnlTerm"
        Me.pnlTerm.Size = New System.Drawing.Size(612, 120)
        Me.pnlTerm.TabIndex = 1
        '
        'ugCCATTerm
        '
        Me.ugCCATTerm.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCCATTerm.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugCCATTerm.Location = New System.Drawing.Point(0, 0)
        Me.ugCCATTerm.Name = "ugCCATTerm"
        Me.ugCCATTerm.Size = New System.Drawing.Size(612, 120)
        Me.ugCCATTerm.TabIndex = 0
        Me.ugCCATTerm.Text = "CCAT Term"
        '
        'pnlPipe
        '
        Me.pnlPipe.Controls.Add(Me.ugCCATPipe)
        Me.pnlPipe.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlPipe.Location = New System.Drawing.Point(0, 152)
        Me.pnlPipe.Name = "pnlPipe"
        Me.pnlPipe.Size = New System.Drawing.Size(612, 120)
        Me.pnlPipe.TabIndex = 0
        '
        'ugCCATPipe
        '
        Me.ugCCATPipe.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCCATPipe.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugCCATPipe.Location = New System.Drawing.Point(0, 0)
        Me.ugCCATPipe.Name = "ugCCATPipe"
        Me.ugCCATPipe.Size = New System.Drawing.Size(612, 120)
        Me.ugCCATPipe.TabIndex = 0
        Me.ugCCATPipe.Text = "CCAT Pipe"
        '
        'pnlTank
        '
        Me.pnlTank.Controls.Add(Me.ugCCATTank)
        Me.pnlTank.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTank.Location = New System.Drawing.Point(0, 32)
        Me.pnlTank.Name = "pnlTank"
        Me.pnlTank.Size = New System.Drawing.Size(612, 120)
        Me.pnlTank.TabIndex = 0
        '
        'ugCCATTank
        '
        Me.ugCCATTank.Cursor = System.Windows.Forms.Cursors.Default
        Me.ugCCATTank.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ugCCATTank.Location = New System.Drawing.Point(0, 0)
        Me.ugCCATTank.Name = "ugCCATTank"
        Me.ugCCATTank.Size = New System.Drawing.Size(612, 120)
        Me.ugCCATTank.TabIndex = 0
        Me.ugCCATTank.Text = "CCAT Tank"
        '
        'pnlTermTop
        '
        Me.pnlTermTop.Controls.Add(Me.chkTankMarkAll)
        Me.pnlTermTop.Controls.Add(Me.chkPipeMarkAll)
        Me.pnlTermTop.Controls.Add(Me.chkTermMarkAll)
        Me.pnlTermTop.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlTermTop.Location = New System.Drawing.Point(0, 0)
        Me.pnlTermTop.Name = "pnlTermTop"
        Me.pnlTermTop.Size = New System.Drawing.Size(612, 32)
        Me.pnlTermTop.TabIndex = 0
        '
        'chkTankMarkAll
        '
        Me.chkTankMarkAll.Location = New System.Drawing.Point(24, 4)
        Me.chkTankMarkAll.Name = "chkTankMarkAll"
        Me.chkTankMarkAll.Size = New System.Drawing.Size(98, 24)
        Me.chkTankMarkAll.TabIndex = 0
        Me.chkTankMarkAll.Text = "Mark All Tanks"
        '
        'chkPipeMarkAll
        '
        Me.chkPipeMarkAll.Location = New System.Drawing.Point(128, 4)
        Me.chkPipeMarkAll.Name = "chkPipeMarkAll"
        Me.chkPipeMarkAll.Size = New System.Drawing.Size(96, 24)
        Me.chkPipeMarkAll.TabIndex = 1
        Me.chkPipeMarkAll.Text = "Mark All Pipes"
        '
        'chkTermMarkAll
        '
        Me.chkTermMarkAll.Location = New System.Drawing.Point(224, 4)
        Me.chkTermMarkAll.Name = "chkTermMarkAll"
        Me.chkTermMarkAll.Size = New System.Drawing.Size(99, 24)
        Me.chkTermMarkAll.TabIndex = 2
        Me.chkTermMarkAll.Text = "Mark All Terms"
        '
        'pnlCCATHeader
        '
        Me.pnlCCATHeader.Controls.Add(Me.lblCaption)
        Me.pnlCCATHeader.Controls.Add(Me.lblCaptionNum)
        Me.pnlCCATHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlCCATHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlCCATHeader.Name = "pnlCCATHeader"
        Me.pnlCCATHeader.Size = New System.Drawing.Size(612, 48)
        Me.pnlCCATHeader.TabIndex = 0
        '
        'lblCaption
        '
        Me.lblCaption.Location = New System.Drawing.Point(64, 8)
        Me.lblCaption.Name = "lblCaption"
        Me.lblCaption.Size = New System.Drawing.Size(536, 34)
        Me.lblCaption.TabIndex = 0
        Me.lblCaption.Text = "THERE ARE ONE OR MORE PIPING SHEAR VALVES THAT ARE NOT PROPERLY ANCHORED. THERE A" & _
        "RE ONE OR MORE PIPING SHEAR VALVES THAT ARE NOT PROPERLY ANCHORED"
        '
        'lblCaptionNum
        '
        Me.lblCaptionNum.Location = New System.Drawing.Point(8, 8)
        Me.lblCaptionNum.Name = "lblCaptionNum"
        Me.lblCaptionNum.Size = New System.Drawing.Size(56, 23)
        Me.lblCaptionNum.TabIndex = 0
        Me.lblCaptionNum.Text = "1.1.1.1.1:"
        Me.lblCaptionNum.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'CCAT
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(616, 484)
        Me.ControlBox = False
        Me.Controls.Add(Me.pnlCCATDetails)
        Me.Name = "CCAT"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "CCAT"
        Me.pnlCCATBottom.ResumeLayout(False)
        Me.pnlCCATDetails.ResumeLayout(False)
        Me.pnlCCATGrid.ResumeLayout(False)
        Me.pnlTerm.ResumeLayout(False)
        CType(Me.ugCCATTerm, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlPipe.ResumeLayout(False)
        CType(Me.ugCCATPipe, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlTank.ResumeLayout(False)
        CType(Me.ugCCATTank, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlTermTop.ResumeLayout(False)
        Me.pnlCCATHeader.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

#End Region
#Region "UI Support Routines"
    Private Function UpdateCitation() As Boolean
        Dim citation As MUSTER.Info.InspectionCitationInfo
        Dim bolUpdate As Boolean = False
        Dim strCCAT As String = String.Empty
        Dim dr As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ccat As MUSTER.Info.InspectionCCATInfo
        Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
        Try
            checkList = oInspection.InspectionInfo.ChecklistMasterCollection.Item(drow.Cells("QUESTION_ID").Value)

            ' citation exists only if citation is not equal to -1
            'If checkList.Citation <> -1 Then
            For Each citation In oInspection.InspectionInfo.CitationsCollection.Values
                If citation.QuestionID = checkList.ID Then
                    bolUpdate = True
                    Exit For
                End If
            Next

            ' tank
            Dim lastTank As Integer = -1
            For Each dr In ugCCATTank.Rows
                ccat = oInspection.InspectionInfo.CCATsCollection.Item(CType(dr.Cells("ID").Value, Int64))

                If lastTank <> ccat.TankPipeID Then
                    ccat.FirstCompartment = True
                End If

                lastTank = ccat.TankPipeID

                ccat.TankPipeResponse = IIf(dr.Cells("CCAT").Text.ToUpper = "TRUE", True, False)
                ccat.TankPipeResponseDetail = dr.Cells("Additional Details").Value

                ccat.Deleted = False
                If dr.Cells("CCAT").Value = True Then
                    strCCAT = String.Format("{0}{1}T{2} - {3}", strCCAT, IIf(strCCAT.Length > 0, ", ", String.Empty), dr.Cells("Tank#").Value, ccat.TankPipeResponseDetail)
                End If
            Next
            ' pipe
            For Each dr In ugCCATPipe.Rows
                If dr.HasChild Then
                    For Each dr2 As Infragistics.Win.UltraWinGrid.UltraGridRow In dr.ChildBands(0).Rows

                        ccat = oInspection.InspectionInfo.CCATsCollection.Item(CType(dr2.Cells("ID").Value, Int64))
                        ccat.TankPipeResponse = IIf(dr2.Cells("CCAT").Text.ToUpper = "TRUE", True, False)
                        ccat.TankPipeResponseDetail = dr2.Cells("Additional Details").Value
                        ccat.Deleted = False

                        If dr2.Cells("CCAT").Value = True Then
                            strCCAT = String.Format("{0}{1}SP{2} - {3}", strCCAT, IIf(strCCAT.Length > 0, ", ", String.Empty), dr2.Cells("Pipe#").Value, ccat.TankPipeResponseDetail)
                        End If


                    Next
                End If
                ccat = oInspection.InspectionInfo.CCATsCollection.Item(CType(dr.Cells("ID").Value, Int64))
                ccat.TankPipeResponse = IIf(dr.Cells("CCAT").Text.ToUpper = "TRUE", True, False)
                ccat.TankPipeResponseDetail = dr.Cells("Additional Details").Value
                ccat.Deleted = False

                If dr.Cells("CCAT").Value = True Then
                    strCCAT = String.Format("{0}{1}P{2} - {3}", strCCAT, IIf(strCCAT.Length > 0, ", ", String.Empty), dr.Cells("Pipe#").Value, ccat.TankPipeResponseDetail)
                End If
            Next
            ' term
            For Each dr In ugCCATTerm.Rows
                If dr.HasChild Then
                    For Each dr2 As Infragistics.Win.UltraWinGrid.UltraGridRow In dr.ChildBands(0).Rows

                        ccat = oInspection.InspectionInfo.CCATsCollection.Item(CType(dr2.Cells("ID").Value, Int64))
                        ccat.TankPipeResponse = IIf(dr2.Cells("CCAT").Text.ToUpper = "TRUE", True, False)
                        ccat.TankPipeResponseDetail = dr2.Cells("Additional Details").Value
                        ccat.Deleted = False

                        If dr2.Cells("CCAT").Value = True Then
                            strCCAT = String.Format("{0}{1}ST{2} - {3}", strCCAT, IIf(strCCAT.Length > 0, ", ", String.Empty), dr2.Cells("Term#").Value, ccat.TankPipeResponseDetail)
                        End If


                    Next
                End If
                ccat = oInspection.InspectionInfo.CCATsCollection.Item(CType(dr.Cells("ID").Value, Int64))
                ccat.TankPipeResponse = IIf(dr.Cells("CCAT").Text.ToUpper = "TRUE", True, False)
                ccat.TankPipeResponseDetail = dr.Cells("Additional Details").Value
                ccat.Deleted = False
                If dr.Cells("CCAT").Value = True Then
                    strCCAT = String.Format("{0}{1}PT{2} - {3}", strCCAT, IIf(strCCAT.Length > 0, ", ", String.Empty), dr.Cells("Term#").Value, ccat.TankPipeResponseDetail)
                End If
            Next

            If Not strCCAT.Length > 0 Then
                MsgBox("Please select atleast one CCAT")
                Return False
            End If

            ' if citation exists, update ccat else create new instance and update ccat
            If bolUpdate Then
                citation.CCAT = strCCAT
            Else
                citation = New MUSTER.Info.InspectionCitationInfo(0, _
                oInspection.ID, _
                checkList.ID, _
                oInspection.FacilityID, _
                0, _
                0, _
                checkList.Citation, _
                String.Empty, _
                False, _
                CDate("01/01/0001"), _
                Date.Now, _
                CDate("01/01/0001"), _
                False, _
                String.Empty, _
                CDate("01/01/0001"), _
                String.Empty, _
                CDate("01/01/0001"))
                oInspection.CheckListMaster.InspectionCitation.Add(citation)
                oInspection.CheckListMaster.InspectionCitation.CCAT = strCCAT
            End If
            ' update calling row
            If grid = "" Then
                drow.Cells("CCAT").Value = strCCAT
            End If
            Return True
        Catch ex As Exception
            Throw ex
        End Try
    End Function
    Private Sub SetupGrid(ByVal term As Boolean)
        Try
            Me.lblCaptionNum.Text = drow.Cells("Line#").Value.ToString
            Me.lblCaption.Text = drow.Cells("Question").Value.ToString

            If term = True Then
                oInspection.CheckListMaster.Retrieve(oInspection.InspectionInfo, oInspection.ID, oInspection.FacilityID, oInspection.OwnerID)
            End If

            If grid = "ugCP" Then
                If drow.Cells.Exists("Tank#") AndAlso Not term Then
                    dtTank = oInspection.CheckListMaster.CCATTankTable(CType(drow.Cells("QUESTION_ID").Value, Int64), bolReadOnly, CType(drow.Cells("TANK_PIPE_ID").Value, Int64))
                ElseIf drow.Cells.Exists("Pipe#") AndAlso Not term Then
                    dtPipe = oInspection.CheckListMaster.CCATPipeTable(CType(drow.Cells("QUESTION_ID").Value, Int64), bolReadOnly, CType(drow.Cells("TANK_PIPE_ID").Value, Int64))
                ElseIf drow.Cells.Exists("Term#") Then
                    dtTerm = oInspection.CheckListMaster.CCATTermTable(CType(drow.Cells("QUESTION_ID").Value, Int64), bolReadOnly, CType(drow.Cells("TANK_PIPE_ID").Value, Int64))
                Else
                    btnCCATCancel.PerformClick()
                End If
                'For i As Integer = 0 To drow.Cells.All.Length
                '    If drow.Cells.All(i).Column.key = "Tank#" Then
                '        dtTank = oInspection.CheckListMaster.CCATTankTable(CType(drow.Cells("QUESTION_ID").Value, Int64), bolReadOnly, CType(drow.Cells("TANK_PIPE_ID").Value, Int64))
                '        Exit For
                '    ElseIf drow.Cells.All(i).Column.key = "Pipe#" Then
                '        dtPipe = oInspection.CheckListMaster.CCATPipeTable(CType(drow.Cells("QUESTION_ID").Value, Int64), bolReadOnly, CType(drow.Cells("TANK_PIPE_ID").Value, Int64))
                '        Exit For
                '    ElseIf drow.Cells.All(i).Column.key = "Term#" Then
                '        dtTerm = oInspection.CheckListMaster.CCATTermTable(CType(drow.Cells("QUESTION_ID").Value, Int64), bolReadOnly, CType(drow.Cells("TANK_PIPE_ID").Value, Int64))
                '        Exit For
                '    End If
                'Next
            Else
                If Not term Then
                    dtTank = oInspection.CheckListMaster.CCATTankTable(CType(drow.Cells("QUESTION_ID").Value, Int64), bolReadOnly)
                    dtPipe = oInspection.CheckListMaster.CCATPipeTable(CType(drow.Cells("QUESTION_ID").Value, Int64), bolReadOnly)
                End If

                dtTerm = oInspection.CheckListMaster.CCATTermTable(CType(drow.Cells("QUESTION_ID").Value, Int64), bolReadOnly)

            End If


            If Not term Then
                If dtTank.Rows.Count = 0 Then
                    chkTankMarkAll.Enabled = False
                Else
                    ugCCATTank.DataSource = dtTank
                    ugCCATTank.DrawFilter = rp
                    dtTank.DefaultView.Sort = "Tank#"

                    ugCCATTank.DisplayLayout.AutoFitColumns = True
                    ugCCATTank.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False

                    For Each band As Infragistics.Win.UltraWinGrid.UltraGridBand In ugCCATTank.DisplayLayout.Bands

                        With band
                            .Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
                            .Columns("ID").Hidden = True
                            .Columns("INSPECTION_ID").Hidden = True
                            .Columns("QUESTION_ID").Hidden = True
                            .Columns("DELETED").Hidden = True
                            .Columns("Tank#").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                            .Columns("Substance").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                            .Columns("FuelType").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                            .Columns("Tank#").Width = 25
                            .Columns("CCAT").Width = 25

                            With .Columns("CompartmentID")

                                .Width = 35
                                .Header.Caption = "Comp #"
                                .DefaultCellValue = 0
                                .CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                                .NullText = "N/A"
                            End With

                            .Columns("Tank#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                        End With

                    Next band

                End If

                If dtPipe.Tables.Count = 0 OrElse dtPipe.Tables(0).Rows.Count = 0 Then
                    chkPipeMarkAll.Enabled = False
                Else
                    ugCCATPipe.DataSource = dtPipe
                    ugCCATPipe.DrawFilter = rp

                    For Each tb As DataTable In dtPipe.Tables
                        tb.DefaultView.Sort = "Pipe#"
                    Next


                    ugCCATPipe.DisplayLayout.AutoFitColumns = True
                    ugCCATPipe.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False

                    For Each band As Infragistics.Win.UltraWinGrid.UltraGridBand In ugCCATPipe.DisplayLayout.Bands

                        With band

                            If band.Index = 0 Then
                                .Override.RowAppearance.BackColor = Color.LightGoldenrodYellow
                                .Columns("PIPEID").Hidden = True
                            Else
                                .Override.RowAppearance.BackColor = Color.LightSkyBlue
                                .Columns("PARENTPIPE").Hidden = True
                            End If

                            .Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
                            .Columns("ID").Hidden = True
                            .Columns("INSPECTION_ID").Hidden = True
                            .Columns("QUESTION_ID").Hidden = True
                            .Columns("DELETED").Hidden = True
                            .Columns("Substance").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

                            .Columns("FuelType").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

                            .Columns("Pipe#").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                            .Columns("Pipe#").Width = 25
                            .Columns("CCAT").Width = 25
                            .Columns("Pipe#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending

                        End With
                    Next band


                    For Each row As Infragistics.Win.UltraWinGrid.UltraGridRow In ugCCATPipe.Rows
                        If Not row.HasChild Then
                            row.Appearance.BackColor = Color.White
                        End If
                    Next

                End If
            End If


            If dtTerm Is Nothing OrElse dtTerm.Tables.Count = 0 OrElse dtTerm.Tables(0).Rows.Count = 0 Then
                chkTermMarkAll.Enabled = False
            Else
                ugCCATTerm.DataSource = dtTerm
                ugCCATTerm.DrawFilter = rp

                ugCCATTerm.DisplayLayout.AutoFitColumns = True
                ugCCATTerm.DisplayLayout.Override.AllowDelete = Infragistics.Win.DefaultableBoolean.False

                For Each band As Infragistics.Win.UltraWinGrid.UltraGridBand In ugCCATTerm.DisplayLayout.Bands

                    With band

                        If band.Index = 0 Then
                            .Override.RowAppearance.BackColor = Color.LightGoldenrodYellow
                            .Columns("PIPEID").Hidden = True
                        Else
                            .Override.RowAppearance.BackColor = Color.LightSkyBlue
                            .Columns("PARENTPIPE").Hidden = True
                        End If



                        .Override.AllowAddNew = Infragistics.Win.UltraWinGrid.AllowAddNew.No
                        .Columns("ID").Hidden = True
                        .Columns("INSPECTION_ID").Hidden = True
                        .Columns("QUESTION_ID").Hidden = True
                        .Columns("DELETED").Hidden = True
                        .Columns("Substance").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit

                        .Columns("FuelType").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                        .Columns("Term#").CellActivation = Infragistics.Win.UltraWinGrid.Activation.NoEdit
                        .Columns("Term#").Width = 25
                        .Columns("CCAT").Width = 55
                        .Columns("Term#").SortIndicator = Infragistics.Win.UltraWinGrid.SortIndicator.Ascending
                    End With
                Next


            End If

            'If oInspection.ID > 0 Then
            '    CommentsMaintenance(, , True)
            'Else
            '    CommentsMaintenance(, , True, True)
            'End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub CommentsMaintenance(Optional ByVal sender As System.Object = Nothing, Optional ByVal e As System.EventArgs = Nothing, Optional ByVal bolSetCounts As Boolean = False, Optional ByVal resetBtnColor As Boolean = False)
        Dim SC As ShowComments
        Dim nEntityType As Integer = 0
        Dim nEntityID As Integer = 0
        Dim strEntityName As String = String.Empty
        Dim strEntityAddnInfo As String = String.Empty
        Dim oComments As MUSTER.BusinessLogic.pComments
        Dim bolEnableShowAllModules As Boolean = True
        Dim nCommentsCount As Integer = 0
        Try
            If oInspection.ID <= 0 Then
                MsgBox("Please save inspection before entering comments")
                Exit Sub
            End If

            strEntityName = "Inspection : ID #" + CStr(oInspection.ID) + " "
            strEntityAddnInfo = lblCaptionNum.Text.ToString
            nEntityID = oInspection.ID
            nEntityType = UIUtilsGen.EntityTypes.Inspection
            oComments = New MUSTER.BusinessLogic.pComments

            If Not resetBtnColor Then
                SC = New ShowComments(nEntityID, nEntityType, IIf(bolSetCounts, "", "Inspeciton"), strEntityName, oComments, Me.Text, strEntityAddnInfo, False)
                If bolSetCounts Then
                    nCommentsCount = SC.GetCounts()
                Else
                    SC.ShowDialog()
                    nCommentsCount = IIf(SC.nCommentsCount <= 0, SC.GetCounts(), SC.nCommentsCount)
                End If
            End If
            If nEntityType = UIUtilsGen.EntityTypes.Inspection Then
                If nCommentsCount > 0 Then
                    btnComments.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_HasCmts)
                Else
                    btnComments.BackColor = Drawing.Color.FromName(UIUtilsGen.CommentsButton_NoCmts)
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region
#Region "UI Control Events"
    Private Sub btnCCATOk_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCCATOk.Click
        Try
            Cursor.Current = Cursors.AppStarting
            If Not bolReadOnly Then
                If UpdateCitation() Then
                    bolAllowClose = True
                    Me.Close()
                End If
            Else
                bolAllowClose = True
                Me.Close()
            End If
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnCCATCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCCATCancel.Click
        Dim dr As Infragistics.Win.UltraWinGrid.UltraGridRow
        Dim ccat As MUSTER.Info.InspectionCCATInfo
        Dim citation As MUSTER.Info.InspectionCitationInfo
        Dim discrep As MUSTER.Info.InspectionDiscrepInfo
        Dim checkList As MUSTER.Info.InspectionChecklistMasterInfo
        Try
            Cursor.Current = Cursors.AppStarting
            If Not bolReadOnly And Not bolModifyingCCAT Then
                For Each dr In ugCCATTank.Rows
                    ccat = oInspection.InspectionInfo.CCATsCollection.Item(dr.Cells("ID").Value)
                    'ccat.Reset()
                    ccat.Deleted = True
                Next
                For Each dr In ugCCATPipe.Rows
                    ccat = oInspection.InspectionInfo.CCATsCollection.Item(dr.Cells("ID").Value)
                    'ccat.Reset()
                    ccat.Deleted = True
                Next
                For Each dr In ugCCATTerm.Rows
                    ccat = oInspection.InspectionInfo.CCATsCollection.Item(Convert.ToInt16(dr.Cells("ID").Value.ToString.Substring(1)))
                    'ccat.Reset(
                    If Not ccat Is Nothing Then
                        ccat.Deleted = True
                    End If

                Next

                checkList = oInspection.InspectionInfo.ChecklistMasterCollection.Item(drow.Cells("QUESTION_ID").Value)

                ' citation exists only if citation is not equal to -1
                If checkList.Citation <> -1 Then
                    For Each citation In oInspection.InspectionInfo.CitationsCollection.Values
                        If citation.QuestionID = CType(drow.Cells("QUESTION_ID").Value, Int64) Then
                            citation.Deleted = True
                            Exit For
                        End If
                    Next
                    'If Not (citation Is Nothing) Then
                    '    If citation.ID < 0 Then
                    '        oInspection.CheckListMaster.InspectionCitation.Remove(citation)
                    '    End If
                    'End If
                End If

                ' discrep exists only if DiscrepText is not empty
                If checkList.DiscrepText <> String.Empty Then
                    For Each discrep In oInspection.InspectionInfo.DiscrepsCollection.Values
                        If discrep.QuestionID = CType(drow.Cells("QUESTION_ID").Value, Int64) Then
                            discrep.Deleted = True
                            Exit For
                        End If
                    Next
                    'If Not (discrep Is Nothing) Then
                    '    If discrep.ID < 0 Then
                    '        oInspection.CheckListMaster.InspectionDiscrep.Remove(discrep)
                    '    End If
                    'End If
                End If

                ' reset values on row which called ccat form and assign value to Object
                If cell = "" And grid = "" Then
                    drow.Cells("RESPONSE").Value = drow.Cells("RESPONSE").OriginalValue
                    drow.Cells("N/A").Value = True
                    drow.Cells("No").Value = drow.Cells("No").OriginalValue
                    drow.Cells("Yes").Value = drow.Cells("Yes").OriginalValue
                    drow.Cells("N/A").Value = drow.Cells("N/A").OriginalValue

                    If oInspection.CheckListMaster.InspectionResponses.ID <> CType(drow.Cells("ID").Value, Int64) Then
                        oInspection.CheckListMaster.InspectionResponses.Retrieve(oInspection.InspectionInfo, drow.Cells("ID").Value)
                    End If
                    oInspection.CheckListMaster.InspectionResponses.Response = drow.Cells("RESPONSE").Value
                ElseIf grid = "ugCP" Then
                    drow.Cells("PASSFAILINCON").Value = drow.Cells("PASSFAILINCON").OriginalValue
                    drow.Cells("Pass").Value = drow.Cells("Pass").OriginalValue
                    drow.Cells("Fail").Value = IIf(drow.Cells("PASSFAILINCON").Value = 0, True, False)
                    drow.Cells("Incon").Value = IIf(drow.Cells("PASSFAILINCON").Value = 2, True, False)
                    If oInspection.CheckListMaster.InspectionCPReadings.ID <> CType(drow.Cells("ID").Value, Int64) Then
                        oInspection.CheckListMaster.InspectionCPReadings.Retrieve(oInspection.InspectionInfo, drow.Cells("ID").Value)
                    End If
                    oInspection.CheckListMaster.InspectionCPReadings.PassFailIncon = drow.Cells("PASSFAILINCON").Value
                ElseIf grid = "ugPipeLeak" Or grid = "ugTankLeak" Then
                    If oInspection.CheckListMaster.InspectionMonitorWells.ID <> CType(drow.Cells("ID").Value, Int64) Then
                        oInspection.CheckListMaster.InspectionMonitorWells.Retrieve(oInspection.InspectionInfo, drow.Cells("ID").Value)
                    End If
                    Select Case cell
                        Case ("Surface Sealed" + vbCrLf + "Yes")
                            drow.Cells("SURFACE_SEALED").Value = drow.Cells("SURFACE_SEALED").OriginalValue
                            drow.Cells("Surface Sealed" + vbCrLf + "Yes").Value = drow.Cells("Surface Sealed" + vbCrLf + "Yes").OriginalValue
                            drow.Cells("Surface Sealed" + vbCrLf + "No").Value = drow.Cells("Surface Sealed" + vbCrLf + "No").OriginalValue
                            oInspection.CheckListMaster.InspectionMonitorWells.SurfaceSealed = drow.Cells("SURFACE_SEALED").Value
                        Case ("Well Caps" + vbCrLf + "Yes")
                            drow.Cells("WELL_CAPS").Value = drow.Cells("WELL_CAPS").OriginalValue
                            drow.Cells("Well Caps" + vbCrLf + "Yes").Value = drow.Cells("Well Caps" + vbCrLf + "Yes").OriginalValue
                            drow.Cells("Well Caps" + vbCrLf + "No").Value = drow.Cells("Well Caps" + vbCrLf + "No").OriginalValue
                            oInspection.CheckListMaster.InspectionMonitorWells.WellCaps = drow.Cells("WELL_CAPS").Value
                    End Select
                End If
            End If
            bolAllowClose = True
            Cursor.Current = Cursors.Default
            Me.Close()
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub btnComments_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnComments.Click
        CommentsMaintenance(sender, e)
    End Sub
    Private Sub chkTankMarkAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkTankMarkAll.CheckedChanged
        Dim dr As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            Cursor.Current = Cursors.AppStarting
            For Each dr In ugCCATTank.Rows
                dr.Cells("CCAT").Value = chkTankMarkAll.Checked
                If Not chkTankMarkAll.Checked Then
                    dr.Cells("Additional Details").Value = String.Empty
                End If
            Next
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkPipeMarkAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkPipeMarkAll.CheckedChanged
        Dim dr As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            Cursor.Current = Cursors.AppStarting
            For Each dr In ugCCATPipe.Rows
                dr.Cells("CCAT").Value = chkPipeMarkAll.Checked
                If Not chkPipeMarkAll.Checked Then
                    dr.Cells("Additional Details").Value = String.Empty
                End If
            Next
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub chkTermMarkAll_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles chkTermMarkAll.CheckedChanged
        Dim dr As Infragistics.Win.UltraWinGrid.UltraGridRow
        Try
            Cursor.Current = Cursors.AppStarting
            For Each dr In ugCCATTerm.Rows
                dr.Cells("CCAT").Value = chkTermMarkAll.Checked
                If Not chkTermMarkAll.Checked Then
                    dr.Cells("Additional Details").Value = String.Empty
                End If
            Next
            Cursor.Current = Cursors.Default
        Catch ex As Exception
            Cursor.Current = Cursors.Default
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    'Private Sub ugCCATTank_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles ugCCATTank.BeforeRowUpdate
    '    Try
    '        CellUpdate(e.Row)
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub ugCCATPipe_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles ugCCATPipe.BeforeRowUpdate
    '    Try
    '        CellUpdate(e.Row)
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub ugCCATTerm_BeforeRowUpdate(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CancelableRowEventArgs) Handles ugCCATTerm.BeforeRowUpdate
    '    Try
    '        CellUpdate(e.Row)
    '    Catch ex As Exception
    '        Dim MyErr As New ErrorReport(ex)
    '        MyErr.ShowDialog()
    '    End Try
    'End Sub
    'Private Sub CellUpdate(ByRef row As Infragistics.Win.UltraWinGrid.UltraGridRow)
    '    Try
    '        If Not row.Cells("Additional Details").Value Is DBNull.Value Then
    '            If row.Cells("Additional Details").Value <> String.Empty And Not row.Cells("CCAT").Value Then
    '                row.Cells("CCAT").Value = True
    '            End If
    '        End If
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub
    Private Sub ugCCATTank_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugCCATTank.CellChange
        Try
            If "CCAT".Equals(e.Cell.Column.Key) Then
                If e.Cell.Text = False Then
                    e.Cell.Row.Cells("Additional Details").Value = String.Empty
                End If
            ElseIf "Additional Details".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("CCAT").Value = False Then
                    e.Cell.Row.Cells("CCAT").Value = True
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugCCATPipe_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugCCATPipe.CellChange
        Try
            If "CCAT".Equals(e.Cell.Column.Key) Then
                If e.Cell.Text = False Then
                    e.Cell.Row.Cells("Additional Details").Value = String.Empty
                End If
            ElseIf "Additional Details".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("CCAT").Value = False Then
                    e.Cell.Row.Cells("CCAT").Value = True
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Private Sub ugCCATTerm_CellChange(ByVal sender As Object, ByVal e As Infragistics.Win.UltraWinGrid.CellEventArgs) Handles ugCCATTerm.CellChange
        Try
            If "CCAT".Equals(e.Cell.Column.Key) Then
                If e.Cell.Text = False Then
                    e.Cell.Row.Cells("Additional Details").Value = String.Empty
                End If
            ElseIf "Additional Details".Equals(e.Cell.Column.Key) Then
                If e.Cell.Row.Cells("CCAT").Value = False Then
                    e.Cell.Row.Cells("CCAT").Value = True
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
    Protected Overrides Sub OnClosing(ByVal e As System.ComponentModel.CancelEventArgs)
        Try
            If Not bolAllowClose Then
                ' if you close the window without pressing ok / cancel (alt + f4)
                Dim Results As Long = MsgBox("Do you want to accept any changes made?", MsgBoxStyle.Question & MsgBoxStyle.YesNoCancel)
                If Results = MsgBoxResult.Yes Then
                    btnCCATOk.PerformClick()
                ElseIf Results = MsgBoxResult.No Then
                    btnCCATCancel.PerformClick()
                Else
                    e.Cancel = True
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()
        End Try
    End Sub
#End Region

    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click


        Dim str As String = String.Empty

        Try

            oInspection.CheckListMaster.InspectionCCAT = New BusinessLogic.pInspectionCCAT

            With oInspection.CheckListMaster.InspectionCCAT
                .QuestionID = drow.Cells("QUESTION_ID").Value
                .InspectionID = oInspection.ID
                .TankPipeID = 30431 '30827
                .TankPipeEntityID = 10
                .Termination = 1
                .TankPipeResponse = False
                .Save(UIUtilsGen.ModuleID.Inspection, MusterContainer.AppUser.UserKey, str)

                If Not str Is Nothing AndAlso str.Length > 0 Then
                    Throw New Exception(String.Format("Error adding New term. {0}", str))
                End If

            End With

            '  oInspection.CheckListMaster.RefreshCCAT()

            SetupGrid(True)
        Catch ex As Exception
            Dim MyErr As New ErrorReport(ex)
            MyErr.ShowDialog()

        End Try




    End Sub
End Class
